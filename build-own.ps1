<#
  Synthetic feed for the client's own goods (makeup tables, mirrors, poufs) built from her Prom YML export.
    powershell -File build-own.ps1 -Source "<Prom YML export link with languages=ru,uk>" -OutFile stoly-dzerkala.xml
  - no variants: every colour / size / light option is its own card
  - copied listings are merged: one card per table model and colour, identical mirrors once
  - goods of the VAYS supplier are skipped (they are in vays.xml)
  - brands and shop names removed from titles; texts follow the same structure as the VAYS cards
#>
param(
  [string]$Source = $env:OWN_FEED_URL,
  [string]$OutFile = 'stoly-dzerkala.xml',
  [int]$MinOffers = 150,
  [string]$ReportFile = ''
)
$ErrorActionPreference = 'Stop'
if (-not $Source) { throw 'No source: pass -Source or set OWN_FEED_URL' }
. (Join-Path $PSScriptRoot 'synthetic-map.ps1')
$inv = [Globalization.CultureInfo]::InvariantCulture
$IC = [Text.RegularExpressions.RegexOptions]::IgnoreCase

# ======================================================================
#  HELPERS
# ======================================================================
function Cap([string]$s) { if (-not $s) { return $s }; return $s.Substring(0,1).ToUpper() + $s.Substring(1) }
function Get-Hash32([string]$s) { $md5 = [Security.Cryptography.MD5]::Create(); $b = $md5.ComputeHash([Text.Encoding]::UTF8.GetBytes($s)); $md5.Dispose(); return [BitConverter]::ToUInt32($b, 0) }
function Select-ByHash([string]$seed, [object[]]$arr) { return $arr[[int]((Get-Hash32 $seed) % [uint32]$arr.Count)] }
function Format-Template([string]$t, [hashtable]$v) { foreach ($k in $v.Keys) { $t = $t.Replace('{' + $k + '}', [string]$v[$k]) }; return $t }
function Safe-CData([string]$s) { return $s.Replace(']]>', ']]&gt;') }
function Encode([string]$s) { return $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;') }
function Format-Num([double]$v) { if ($v -eq [math]::Floor($v)) { return ([long]$v).ToString($inv) }; return $v.ToString('0.#', $inv).Replace('.', ',') }
function Get-Plural([int]$n, [string[]]$f) { $a = $n % 10; $b = $n % 100; if ($a -eq 1 -and $b -ne 11) { return $f[0] }; if ($a -ge 2 -and $a -le 4 -and ($b -lt 12 -or $b -gt 14)) { return $f[1] }; return $f[2] }
function Norm-Code([string]$s) { return ("$s".ToLower().Replace('є', 'е') -replace '\s+', '') }
function To-Num([string]$s) { $m = [regex]::Match("$s", '\d+(?:[.,]\d+)?'); if (-not $m.Success) { return 0 }; return [double]::Parse($m.Value.Replace(',', '.'), $inv) }

# sizes from text, always returned in cm
function Get-Dims([string]$s) {
  $m = [regex]::Match("$s", '(\d{2,4}(?:[.,]\d)?)\s*[*xхX×Х]\s*(\d{2,4}(?:[.,]\d)?)(?:\s*[*xхX×Х]\s*(\d{2,4}(?:[.,]\d)?))?\s*(мм|см)?', $IC)
  if (-not $m.Success) { return $null }
  $v = @(); foreach ($g in 1..3) { if ($m.Groups[$g].Success) { $v += [double]::Parse($m.Groups[$g].Value.Replace(',', '.'), $inv) } }
  $unit = $m.Groups[4].Value.ToLower(); $max = ($v | Measure-Object -Maximum).Maximum
  if ($unit -eq 'мм' -or ($unit -ne 'см' -and $max -ge 250)) { $v = @($v | ForEach-Object { $_ / 10 }) }
  return @{ vals = $v; text = (($v | ForEach-Object { Format-Num $_ }) -join 'x') }
}

# shop names, brands and leftovers that must not reach the titles
function Clean-Name([string]$s) {
  $s = [Net.WebUtility]::HtmlDecode("$s")
  $s = $s -replace '^\s*\d+\s*(?=\p{L})', ''                                   # "1Туалетный" = copied listing
  $s = $s -replace '(?i)(?<![\p{L}])(Rika[\s-]*mebli|Rika[\s-]*market|RikaMarket|MirroRika)(?![\p{L}])', ' '
  $s = $s -replace '(?i)(?<![\p{L}])(Lion|Doros|Perfect\s+Home|Facet|Новый\s+Стиль|Новий\s+Стиль)(?![\p{L}])', ' '
  $s = $s -replace '\([A-Z]{2,4}-[\d-]+\)', ' '
  $s = $s -replace '(?<![\p{L}])ТОП(?![\p{L}])', ' '
  return ($s -replace '\s+', ' ').Trim(' ', ',', '.')
}

# ======================================================================
#  COLOURS  (longest names first; ua / ru as they appear in titles)
# ======================================================================
$Colors = @(
  @{ re = 'н[иі]мфея\s+альба'; ua = 'німфея альба'; ru = 'нимфея альба' }
  @{ re = 'орех\s+франклин(\s+табачный)?|горіх\s+франклін(\s+табачний)?'; ua = 'горіх франклін'; ru = 'орех франклин' }
  @{ re = 'самоа\s+д[ыи]мчат(ый|ий)'; ua = 'самоа димчастий'; ru = 'самоа дымчатый' }
  @{ re = 'самоа\s+димчаст(ий)'; ua = 'самоа димчастий'; ru = 'самоа дымчатый' }
  @{ re = 'дуб\s+кремона\s+торро'; ua = 'дуб кремона торро'; ru = 'дуб кремона торро' }
  @{ re = 'дуб\s+крафт\s+(белый|білий)|дуб\s+(белый|білий)\s+крафт'; ua = 'дуб крафт білий'; ru = 'дуб крафт белый' }
  @{ re = 'дуб\s+крафт\s+(серый|сірий)'; ua = 'дуб крафт сірий'; ru = 'дуб крафт серый' }
  @{ re = 'дуб\s+(сонома|санома)'; ua = 'дуб сонома'; ru = 'дуб сонома' }
  @{ re = 'дуб\s+арт[иі][зс]ан'; ua = 'дуб артизан'; ru = 'дуб артизан' }
  @{ re = 'дуб\s+л[иі]ворно'; ua = 'дуб ліворно'; ru = 'дуб ливорно' }
  @{ re = 'дуб\s+аппалач[иі]'; ua = 'дуб аппалачі'; ru = 'дуб аппалачи' }
  @{ re = 'бетон\s+style'; ua = 'бетон Style'; ru = 'бетон Style' }
  @{ re = 'бетон\s+т[её]мн(ый|ий)'; ua = 'бетон темний'; ru = 'бетон тёмный' }
  @{ re = 'венге\s+маг[иі]я'; ua = 'венге магія'; ru = 'венге магия' }
  @{ re = 'урбан\s+лайт'; ua = 'урбан лайт'; ru = 'урбан лайт' }
  @{ re = 'сер(ая|а)\s+мыш(ка|ь)|сіра\s+миш(ка)?'; ua = 'сіра мишка'; ru = 'серая мышка' }
  @{ re = '(серый|сірий)\s+т[её]мн(ый|ий)|т[её]мно-(серый|сірий)'; ua = 'темно-сірий'; ru = 'тёмно-серый' }
  @{ re = '(серый|сірий)\s+светл(ый)|светло-серый|світло-сірий'; ua = 'світло-сірий'; ru = 'светло-серый' }
  @{ re = 'светло-бежев(ый)|світло-бежев(ий)'; ua = 'світло-бежевий'; ru = 'светло-бежевый' }
  @{ re = 'светло-коричнев(ый)|світло-коричнев(ий)'; ua = 'світло-коричневий'; ru = 'светло-коричневый' }
  @{ re = 'кашем[иі]р'; ua = 'кашемір'; ru = 'кашемир' }
  @{ re = 'антрацит'; ua = 'антрацит'; ru = 'антрацит' }
  @{ re = 'графит|графіт'; ua = 'графіт'; ru = 'графит' }
  @{ re = 'бетон'; ua = 'бетон'; ru = 'бетон' }
  @{ re = 'ч[её]рн(ый|ая)|чорн(ий|а)'; ua = 'чорний'; ru = 'чёрный' }
  @{ re = 'бел(ый|ая)|біл(ий|а)'; ua = 'білий'; ru = 'белый' }
  @{ re = 'бежев(ый|ий)|беж'; ua = 'бежевий'; ru = 'бежевый' }
  @{ re = 'сер(ый)|сірий'; ua = 'сірий'; ru = 'серый' }
  @{ re = 'серебрист(ый)|сріблястий'; ua = 'сріблястий'; ru = 'серебристый' }
  @{ re = 'орех|горіх'; ua = 'горіх'; ru = 'орех' }
)
function Find-Colors([string]$s, [bool]$keepBrackets = $false) {
  # "(белый)" after a colour name is a hint, not the colour; brackets are read only when nothing else is found
  $t = $(if ($keepBrackets) { " $s " } else { " $s " -replace '\([^)]*\)', ' ' }).ToLower().Replace([char]0x0063, [char]0x0441)   # latin c in "Дуб cонома"
  $found = New-Object System.Collections.Generic.List[object]
  foreach ($c in $Colors) {
    foreach ($m in [regex]::Matches($t, "(?<![\p{L}])(?:$($c.re))(?![\p{L}])", $IC)) {
      $found.Add(@{ pos = $m.Index; c = $c })
      $t = $t.Substring(0, $m.Index) + (' ' * $m.Length) + $t.Substring($m.Index + $m.Length)
    }
  }
  if (-not $found.Count -and -not $keepBrackets -and $s -match '\(') { return @(Find-Colors $s $true) }
  return @($found | Sort-Object { $_.pos } | ForEach-Object { $_.c })
}
function Get-FirstColor([string]$s) { if (-not $s -or $s -match '\|') { return $null }; $f = @(Find-Colors $s); if ($f.Count) { return $f[0] }; return $null }
function Get-LastColor([string]$s) { $f = @(Find-Colors $s); if ($f.Count) { return $f[$f.Count - 1] }; return $null }
$ColorFields = @('Цвет', 'Цвет (оттенок) мебели', 'Цвет корпуса', 'Цвет покрытия', 'Основной цвет')

# ======================================================================
#  CATEGORIES (her Prom groups, names in Ukrainian)
# ======================================================================
$CatUa = @{
  '127281155' = 'Туалетні столики для макіяжу'; '158648863' = 'Трюмо (туалетні столики)'
  '127281154' = 'Дзеркала'; '127281157' = 'Настінні дзеркала'; '127281159' = 'Дзеркала в повний зріст з підсвіткою'
  '127281164' = 'Дзеркала в повний зріст без підсвітки'; '127281156' = 'Дзеркала для ванної кімнати'
  '127281161' = 'Круглі дзеркала'; '127281162' = 'Овальні дзеркала'; '127281163' = 'Прямокутні та квадратні дзеркала'
  '135724596' = 'Пуфи, стільці, крісла'
}
$TableCats = @('127281155', '158648863')
$BathCats = @('127281161', '127281162', '127281163')
$MirrorCats = @('127281157', '127281159', '127281164') + $BathCats
$SeatCats = @('135724596')

# goods with no usable title: fixed names (key = her article)
$Override = @{
  'PFH-091169' = @{ kind = 'table'; ua = 'Туалетний столик з шухлядою на золотих ніжках'; ru = 'Туалетный столик с ящиком на золотых ножках' }
  'NWS-122176' = @{ kind = 'barstool'; ua = 'Барний стілець на хромованій основі'; ru = 'Барный стул на хромированной основе' }
  'NWS-128388' = @{ kind = 'armchair'; ua = 'Поворотне крісло на коліщатах'; ru = 'Поворотное кресло на колёсиках' }
  'NWS-129751' = @{ kind = 'chair'; ua = 'Поворотний стілець на коліщатах'; ru = 'Поворотный стул на колёсиках' }
  'л1' = @{ kind = 'lamp'; ua = 'Світлодіодна економна лампа для гримерного дзеркала'; ru = 'Светодиодная экономная лампа для гримерного зеркала' }
}

# ======================================================================
#  TEXTS (same blocks as the VAYS cards)
# ======================================================================
$K = @{   # kind -> [ua nominative, ua genitive, ua accusative, ru nom, ru gen, ru acc, ua use, ru use]
  table    = @('туалетний столик', 'туалетного столика', 'туалетний столик', 'туалетный столик', 'туалетного столика', 'туалетный столик', 'облаштувати зручне місце для макіяжу, зачіски й щоденного догляду', 'обустроить удобное место для макияжа, причёски и ежедневного ухода')
  set      = @('туалетний столик з пуфом', 'туалетного столика з пуфом', 'туалетний столик з пуфом', 'туалетный столик с пуфом', 'туалетного столика с пуфом', 'туалетный столик с пуфом', 'облаштувати зручне місце для макіяжу, зачіски й щоденного догляду', 'обустроить удобное место для макияжа, причёски и ежедневного ухода')
  komod    = @('комод-трюмо', 'комода-трюмо', 'комод-трюмо', 'комод-трюмо', 'комода-трюмо', 'комод-трюмо', 'поєднати місце для макіяжу та зберігання в одному компактному предметі', 'совместить место для макияжа и хранение в одном компактном предмете')
  mirror   = @('дзеркало', 'дзеркала', 'дзеркало', 'зеркало', 'зеркала', 'зеркало', '', '')
  shelf    = @('полиця', 'полиці', 'полицю', 'полка', 'полки', 'полку', 'розмістити книги, декор і дрібниці, не займаючи місця на підлозі', 'разместить книги, декор и мелочи, не занимая места на полу')
  pouf     = @('пуф', 'пуфа', 'пуф', 'пуф', 'пуфа', 'пуф', 'зручно сісти біля туалетного столика, у спальні чи передпокої', 'удобно присесть у туалетного столика, в спальне или прихожей')
  chair    = @('стілець', 'стільця', 'стілець', 'стул', 'стула', 'стул', 'облаштувати зручне місце для роботи, навчання чи б''юті-процедур', 'обустроить удобное место для работы, учёбы или бьюти-процедур')
  barstool = @('барний стілець', 'барного стільця', 'барний стілець', 'барный стул', 'барного стула', 'барный стул', 'облаштувати зручне місце біля барної стійки, острова чи б''юті-столу', 'обустроить удобное место у барной стойки, острова или бьюти-стола')
  armchair = @('крісло', 'крісла', 'крісло', 'кресло', 'кресла', 'кресло', 'облаштувати зручне місце для роботи, навчання чи б''юті-процедур', 'обустроить удобное место для работы, учёбы или бьюти-процедур')
  lamp     = @('лампа', 'лампи', 'лампу', 'лампа', 'лампы', 'лампу', 'замінити лампу в гримерному дзеркалі чи столику', 'заменить лампу в гримерном зеркале или столике')
  other    = @('товар', 'товару', 'товар', 'товар', 'товара', 'товар', '', '')
}
$MirrorUse = @{
  bath     = @('зручно доглядати за собою у ванній кімнаті при рівному світлі', 'удобно ухаживать за собой в ванной комнате при ровном свете')
  makeup   = @('наносити макіяж і доглядати за собою при яскравому рівномірному світлі', 'наносить макияж и ухаживать за собой при ярком равномерном свете')
  floorlit = @('бачити весь образ на повний зріст — удома, в салоні краси чи студії', 'видеть весь образ в полный рост — дома, в салоне красоты или студии')
  frame    = @('доповнити передпокій, спальню чи вітальню й візуально розширити простір', 'дополнить прихожую, спальню или гостиную и визуально расширить пространство')
  cosmetic = @('робити макіяж і догляд за дрібними деталями', 'делать макияж и уход за мелкими деталями')
  loft     = @('додати інтер''єру характеру й візуально розширити простір', 'добавить интерьеру характера и визуально расширить пространство')
}
$T = @{
  ua = @{
    intro = @('{Head} допоможе {use}.', 'Шукаєте, як {use}? {Head} стане вдалим вибором.')
    colorPart = ' у кольорі «{c}»'
    advOf = 'Переваги {g}:'
    design = 'Сучасний лаконічний дизайн, який гармонійно доповнить {place}.'
    body = @('Продумана конструкція та універсальний дизайн забезпечать комфорт щодня.', 'Стриманий сучасний стиль легко поєднати з меблями інших кольорів і форм.')
    sloganFurniture = 'Практичні меблі для тих, хто цінує порядок, зручність та сучасний стиль!'
    sloganMirror = 'Стильне дзеркало для тих, хто цінує зручність і бездоганний вигляд!'
    sizes = 'Розміри:'; sizeW = 'Ширина {v} см'; sizeD = 'Глибина {v} см'; sizeH = 'Висота {v} см'; size2 = 'Розмір: {v} см'; diam = 'Діаметр: {v} см'; frameW = 'Ширина рамки: {v} см'
    closing = 'Можливо виготовлення в кольорах: {list}.'
    cta = 'Купити {t} з доставкою по Україні просто: оформіть замовлення, а ми подбаємо про решту.'
    delivery = 'Доставка та оплата:'
    ship = 'Відправка здійснюється транспортними компаніями Нова пошта, Міст Експрес за тарифами перевізника.'
    pay = 'Можлива відправка накладеним платежем за умови часткової передоплати.'
    trustHead = 'Чому клієнти довіряють Rika-mebli?'
    trust1 = 'Персональний супровід на всіх етапах покупки.'; trustWar = 'Гарантія {n}.'
    trust3 = 'Надійне пакування для безпечного транспортування.'; trust4 = 'Уважне ставлення до кожного замовлення.'
    trustEnd = 'Ви обираєте меблі — ми дбаємо про кожен етап вашого замовлення!'
    buy = 'купити {t}'
    places = @{ table = 'спальню чи beauty-зону'; mirror = 'ваш інтер''єр'; other = 'ваш інтер''єр' }
    a = @{
      ledTable = 'LED-підсвітка дзеркала дає рівномірне світло для макіяжу та догляду.'
      lampsTable = 'Підсвітка лампами по периметру дзеркала, як у професійній гримерці.'
      drawers = 'Місткі шухляди ({n} шт.) для косметики, прикрас і аксесуарів.'; drawer1 = 'Шухляда для косметики й дрібниць.'
      shelves = 'Додаткові полиці, щоб усе потрібне було під рукою.'
      pouf = 'У комплекті м''який пуф.'; sideL = 'Тумба з шухлядами розташована ліворуч.'; sideR = 'Тумба з шухлядами розташована праворуч.'
      socket = 'Вбудована розетка для фена чи плойки.'
      telescopic = 'Телескопічні напрямні забезпечують повне висування шухляд.'
      matLdsp = 'Корпус і стільницю виготовлено з ламінованої ДСП{t} — практичного матеріалу, який легко доглядати.'
      niche = 'Ніша під сидіння: пуф чи стілець не займає зайвого місця.'; komodMirror = 'Дзеркало зверху, щоб зручно наносити макіяж.'
      front = 'Фронтальна LED-підсвітка рівномірно освітлює обличчя — зручно для макіяжу, гоління й догляду.'
      back = 'Фонова LED-підсвітка (ефект ореолу) створює м''яке розсіяне світло й затишну атмосферу.'
      both = 'Фронтальна та фонова LED-підсвітка: яскраве світло для догляду й декоративне сяйво навколо дзеркала.'
      led = 'Вбудована LED-підсвітка дає рівномірне світло без тіней.'
      warm = 'Тепле світло — м''яке й затишне.'; neutral = 'Нейтральне світло — природна передача кольорів.'; cold = 'Холодне світло — яскраве й чітке.'
      sensor = 'Сенсорна кнопка вмикає підсвітку одним дотиком.'
      moist = 'Вологостійке виконання — підходить для ванної кімнати.'
      frameless = 'Безрамний дизайн візуально полегшує дзеркало й пасує до будь-якого інтер''єру.'
      lampsMirror = 'Підсвітка лампами дає яскраве рівномірне світло без тіней — як у професійній гримерці.'
      bulbsYes = 'Лампи входять у комплект — дзеркало готове до використання.'
      bulbsNo = 'Лампи в комплект не входять, їх можна замовити окремо.'
      shelf = 'Зручна поличка для косметики й дрібниць.'; floorShelves = 'Полиці для косметики, щіток і аксесуарів.'
      frameMat = 'Рама з {m} надає дзеркалу завершеного вигляду.'; frameMatW = 'Рама з {m} шириною {w} см надає дзеркалу завершеного вигляду.'
      podium = 'Подіум на колесах — дзеркало легко пересувати.'; wheels = 'Колеса — дзеркало легко пересувати.'
      full = 'Дзеркало в повний зріст — видно весь образ.'
      loft = 'Металева рама в стилі лофт.'; magnify = 'Одна сторона зі збільшенням — зручно для макіяжу й дрібних деталей.'
      salon = 'Підходить для дому, салону краси чи студії.'
      wall = 'Навісний монтаж — не займає місця на підлозі.'
      cube = 'Каркасна конструкція надійно тримає форму.'; soft = 'М''яке сидіння для комфорту щодня.'
      rollers = 'Коліщата для легкого пересування.'
      lamp = 'Економне світлодіодне світло.'
    }
    p = @{ ledType = 'Світлодіодна'; lampsType = 'Лампи'; none = 'Немає'; front = 'Фронтальна'; back = 'Фонова'; both = 'Фронтальна і фонова'
           warm = 'Тепле'; neutral = 'Нейтральне'; cold = 'Холодне'; yes = 'Так'; no = 'Ні'; left = 'Ліворуч'; right = 'Праворуч'
           round = 'Кругла'; oval = 'Овальна'; rect = 'Прямокутна'; square = 'Квадратна'; ldsp = 'Ламінована ДСП'; dsp = 'ДСП'; mdf = 'МДФ'; metal = 'Метал'; months = '{n} місяців'; ua = 'Україна'; new = 'Новий' }
  }
  ru = @{
    intro = @('{Head} поможет {use}.', 'Ищете, как {use}? {Head} станет удачным выбором.')
    colorPart = ' в цвете «{c}»'
    advOf = 'Преимущества {g}:'
    design = 'Современный лаконичный дизайн, который гармонично дополнит {place}.'
    body = @('Продуманная конструкция и универсальный дизайн обеспечат комфорт каждый день.', 'Сдержанный современный стиль легко сочетать с мебелью других цветов и форм.')
    sloganFurniture = 'Практичная мебель для тех, кто ценит порядок, удобство и современный стиль!'
    sloganMirror = 'Стильное зеркало для тех, кто ценит удобство и безупречный вид!'
    sizes = 'Размеры:'; sizeW = 'Ширина {v} см'; sizeD = 'Глубина {v} см'; sizeH = 'Высота {v} см'; size2 = 'Размер: {v} см'; diam = 'Диаметр: {v} см'; frameW = 'Ширина рамки: {v} см'
    closing = 'Возможно изготовление в цветах: {list}.'
    cta = 'Купить {t} с доставкой по Украине просто: оформите заказ, а мы позаботимся об остальном.'
    delivery = 'Доставка и оплата:'
    ship = 'Отправка осуществляется транспортными компаниями Новая почта, Мист Экспресс по тарифам перевозчика.'
    pay = 'Возможна отправка наложенным платежом при условии частичной предоплаты.'
    trustHead = 'Почему клиенты доверяют Rika-mebli?'
    trust1 = 'Персональное сопровождение на всех этапах покупки.'; trustWar = 'Гарантия {n}.'
    trust3 = 'Надёжная упаковка для безопасной транспортировки.'; trust4 = 'Внимательное отношение к каждому заказу.'
    trustEnd = 'Вы выбираете мебель — мы заботимся о каждом этапе вашего заказа!'
    buy = 'купить {t}'
    places = @{ table = 'спальню или beauty-зону'; mirror = 'ваш интерьер'; other = 'ваш интерьер' }
    a = @{
      ledTable = 'LED-подсветка зеркала даёт равномерный свет для макияжа и ухода.'
      lampsTable = 'Подсветка лампами по периметру зеркала, как в профессиональной гримёрке.'
      drawers = 'Вместительные ящики ({n} шт.) для косметики, украшений и аксессуаров.'; drawer1 = 'Ящик для косметики и мелочей.'
      shelves = 'Дополнительные полки, чтобы всё нужное было под рукой.'
      pouf = 'В комплекте мягкий пуф.'; sideL = 'Тумба с ящиками расположена слева.'; sideR = 'Тумба с ящиками расположена справа.'
      socket = 'Встроенная розетка для фена или плойки.'
      telescopic = 'Телескопические направляющие обеспечивают полное выдвижение ящиков.'
      matLdsp = 'Корпус и столешница изготовлены из ламинированного ДСП{t} — практичного материала, за которым легко ухаживать.'
      niche = 'Ниша под сиденье: пуф или стул не занимает лишнего места.'; komodMirror = 'Зеркало сверху, чтобы удобно наносить макияж.'
      front = 'Фронтальная LED-подсветка равномерно освещает лицо — удобно для макияжа, бритья и ухода.'
      back = 'Фоновая LED-подсветка (эффект ореола) создаёт мягкий рассеянный свет и уютную атмосферу.'
      both = 'Фронтальная и фоновая LED-подсветка: яркий свет для ухода и декоративное сияние вокруг зеркала.'
      led = 'Встроенная LED-подсветка даёт равномерный свет без теней.'
      warm = 'Тёплый свет — мягкий и уютный.'; neutral = 'Нейтральный свет — естественная передача цветов.'; cold = 'Холодный свет — яркий и чёткий.'
      sensor = 'Сенсорная кнопка включает подсветку одним касанием.'
      moist = 'Влагостойкое исполнение — подходит для ванной комнаты.'
      frameless = 'Безрамный дизайн визуально облегчает зеркало и подходит к любому интерьеру.'
      lampsMirror = 'Подсветка лампами даёт яркий равномерный свет без теней — как в профессиональной гримёрке.'
      bulbsYes = 'Лампы входят в комплект — зеркало готово к использованию.'
      bulbsNo = 'Лампы в комплект не входят, их можно заказать отдельно.'
      shelf = 'Удобная полочка для косметики и мелочей.'; floorShelves = 'Полки для косметики, щёток и аксессуаров.'
      frameMat = 'Рама из {m} придаёт зеркалу завершённый вид.'; frameMatW = 'Рама из {m} шириной {w} см придаёт зеркалу завершённый вид.'
      podium = 'Подиум на колёсах — зеркало легко передвигать.'; wheels = 'Колёса — зеркало легко передвигать.'
      full = 'Зеркало в полный рост — виден весь образ.'
      loft = 'Металлическая рама в стиле лофт.'; magnify = 'Одна сторона с увеличением — удобно для макияжа и мелких деталей.'
      salon = 'Подходит для дома, салона красоты или студии.'
      wall = 'Навесной монтаж — не занимает места на полу.'
      cube = 'Каркасная конструкция надёжно держит форму.'; soft = 'Мягкое сиденье для комфорта каждый день.'
      rollers = 'Колёсики для лёгкого перемещения.'
      lamp = 'Экономный светодиодный свет.'
    }
  }
}

# ======================================================================
#  LOAD
# ======================================================================
$xml = New-Object System.Xml.XmlDocument; $xml.XmlResolver = $null
if ($Source -match '^https?://') {
  $tmp = [IO.Path]::GetTempFileName(); $ProgressPreference = 'SilentlyContinue'
  Invoke-WebRequest -Uri $Source -OutFile $tmp -UseBasicParsing -TimeoutSec 300
  $xml.Load($tmp); Remove-Item $tmp
} else { $xml.Load((Resolve-Path $Source).Path) }
$offers = $xml.SelectNodes('/yml_catalog/shop/offers/offer')
if ($offers.Count -lt $MinOffers) { throw "Prom feed has only $($offers.Count) offers (< $MinOffers), refusing to publish a broken feed" }
if (-not $xml.SelectSingleNode('/yml_catalog/shop/offers/offer/name_ua')) { Write-Warning 'feed has no name_ua: the link should have languages=ru,uk' }
$srcCats = [ordered]@{}
foreach ($c in $xml.SelectNodes('/yml_catalog/shop/categories/category')) { $srcCats[$c.GetAttribute('id')] = @{ name = $c.InnerText.Trim(); parent = $c.GetAttribute('parentId') } }

$items = New-Object System.Collections.Generic.List[object]; $skippedVays = 0
foreach ($o in $offers) {
  $vendor = "$($o.SelectSingleNode('vendor').InnerText)".Trim(); $code = "$($o.SelectSingleNode('vendorCode').InnerText)".Trim()
  if ($vendor -eq 'VAYS' -or $code -match '^VAYS') { $skippedVays++; continue }
  $pr = @{}
  foreach ($p in $o.SelectNodes('param')) { $pn = $p.GetAttribute('name').Trim(); $pv = $p.InnerText.Trim(); if (-not $pv) { continue }; if ($pr.ContainsKey($pn)) { if ($pr[$pn] -ne $pv) { $pr[$pn] += '|' + $pv } } else { $pr[$pn] = $pv } }
  $nameRu = [Net.WebUtility]::HtmlDecode("$($o.SelectSingleNode('name').InnerText)").Trim()
  $nu = $o.SelectSingleNode('name_ua'); $nameUa = $(if ($nu) { [Net.WebUtility]::HtmlDecode($nu.InnerText).Trim() } else { $nameRu })
  $oldN = $o.SelectSingleNode('oldprice')
  $items.Add(@{
    id = $o.GetAttribute('id'); grp = $o.GetAttribute('group_id'); cat = "$($o.SelectSingleNode('categoryId').InnerText)".Trim()
    avail = ($o.GetAttribute('available') -eq 'true'); code = $code; vendor = $vendor
    price = [decimal]::Parse($o.SelectSingleNode('price').InnerText.Trim(), $inv)
    old = $(if ($oldN -and $oldN.InnerText.Trim()) { [decimal]::Parse($oldN.InnerText.Trim(), $inv) } else { [decimal]0 })
    pics = @($o.SelectNodes('picture') | ForEach-Object { $_.InnerText.Trim() } | Where-Object { $_ } | Select-Object -First 10)
    nameRu = $nameRu; nameUa = $nameUa; descRu = ("$($o.SelectSingleNode('description').InnerText)" -replace '\s+', ' '); pr = $pr
    clean = (Clean-Name $nameRu)
  })
}

# ======================================================================
#  COLOUR: the field that really differs inside a Prom variant group wins, then the colour named last in the title
# ======================================================================
$byGrp = @{}
foreach ($it in $items) { $gk = $(if ($it.grp) { "g$($it.grp)" } else { "i$($it.id)" }); if (-not $byGrp.ContainsKey($gk)) { $byGrp[$gk] = New-Object System.Collections.Generic.List[object] }; $byGrp[$gk].Add($it); $it.gk = $gk }
foreach ($gk in $byGrp.Keys) {
  $members = $byGrp[$gk]; $bestField = $null; $bestN = 1
  foreach ($f in $ColorFields) {
    $keys = @($members | ForEach-Object { $c = Get-FirstColor $_.pr[$f]; if ($c) { $c.ua } } | Select-Object -Unique)
    if ($keys.Count -gt $bestN) { $bestN = $keys.Count; $bestField = $f }
  }
  foreach ($it in $members) {
    $it.nameColor = Get-LastColor $it.clean
    $c = $null
    if ($bestField) { $c = Get-FirstColor $it.pr[$bestField] }
    if (-not $c) { $c = $it.nameColor }
    if (-not $c) { foreach ($f in $ColorFields) { $c = Get-FirstColor $it.pr[$f]; if ($c) { break } } }
    $it.color = $c
  }
}

# ======================================================================
#  WHAT IS IT
# ======================================================================
foreach ($it in $items) {
  $n = $it.clean; $pr = $it.pr; $cat = $it.cat
  $it.kind = 'other'; $it.sub = ''
  if ($Override.ContainsKey($it.code)) { $it.kind = $Override[$it.code].kind }
  elseif ($TableCats -contains $cat) {
    if ($n -match '(?i)комод-трюм|трюмо-комод') { $it.kind = 'komod' }
    elseif ($n -match '(?i)(?<![\p{L}])лампа' -and $it.price -lt 300) { $it.kind = 'lamp' }
    else { $it.kind = 'table' }
  }
  elseif ($MirrorCats -contains $cat) {
    if ($n -match '(?i)полка') { $it.kind = 'shelf' } else { $it.kind = 'mirror' }
  }
  elseif ($SeatCats -contains $cat) {
    if ($n -match '(?i)пуф') { $it.kind = 'pouf' } elseif ($n -match '(?i)барн') { $it.kind = 'barstool' } elseif ($n -match '(?i)кресл') { $it.kind = 'armchair' } else { $it.kind = 'chair' }
  }
  # --- common
  $it.dims = Get-Dims $n
  $war = To-Num $pr['Гарантийный срок']; $it.warranty = $(if ($war -ge 1) { [int]$war } else { 0 })
  $it.country = $(if ("$($pr['Страна-производитель'])" -match 'Украин' -or $it.vendor -match 'Собственное') { 'ua' } else { '' })
  # --- tables
  if ($it.kind -in 'table', 'komod') {
    $it.light = $(if ($n -match '(?i)подсвет|LED') { 'led' } elseif ($n -match '(?i)ламп') { 'lamps' } else { '' })
    $it.shelves = ($n -match '(?i)полк|полоч')
    $it.pouf = ($n -match '(?i)пуф' -or $it.code -match '(?i)пуф' -or "$($pr['Комплектация'])" -match '(?i)пуф')
    $it.side = $(if ($n -match '(?i)тумба\s+слева|\(лев') { 'L' } elseif ($n -match '(?i)тумба\s+справа|\(прав') { 'R' } elseif ("$($pr['Расположение тумбы'])" -match 'Ліворуч') { 'L' } elseif ("$($pr['Расположение тумбы'])" -match 'Праворуч') { 'R' } else { '' })
    $dr = To-Num $(if ($pr['Выдвижной ящик']) { $pr['Выдвижной ящик'] } elseif ($pr['Кількість шухляд']) { $pr['Кількість шухляд'] } else { $pr['Количество ящиков'] })
    if (-not $dr) { $m = [regex]::Match($n, '(\d+)\s*(ящик|яшик)', $IC); if ($m.Success) { $dr = [int]$m.Groups[1].Value } }
    $it.drawers = [int]$dr
    $it.socket = ($n -match '(?i)розетк')
    $it.telescopic = ("$($pr['Тип направляющих для выдвижных ящиков'])" -match '(?i)телескоп' -or $pr['Полное выдвижение ящика'] -eq 'Да')
    $it.ldsp = ("$($pr['Корпус (каркас)'])$($pr['Материал корпуса'])$($pr['Столешница'])" -match '(?i)ДСП')
    $it.ldspThick = (To-Num $pr['Толщина столешницы, см']) -eq 1.6
    $w = To-Num $(if ($pr['Длина столика']) { $pr['Длина столика'] } else { $pr['Ширина, см'] })
    $d = To-Num $(if ($pr['Глубина столика']) { $pr['Глубина столика'] } else { $pr['Глубина, см'] })
    $h = To-Num $(if ($pr['Высота столика с учетом зеркала']) { $pr['Высота столика с учетом зеркала'] } else { $pr['Высота, см'] })
    if (-not $h) { $h = To-Num $pr['Высота столика с учетом столешницы'] }
    $it.whd = $(if ($w -and $d -and $h -and $w -lt 300 -and $h -lt 300) { @($w, $d, $h) } else { $null })
    if (-not $it.dims -and $it.whd) { $it.dims = @{ vals = $it.whd; text = (($it.whd | ForEach-Object { Format-Num $_ }) -join 'x') } }
  }
  if ($it.kind -eq 'table' -and $it.pouf) { $it.kind = 'set' }
  # --- mirrors
  if ($it.kind -eq 'mirror') {
    $it.sub = $(if ($BathCats -contains $cat) { 'bath' } elseif ($cat -eq '127281159') { 'floorlit' } elseif ($cat -eq '127281164') { 'frame' }
                elseif ($n -match '(?i)косметическ') { 'cosmetic' } elseif ($n -match '(?i)в\s+раме' -and $n -notmatch '(?i)подсвет|ламп') { 'frame' } else { 'makeup' })
    if ($n -match '(?i)металлическ') { $it.sub = 'loft' }
    $fm = "$($pr['Форма'])$($pr['Форма зеркала'])"
    $it.shape = $(if ($cat -eq '127281161' -or $n -match '(?i)кругл|кругов') { 'round' } elseif ($cat -eq '127281162' -or $n -match '(?i)овальн') { 'oval' } elseif ($fm -match 'Квадрат') { 'square' } else { 'rect' })
    if ($it.shape -eq 'round') {
      $dm = "$($pr['Диаметр'])"; if (-not $dm) { $mm = [regex]::Match($n, '(?<![\p{L}])D(\d{2,3})(?!\d)'); if ($mm.Success) { $dm = $mm.Groups[1].Value } }
      if (-not $dm) { $mm = [regex]::Match($n, '(?<![\p{L}])D(\d{2,3})(?!\d)'); if ($mm.Success) { $dm = $mm.Groups[1].Value } }
      $dv = To-Num $dm; if ($dv -ge 150) { $dv = $dv / 10 }
      $it.diam = $dv
    }
    $sz = "$($pr['Размер'])"; if (-not $sz) { $sz = "$($pr['Размеры'])" }
    if ($it.shape -ne 'round' -and $sz -match '^\s*\d{2,4}\s*\*\s*\d{2,4}\s*$') { $it.dims = Get-Dims $sz }
    $it.lit = ($it.sub -in 'bath', 'makeup', 'floorlit') -and ("$($pr['Подсветка (тип ламп)'])" -ne 'Нет')
    $it.pos = $(if ($n -match '(?i)фронтальн\w*\s+и\s+(задн|фонов)') { 'both' } elseif ($n -match '(?i)фонов|задн|ореол') { 'back' } elseif ($n -match '(?i)фронтальн') { 'front' } else { '' })
    $tsrc = "$($it.code) $n $($pr['Подсветка']) $($pr['Подсветка рабочей зоны'])"
    $it.temp = $(if ($tsrc -match '(?i)тепл') { 'warm' } elseif ($tsrc -match '(?i)нейтр|нейт(?![\p{L}])') { 'neutral' } elseif ($tsrc -match '(?i)холод|(?<![\p{L}])хол(?![\p{L}])') { 'cold' } else { '' })
    $it.sensor = ($n -match '(?i)сенсорн' -or "$($pr['Выключатель'])" -match 'сенсорн' -or $pr['Сенсорная кнопка'] -eq 'Да')
    $b = "$($pr['Подсветка рабочей зоны'])"
    $it.bulbs = $(if ($b -match '(?i)с\s+лампами' -or $it.nameRu -match '(?i)с\s+лампами\s*$') { 'yes' } elseif ($b -match '(?i)без\s+ламп' -or $it.descRu -match '(?i)ламп\w*\s+в\s+комплект\s+не\s+вход') { 'no' } else { '' })
    $it.shelf = ($n -match '(?i)с\s+полк|с\s+полоч|полочками')
    $it.frameless = ($n -match '(?i)безрамоч|бескаркас|без\s+рам')
    $fmat = "$($pr['Материал рамы'])$($pr['Рама зеркала'])"
    $it.frameMat = $(if ($it.sub -eq 'loft') { 'metal' } elseif ($fmat -match 'МДФ' -or $n -match 'МДФ') { 'mdf' } elseif ($fmat -match '(?i)ЛДСП|ДСП') { 'ldsp' } else { '' })
    $mw = [regex]::Match($it.descRu, 'Ширина\s+рамки\s*[:\-–]?\s*(\d+(?:[.,]\d+)?)\s*см', $IC); $it.frameW = $(if ($mw.Success) { To-Num $mw.Groups[1].Value } else { 0 })
    $it.podium = ($n -match '(?i)подиум'); $it.wheels = ($n -match '(?i)колес|передвижн')
    $it.mount = $(if ($n -match '(?i)навесное,\s*напольное') { 'both' } elseif ($n -match '(?i)напольн|передвижн|в\s+полный\s+рост') { 'floor' } else { 'wall' })
    $it.moist = ($BathCats -contains $cat) -or ($pr['Влагоустойчивость'] -eq 'Да') -or ($pr['Влагостойкость'] -eq 'Да') -or ($n -match '(?i)влагост')
    $it.desk = ($n -match '(?i)настольн')
    $it.salon = ($n -match '(?i)салон|профессиональн|гример|бьюти|макияж')
  }
}

# group-level features: copies of one listing must look the same
function Get-Mode($vals) { $g = @($vals | Group-Object | Sort-Object Count -Descending); if ($g.Count) { return $g[0].Name }; return '' }
foreach ($gk in $byGrp.Keys) {
  $mem = $byGrp[$gk].ToArray(); if ($mem.Count -lt 2) { continue }
  $tables = @($mem | Where-Object { $_.kind -in 'table', 'set', 'komod' })
  if ($tables.Count -ge 2) {
    $pouf = (@($tables | Where-Object { $_.pouf }).Count * 2 -ge $tables.Count)
    $light = Get-Mode @($tables | Where-Object { $_.light } | ForEach-Object { $_.light })
    $side = Get-Mode @($tables | Where-Object { $_.side } | ForEach-Object { $_.side })
    $shel = (@($tables | Where-Object { $_.shelves }).Count * 2 -ge $tables.Count)
    $sock = (@($tables | Where-Object { $_.socket }).Count * 2 -gt $tables.Count)
    $drw = Get-Mode @($tables | Where-Object { $_.drawers } | ForEach-Object { $_.drawers })
    $dims = Get-Mode @($tables | Where-Object { $_.dims } | ForEach-Object { $_.dims.text })
    foreach ($tb in $tables) {
      $tb.pouf = $pouf; if ($light) { $tb.light = $light }; if ($side) { $tb.side = $side }; $tb.shelves = $shel; $tb.socket = $sock
      if ($drw) { $tb.drawers = [int]$drw }
      if ($dims -and (-not $tb.dims -or $tb.dims.text -ne $dims)) { $tb.dims = Get-Dims ($dims.Replace('x', '*') + ' см') }
      if ($tb.kind -eq 'table' -and $pouf) { $tb.kind = 'set' }; if ($tb.kind -eq 'set' -and -not $pouf) { $tb.kind = 'table' }
    }
  }
  $mir = @($mem | Where-Object { $_.kind -eq 'mirror' })
  if ($mir.Count -ge 2) {
    $pos = @($mir | Where-Object { $_.pos } | ForEach-Object { $_.pos } | Select-Object -Unique)
    if ($pos.Count -eq 1) { foreach ($x in $mir) { if (-not $x.pos) { $x.pos = $pos[0] } } }
    $tmp2 = @($mir | Where-Object { $_.temp } | ForEach-Object { $_.temp } | Select-Object -Unique)
    if ($tmp2.Count -eq 1) { foreach ($x in $mir) { if (-not $x.temp) { $x.temp = $tmp2[0] } } }
    if (@($mir | Where-Object { $_.sensor }).Count * 2 -gt $mir.Count) { foreach ($x in $mir) { $x.sensor = $true } }
    # "з лампами" is a paid option of the same mirror: the other cards of that listing are the ones without lamps
    if (@($mir | Where-Object { $_.bulbs -eq 'yes' }).Count) { foreach ($x in $mir) { if (-not $x.bulbs -and $x.sub -in 'makeup', 'floorlit') { $x.bulbs = 'no' } } }
    $fl = @($mir | Where-Object { $_.sub -eq 'floorlit' })
    if ($fl.Count -ge 2) {
      $anyShelf = [bool](@($fl | Where-Object { $_.shelf }).Count); $anyPodium = [bool](@($fl | Where-Object { $_.podium }).Count); $anyWheels = [bool](@($fl | Where-Object { $_.wheels }).Count)
      foreach ($x in $fl) { $x.shelf = $anyShelf; $x.podium = $anyPodium; $x.wheels = $anyWheels }
    }
  }
}

# ======================================================================
#  TITLES
# ======================================================================
function Get-SizeText($it, [string]$lang) {
  if ($it.kind -eq 'mirror' -and $it.shape -eq 'round' -and $it.diam) { return "Ø $(Format-Num $it.diam) см" }
  if ($it.dims) { return "$($it.dims.text) см" }
  return ''
}
function Join-Feats([string[]]$f, [string]$and) { $f = @($f | Where-Object { $_ }); if ($f.Count -le 1) { return ($f -join '') }; return (($f[0..($f.Count - 2)] -join ', ') + " $and " + $f[-1]) }
function Get-Title($it, [string]$lang, [bool]$withColor) {
  $ua = ($lang -eq 'ua'); $parts = New-Object System.Collections.Generic.List[string]
  $size = Get-SizeText $it $lang
  $colorWord = $(if ($ua) { 'колір' } else { 'цвет' })
  $col = $(if ($it.color) { $(if ($ua) { $it.color.ua } else { $it.color.ru }) } else { '' })
  $showColor = $withColor -and $col
  switch ($it.kind) {
    { $_ -in 'table', 'set' } {
      $f = @()
      if ($it.light -eq 'led') { $f += $(if ($ua) { 'LED-підсвіткою' } else { 'LED-подсветкой' }) } elseif ($it.light -eq 'lamps') { $f += $(if ($ua) { 'лампами' } else { 'лампами' }) }
      if ($it.shelves) { $f += $(if ($ua) { 'полицями' } else { 'полками' }) }
      if ($it.kind -eq 'set') { $f += $(if ($ua) { 'пуфом' } else { 'пуфом' }) }
      $head = $(if ($ua) { 'Туалетний столик з дзеркалом' } else { 'Туалетный столик с зеркалом' })
      $and = $(if ($ua) { 'і' } else { 'и' })
      if ($f.Count -eq 1) { $head += " $and " + $f[0] } elseif ($f.Count -gt 1) { $head += ', ' + (Join-Feats $f $and) }
      if ($Override.ContainsKey($it.code)) { $head = $(if ($ua) { $Override[$it.code].ua } else { $Override[$it.code].ru }) }
      $parts.Add($head)
      if ($it.drawers -and -not $Override.ContainsKey($it.code)) { $parts.Add("$($it.drawers) " + $(if ($ua) { Get-Plural $it.drawers @('шухляда', 'шухляди', 'шухляд') } else { Get-Plural $it.drawers @('ящик', 'ящика', 'ящиков') })) }
      if ($it.side -eq 'L') { $parts.Add($(if ($ua) { 'тумба ліворуч' } else { 'тумба слева' })) } elseif ($it.side -eq 'R') { $parts.Add($(if ($ua) { 'тумба праворуч' } else { 'тумба справа' })) }
    }
    'komod' { $parts.Add($(if ($ua) { 'Комод-трюмо з дзеркалом, нішею під сидіння та шухлядою' } else { 'Комод-трюмо с зеркалом, нишей под сиденье и ящиком' })) }
    'mirror' {
      $shapeUa = @{ round = 'Кругле'; oval = 'Овальне'; rect = 'Прямокутне'; square = 'Квадратне' }; $shapeRu = @{ round = 'Круглое'; oval = 'Овальное'; rect = 'Прямоугольное'; square = 'Квадратное' }
      switch ($it.sub) {
        'bath' {
          $h = $(if ($ua) { "$($shapeUa[$it.shape]) дзеркало" } else { "$($shapeRu[$it.shape]) зеркало" })
          if ($it.frameless) { $h += $(if ($ua) { ' без рами' } else { ' без рамы' }) }
          if ($it.lit) {
            $posUa = @{ front = 'фронтальною '; back = 'фоновою '; both = 'фронтальною і фоновою '; '' = '' }; $posRu = @{ front = 'фронтальной '; back = 'фоновой '; both = 'фронтальной и фоновой '; '' = '' }
            $h += $(if ($ua) { " з $($posUa[$it.pos])LED-підсвіткою" } else { " с $($posRu[$it.pos])LED-подсветкой" })
            if ($it.sensor) { $h += $(if ($ua) { ' і сенсорною кнопкою' } else { ' и сенсорной кнопкой' }) }
          }
          $h += $(if ($ua) { ' для ванної' } else { ' для ванной' })
          $parts.Add($h)
        }
        'loft' { $parts.Add($(if ($ua) { 'Кругле дзеркало в металевій рамі в стилі лофт' } else { 'Круглое зеркало в металлической раме в стиле лофт' })) }
        'cosmetic' { $parts.Add($(if ($ua) { 'Косметичне кругле дзеркало зі збільшенням' } else { 'Косметическое круглое зеркало с увеличением' })) }
        'makeup' {
          $h = $(if ($ua) { $(if ($it.desk) { 'Настільне гримерне дзеркало' } else { 'Гримерне дзеркало' }) } else { $(if ($it.desk) { 'Настольное гримерное зеркало' } else { 'Гримерное зеркало' }) })
          if ($it.frameless) { $h += $(if ($ua) { ' без рами' } else { ' без рамы' }) }
          $h += $(if ($ua) { ' з підсвіткою' } else { ' с подсветкой' })
          if ($it.shelf) { $h += $(if ($ua) { ' і поличкою' } else { ' и полочкой' }) }
          $parts.Add($h)
          if (-not $it.frameless -and $it.frameMat -eq 'mdf') { $parts.Add($(if ($ua) { 'рама МДФ' } else { 'рама МДФ' })) }
        }
        'floorlit' {
          $h = $(if ($ua) { 'Дзеркало в повний зріст з підсвіткою' } else { 'Зеркало в полный рост с подсветкой' })
          if ($it.podium) { $h += $(if ($ua) { ' на подіумі з колесами' } else { ' на подиуме с колёсами' }) } elseif ($it.wheels) { $h += $(if ($ua) { ' на колесах' } else { ' на колёсах' }) }
          if ($it.shelf) { $h += $(if ($ua) { ' і полицями' } else { ' и полками' }) }
          $parts.Add($h)
        }
        default {
          $h = $(if ($it.wheels) { $(if ($ua) { 'Підлогове дзеркало в рамі на колесах' } else { 'Напольное зеркало в раме на колёсах' }) }
                 elseif ($it.mount -eq 'both') { $(if ($ua) { 'Настінне або підлогове дзеркало в рамі' } else { 'Настенное или напольное зеркало в раме' }) }
                 elseif ($it.mount -eq 'floor') { $(if ($ua) { 'Підлогове дзеркало в рамі' } else { 'Напольное зеркало в раме' }) }
                 else { $(if ($ua) { 'Настінне дзеркало в рамі' } else { 'Настенное зеркало в раме' }) })
          if ($it.frameMat -eq 'mdf') { $h += $(if ($ua) { ' з МДФ' } else { ' из МДФ' }) }
          $parts.Add($h)
        }
      }
      # colour only where the title itself names it (black / white podium mirrors); white frames are the default
      $showColor = $withColor -and $it.nameColor -and $col
    }
    'shelf' { $parts.Add($(if ($ua) { 'Навісна полиця для дому' } else { 'Навесная полка для дома' })) }
    'pouf' {
      $parts.Add($(if ($it.clean -match '(?i)кубик|куб') { $(if ($ua) { 'Каркасний м''який пуф-куб' } else { 'Каркасный мягкий пуф-куб' }) }
                   elseif ($it.clean -match '(?i)кругл') { $(if ($ua) { 'Круглий м''який пуф' } else { 'Круглый мягкий пуф' }) }
                   else { $(if ($ua) { 'М''який пуф' } else { 'Мягкий пуф' }) }))
    }
    default {
      if ($Override.ContainsKey($it.code)) { $parts.Add($(if ($ua) { $Override[$it.code].ua } else { $Override[$it.code].ru })) }
      else { $parts.Add((Clean-Name $(if ($ua) { $it.nameUa } else { $it.nameRu }))) }
    }
  }
  if ($size) { $parts.Add($size) }
  if ($it.kind -eq 'mirror') {
    if ($it.sub -eq 'bath' -and $it.temp) { $parts.Add($(if ($ua) { @{ warm = 'тепле світло'; neutral = 'нейтральне світло'; cold = 'холодне світло' }[$it.temp] } else { @{ warm = 'тёплый свет'; neutral = 'нейтральный свет'; cold = 'холодный свет' }[$it.temp] })) }
    if ($it.sub -in 'makeup', 'floorlit' -and $it.bulbs -eq 'yes') { $parts.Add($(if ($ua) { 'лампи в комплекті' } else { 'лампы в комплекте' })) }
    if ($it.sub -in 'makeup', 'floorlit' -and $it.bulbs -eq 'no') { $parts.Add($(if ($ua) { 'без ламп' } else { 'без ламп' })) }
    if ($it.sub -eq 'frame' -and $it.frameW) { $parts.Add($(if ($ua) { "рамка $(Format-Num $it.frameW) см" } else { "рамка $(Format-Num $it.frameW) см" })) }
  }
  if ($showColor) { $parts.Add("$colorWord $col") }
  $s = ($parts -join ', ')
  if ($s.Length -gt 120) { $s = $s.Replace(' для ванної', '').Replace(' для ванной', '') }
  if ($s.Length -gt 120) { $s = $s.Replace(", $colorWord ", ', ') }
  return $s
}
foreach ($it in $items) { $it.titleUa = Get-Title $it 'ua' $true; $it.titleRu = Get-Title $it 'ru' $true; $it.coreUa = Get-Title $it 'ua' $false }

# ======================================================================
#  DUPLICATES
#    tables: one card per model and colour (her request), whatever the price
#    mirrors and the rest: same title (sensor button ignored) AND same model line (Крис, Вуди Макси...) AND
#    (same price OR same name in her shop). When most cards of two Prom listings match like that, the listing is a copy
#    and all its cards are merged even where she later changed a price. Articles are not trusted: she reuses one
#    article for different mirror models and for mirrors with and without a sensor button.
# ======================================================================
$parent = @{}; foreach ($it in $items) { $parent[$it.id] = $it.id }
function Find-Root([string]$x) { while ($parent[$x] -ne $x) { $parent[$x] = $parent[$parent[$x]]; $x = $parent[$x] }; return $x }
function Join-Items([string]$a, [string]$b) { $ra = Find-Root $a; $rb = Find-Root $b; if ($ra -ne $rb) { $parent[$rb] = $ra } }
$isTableKind = { param($x) $x.kind -in 'table', 'set', 'komod' }
foreach ($it in $items) {
  $ml = [regex]::Match($it.clean, '(?i)(Крис|Грей|Вуди|Моши|Мики|Фози)(\s+(мини|макси))?')
  $it.line = $(if ($ml.Success) { $ml.Value.ToLower() -replace '\s+', ' ' } else { '' })
  $it.cleanKey = ($it.clean.ToLower() -replace '\s+', ' ')
  if (& $isTableKind $it) { $it.keyT = 't|' + $it.titleUa.ToLower() }
  else { $s0 = $it.sensor; $it.sensor = $false; $it.keyT = 'm|' + (Get-Title $it 'ua' $true).ToLower() + '|' + $it.line; $it.sensor = $s0 }
}
$pairCount = @{}
foreach ($tg in ($items | Group-Object { $_.keyT })) {
  $l = @($tg.Group); if ($l.Count -lt 2) { continue }
  for ($i = 0; $i -lt $l.Count; $i++) { for ($j = $i + 1; $j -lt $l.Count; $j++) {
    $a = $l[$i]; $b = $l[$j]
    if ((& $isTableKind $a) -or $a.price -eq $b.price -or $a.cleanKey -eq $b.cleanKey) {
      Join-Items $a.id $b.id
      if ($a.gk -ne $b.gk) { $pk2 = @($a.gk, $b.gk) | Sort-Object; $pkey = $pk2 -join '~'; $pairCount[$pkey] = 1 + [int]$pairCount[$pkey] }
    }
  } }
}
foreach ($pkey in @($pairCount.Keys)) {
  $ga, $gb = $pkey -split '~'; $na = $byGrp[$ga].Count; $nb = $byGrp[$gb].Count
  if ($na -lt 2 -or $nb -lt 2 -or $pairCount[$pkey] * 2 -lt [math]::Min($na, $nb)) { continue }
  foreach ($a in $byGrp[$ga]) { foreach ($b in $byGrp[$gb]) { if ($a.keyT -eq $b.keyT) { Join-Items $a.id $b.id } } }
}
$clusters = @{}; foreach ($it in $items) { $r = Find-Root $it.id; if (-not $clusters.ContainsKey($r)) { $clusters[$r] = New-Object System.Collections.Generic.List[object] }; $clusters[$r].Add($it) }
$kept = New-Object System.Collections.Generic.List[object]; $merged = New-Object System.Collections.Generic.List[string]; $priceConflicts = 0
foreach ($r in $clusters.Keys) {
  $c = @($clusters[$r] | Sort-Object @{ e = { [int]$_.avail }; Descending = $true }, @{ e = { $_.pics.Count }; Descending = $true }, @{ e = { [long]$_.id }; Descending = $true })
  $best = $c[0]; $kept.Add($best)
  # copies that disagree about the sensor button: do not promise it
  if ($c.Count -gt 1 -and $best.sensor -and @($c | Where-Object { -not $_.sensor }).Count) {
    $best.sensor = $false; $best.titleUa = Get-Title $best 'ua' $true; $best.titleRu = Get-Title $best 'ru' $true
  }
  if ($c.Count -gt 1) {
    $prices = @($c | ForEach-Object { $_.price } | Select-Object -Unique); if ($prices.Count -gt 1) { $priceConflicts++ }
    $merged.Add("$($best.titleUa) | kept $($best.price) [$($best.code)] id $($best.id)$(if ($best.avail) { '' } else { ' (no stock)' }) | dropped " + ((@($c | Select-Object -Skip 1 | ForEach-Object { "$($_.price) [$($_.code)] id $($_.id)" })) -join '; '))
  }
}
# different goods that still share a title get "варіант N" (cheapest first)
$variantTitles = 0
foreach ($tg in ($kept | Group-Object { $_.titleUa.ToLower() })) {
  if ($tg.Count -lt 2) { continue }; $n = 0
  foreach ($it in @($tg.Group | Sort-Object { $_.price }, { [long]$_.id })) { $n++; $it.titleUa += ", варіант $n"; $it.titleRu += ", вариант $n"; $variantTitles++ }
}

# ======================================================================
#  DESCRIPTIONS, KEYWORDS, CHARACTERISTICS
# ======================================================================
$modelOf = @{}
foreach ($it in $kept) { $it.model = $(if ($it.kind -in 'table', 'set', 'komod', 'shelf', 'pouf') { "c:" + $it.coreUa.ToLower() } else { $it.gk }) }
$modelColors = @{}
foreach ($g in ($kept | Group-Object { $_.model })) { $modelColors[$g.Name] = @($g.Group | Where-Object { $_.color } | ForEach-Object { $_.color } | Sort-Object { $_.ua } -Unique) }

function Get-Texts($it, [string]$lang) {
  $ua = ($lang -eq 'ua'); $L = $T[$lang]; $A = $L.a; $k = $script:K[$it.kind]
  $gen = $(if ($ua) { $k[1] } else { $k[4] }); $acc = $(if ($ua) { $k[2] } else { $k[5] })
  $title = $(if ($ua) { $it.titleUa } else { $it.titleRu })
  $head = ($title -split ', ')[0]
  if ($it.kind -in 'table', 'set') { $head = $(if ($ua) { 'Туалетний столик з дзеркалом' } else { 'Туалетный столик с зеркалом' }); if ($Override.ContainsKey($it.code)) { $head = ($title -split ', ')[0] } }
  $use = $(if ($it.kind -eq 'mirror') { $u = $MirrorUse[$it.sub]; if (-not $u) { $u = $MirrorUse['frame'] }; $(if ($ua) { $u[0] } else { $u[1] }) } else { $(if ($ua) { $k[6] } else { $k[7] }) })
  $col = $(if ($it.color -and $it.kind -ne 'mirror') { $(if ($ua) { $it.color.ua } else { $it.color.ru }) } else { '' })
  $headFull = $head + $(if ($col) { Format-Template $L.colorPart @{ c = $col } } else { '' })
  $seed = "$($it.id)"
  $intro = $(if ($use) { Format-Template (Select-ByHash "$seed|i" $L.intro) @{ Head = $headFull; use = $use } } else { "$headFull." })

  $li = New-Object System.Collections.Generic.List[string]
  $place = $(if ($it.kind -in 'table', 'set', 'komod') { $L.places.table } else { $L.places.mirror })
  switch ($it.kind) {
    { $_ -in 'table', 'set' } {
      $li.Add((Format-Template $L.design @{ place = $place }))
      if ($it.light -eq 'led') { $li.Add($A.ledTable) } elseif ($it.light -eq 'lamps') { $li.Add($A.lampsTable) }
      if ($it.drawers -gt 1) { $li.Add((Format-Template $A.drawers @{ n = $it.drawers })) } elseif ($it.drawers -eq 1) { $li.Add($A.drawer1) }
      if ($it.shelves) { $li.Add($A.shelves) }
      if ($it.kind -eq 'set') { $li.Add($A.pouf) }
      if ($it.side -eq 'L') { $li.Add($A.sideL) } elseif ($it.side -eq 'R') { $li.Add($A.sideR) }
      if ($it.socket) { $li.Add($A.socket) }
      if ($it.telescopic -and $it.drawers) { $li.Add($A.telescopic) }
      if ($it.ldsp) { $li.Add((Format-Template $A.matLdsp @{ t = $(if ($it.ldspThick) { $(if ($ua) { ' товщиною 16 мм' } else { ' толщиной 16 мм' }) } else { '' }) })) }
    }
    'komod' {
      $li.Add((Format-Template $L.design @{ place = $place })); $li.Add($A.komodMirror); $li.Add($A.niche); $li.Add($A.drawer1)
      if ($it.ldsp) { $li.Add((Format-Template $A.matLdsp @{ t = $(if ($it.ldspThick) { $(if ($ua) { ' товщиною 16 мм' } else { ' толщиной 16 мм' }) } else { '' }) })) }
    }
    'mirror' {
      switch ($it.sub) {
        'bath' {
          if ($it.lit) { if ($it.pos) { $li.Add($A[$it.pos]) } else { $li.Add($A.led) } }
          if ($it.temp) { $li.Add($A[$it.temp]) }
          if ($it.sensor) { $li.Add($A.sensor) }
          if ($it.frameless) { $li.Add($A.frameless) }
          $li.Add($A.moist)
        }
        'makeup' {
          $li.Add($A.lampsMirror)
          if ($it.bulbs -eq 'yes') { $li.Add($A.bulbsYes) } elseif ($it.bulbs -eq 'no') { $li.Add($A.bulbsNo) }
          if ($it.shelf) { $li.Add($A.shelf) }
          if ($it.frameless) { $li.Add($A.frameless) } elseif ($it.frameMat -eq 'mdf') { $li.Add((Format-Template $A.frameMat @{ m = 'МДФ' })) }
          $li.Add($A.salon)
        }
        'floorlit' {
          $li.Add($A.full); $li.Add($A.lampsMirror)
          if ($it.podium) { $li.Add($A.podium) } elseif ($it.wheels) { $li.Add($A.wheels) }
          if ($it.shelf) { $li.Add($A.floorShelves) }
          $li.Add($A.salon)
        }
        'loft' { $li.Add($A.loft); $li.Add((Format-Template $L.design @{ place = $place })) }
        'cosmetic' { $li.Add($A.magnify); $li.Add((Format-Template $L.design @{ place = $place })) }
        default {
          if ($it.frameMat -eq 'mdf') { if ($it.frameW) { $li.Add((Format-Template $A.frameMatW @{ m = 'МДФ'; w = (Format-Num $it.frameW) })) } else { $li.Add((Format-Template $A.frameMat @{ m = 'МДФ' })) } }
          if ($it.mount -eq 'floor' -or $it.wheels) { $li.Add($A.full) }
          if ($it.wheels) { $li.Add($A.wheels) }
          $li.Add((Format-Template $L.design @{ place = $place }))
        }
      }
    }
    'shelf' { $li.Add($A.wall); $li.Add((Format-Template $L.design @{ place = $place })) }
    'pouf' { $li.Add($A.soft); if ($it.clean -match '(?i)каркас|куб') { $li.Add($A.cube) }; $li.Add((Format-Template $L.design @{ place = $place })) }
    { $_ -in 'chair', 'armchair', 'barstool' } { $li.Add((Format-Template $L.design @{ place = $place })); $li.Add($A.soft); if ($it.clean -match '(?i)ролик|колес') { $li.Add($A.rollers) } }
    'lamp' { $li.Add($A.lamp) }
    default { $li.Add((Format-Template $L.design @{ place = $place })) }
  }
  $sz = New-Object System.Collections.Generic.List[string]
  if ($it.whd) { $sz.Add((Format-Template $L.sizeW @{ v = (Format-Num $it.whd[0]) })); $sz.Add((Format-Template $L.sizeD @{ v = (Format-Num $it.whd[1]) })); $sz.Add((Format-Template $L.sizeH @{ v = (Format-Num $it.whd[2]) })) }
  elseif ($it.kind -eq 'mirror' -and $it.shape -eq 'round' -and $it.diam) { $sz.Add((Format-Template $L.diam @{ v = (Format-Num $it.diam) })) }
  elseif ($it.dims) { $sz.Add((Format-Template $L.size2 @{ v = $it.dims.text })) }
  if ($it.kind -eq 'mirror' -and $it.frameW) { $sz.Add((Format-Template $L.frameW @{ v = (Format-Num $it.frameW) })) }

  $sb = New-Object Text.StringBuilder
  [void]$sb.Append('<p>' + (Encode $intro) + '</p>')
  if ($li.Count) {
    [void]$sb.Append('<p><strong>' + (Encode (Format-Template $L.advOf @{ g = $gen })) + '</strong></p><ul>')
    foreach ($x in $li) { [void]$sb.Append('<li>' + (Encode $x) + '</li>') }
    [void]$sb.Append('</ul>')
  }
  if ($it.kind -ne 'lamp') {
    [void]$sb.Append('<p>' + (Encode (Select-ByHash "$seed|b" $L.body)) + '</p>')
    [void]$sb.Append('<p>' + (Encode $(if ($it.kind -eq 'mirror') { $L.sloganMirror } else { $L.sloganFurniture })) + '</p>')
  }
  if ($sz.Count) {
    [void]$sb.Append('<p><strong>' + (Encode $L.sizes) + '</strong></p><ul>')
    foreach ($x in $sz) { [void]$sb.Append('<li>' + (Encode $x) + '</li>') }
    [void]$sb.Append('</ul>')
  }
  $cols = @($modelColors[$it.model])
  if ($cols.Count -ge 2) { [void]$sb.Append('<p>' + (Encode (Format-Template $L.closing @{ list = (@($cols | ForEach-Object { $(if ($ua) { $_.ua } else { $_.ru }) }) -join ', ') })) + '</p>') }
  [void]$sb.Append('<p>' + (Encode (Format-Template $L.cta @{ t = $acc })) + '</p>')
  [void]$sb.Append('<p><strong>' + (Encode $L.delivery) + '</strong></p>')
  [void]$sb.Append('<p>' + (Encode $L.ship) + '</p><p>' + (Encode $L.pay) + '</p>')
  [void]$sb.Append('<p><strong>' + (Encode $L.trustHead) + '</strong></p><ul>')
  $trust = @($L.trust1)
  if ($it.warranty) { $trust += (Format-Template $L.trustWar @{ n = "$($it.warranty) " + $(if ($ua) { Get-Plural $it.warranty @('місяць', 'місяці', 'місяців') } else { Get-Plural $it.warranty @('месяц', 'месяца', 'месяцев') }) }) }
  $trust += $L.trust3; $trust += $L.trust4
  foreach ($x in $trust) { [void]$sb.Append('<li>' + (Encode $x) + '</li>') }
  [void]$sb.Append('</ul><p>' + (Encode $L.trustEnd) + '</p>')

  # keywords: what people type in search
  $kw = New-Object System.Collections.Generic.List[string]
  $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
  $add = { param($s) $s = ($s -replace '\s+', ' ').Trim(); if ($s -and $seen.Add($s)) { $kw.Add($s) } }
  $base = @()
  switch ($it.kind) {
    { $_ -in 'table', 'set' } {
      $base = $(if ($ua) { @('туалетний столик', 'туалетний столик з дзеркалом', 'столик для макіяжу', 'гримерний столик', 'трюмо', 'косметичний столик') } else { @('туалетный столик', 'туалетный столик с зеркалом', 'столик для макияжа', 'гримерный столик', 'трюмо', 'косметический столик') })
      if ($it.light) { $base += $(if ($ua) { 'туалетний столик з підсвіткою' } else { 'туалетный столик с подсветкой' }) }
      if ($it.kind -eq 'set') { $base += $(if ($ua) { 'туалетний столик з пуфом' } else { 'туалетный столик с пуфом' }) }
    }
    'komod' { $base = $(if ($ua) { @('комод-трюмо', 'трюмо з дзеркалом', 'туалетний столик', 'комод з дзеркалом', 'столик для макіяжу') } else { @('комод-трюмо', 'трюмо с зеркалом', 'туалетный столик', 'комод с зеркалом', 'столик для макияжа') }) }
    'mirror' {
      switch ($it.sub) {
        'bath' { $base = $(if ($ua) { @('дзеркало у ванну', 'дзеркало з підсвіткою', 'дзеркало з LED-підсвіткою', 'дзеркало для ванної кімнати', "$(@{ round = 'кругле'; oval = 'овальне'; rect = 'прямокутне'; square = 'квадратне' }[$it.shape]) дзеркало з підсвіткою") } else { @('зеркало в ванную', 'зеркало с подсветкой', 'зеркало с LED-подсветкой', 'зеркало для ванной комнаты', "$(@{ round = 'круглое'; oval = 'овальное'; rect = 'прямоугольное'; square = 'квадратное' }[$it.shape]) зеркало с подсветкой") }) }
        'makeup' { $base = $(if ($ua) { @('гримерне дзеркало', 'дзеркало з лампами', 'дзеркало з підсвіткою', 'дзеркало для макіяжу', 'б''юті-дзеркало') } else { @('гримерное зеркало', 'зеркало с лампами', 'зеркало с подсветкой', 'зеркало для макияжа', 'бьюти-зеркало') }) }
        'floorlit' { $base = $(if ($ua) { @('дзеркало в повний зріст', 'дзеркало з підсвіткою', 'підлогове дзеркало', 'дзеркало для салону', 'дзеркало на колесах') } else { @('зеркало в полный рост', 'зеркало с подсветкой', 'напольное зеркало', 'зеркало для салона', 'зеркало на колёсах') }) }
        'cosmetic' { $base = $(if ($ua) { @('косметичне дзеркало', 'дзеркало зі збільшенням', 'дзеркало для макіяжу') } else { @('косметическое зеркало', 'зеркало с увеличением', 'зеркало для макияжа') }) }
        default {
          $floor = ($it.mount -ne 'wall' -or $it.wheels); $tall = ($it.dims -and ($it.dims.vals | Measure-Object -Maximum).Maximum -ge 120)
          $base = $(if ($ua) { @('дзеркало в рамі', $(if ($floor) { 'підлогове дзеркало' } else { 'настінне дзеркало' }), 'дзеркало в передпокій', 'дзеркало для дому') } else { @('зеркало в раме', $(if ($floor) { 'напольное зеркало' } else { 'настенное зеркало' }), 'зеркало в прихожую', 'зеркало для дома') })
          if ($tall) { $base += $(if ($ua) { 'дзеркало в повний зріст' } else { 'зеркало в полный рост' }) }
        }
      }
    }
    'shelf' { $base = $(if ($ua) { @('навісна полиця', 'полиця на стіну', 'полиця для книг', 'полиця для декору') } else { @('навесная полка', 'полка на стену', 'полка для книг', 'полка для декора') }) }
    'pouf' { $base = $(if ($ua) { @('пуф', 'м''який пуф', 'пуф для туалетного столика', 'пуф у спальню', 'пуф у передпокій') } else { @('пуф', 'мягкий пуф', 'пуф для туалетного столика', 'пуф в спальню', 'пуф в прихожую') }) }
    default { $base = @($(if ($ua) { $k[0] } else { $k[3] })) }
  }
  foreach ($s in $base) { & $add $s }
  if ($col) { & $add "$($(if ($ua) { $k[0] } else { $k[3] })) $col" }
  $szT = Get-SizeText $it $lang; if ($szT -and $it.kind -ne 'lamp') { & $add "$($(if ($ua) { $k[0] } else { $k[3] })) $($szT -replace '^Ø ', '')" }
  & $add (Format-Template $L.buy @{ t = $acc })
  return @{ name = $title; desc = $sb.ToString(); kw = ($kw -join ', ') }
}

function Get-Params($it) {
  $p = [ordered]@{}; $PV = $T.ua.p
  $typeUa = @{ table = 'Туалетний столик'; set = 'Туалетний столик з пуфом'; komod = 'Комод-трюмо'; mirror = 'Дзеркало'; shelf = 'Полиця навісна'; pouf = 'Пуф'; chair = 'Стілець'; barstool = 'Стілець барний'; armchair = 'Крісло'; lamp = 'Лампа'; other = '' }
  if ($typeUa[$it.kind]) { $p['Тип товару'] = $typeUa[$it.kind] }
  if ($it.color -and -not ($it.kind -eq 'mirror' -and -not $it.nameColor -and $it.color.ua -in 'білий', 'сріблястий')) { $p['Колір'] = Cap $it.color.ua }
  if ($it.whd) { $p['Ширина'] = Format-Num ($it.whd[0] * 10); $p['Глибина'] = Format-Num ($it.whd[1] * 10); $p['Висота'] = Format-Num ($it.whd[2] * 10) }
  elseif ($it.kind -eq 'mirror' -and $it.shape -eq 'round' -and $it.diam) { $p['Діаметр'] = Format-Num ($it.diam * 10) }
  elseif ($it.dims) { $p['Розмір'] = "$($it.dims.text) см" }
  if ($it.kind -eq 'mirror') {
    $p['Форма'] = $PV[$it.shape]
    if ($it.sub -in 'bath', 'makeup', 'floorlit') { $p['Підсвітка'] = $(if (-not $it.lit) { $PV.none } elseif ($it.sub -eq 'bath') { $PV.ledType } else { $PV.lampsType }) }
    if ($it.sub -eq 'bath' -and $it.pos) { $p['Розташування підсвітки'] = $PV[$it.pos] }
    if ($it.sub -eq 'bath' -and $it.temp) { $p['Колір світла'] = $PV[$it.temp] }
    if ($it.sensor -and $it.lit) { $p['Сенсорна кнопка'] = $PV.yes }
    if ($it.bulbs) { $p['Лампи в комплекті'] = $(if ($it.bulbs -eq 'yes') { $PV.yes } else { $PV.no }) }
    if ($it.moist -and $it.sub -eq 'bath') { $p['Вологостійкість'] = $PV.yes }
    if ($it.frameless) { $p['Рама'] = 'Без рами' } elseif ($it.frameMat) { $p['Матеріал рами'] = $PV[$it.frameMat] }
    if ($it.frameW) { $p['Ширина рамки'] = Format-Num ($it.frameW * 10) }
  }
  if ($it.kind -in 'table', 'set', 'komod') {
    if ($it.drawers -and $it.kind -ne 'komod') { $p['Кількість висувних ящиків'] = "$($it.drawers)" }
    if ($it.drawers -or $it.kind -eq 'komod') { $p['Наявність шухляди'] = $PV.yes }
    if ($it.light) { $p['Підсвітка'] = $(if ($it.light -eq 'led') { $PV.ledType } else { $PV.lampsType }) }
    if ($it.shelves) { $p['Полиці'] = $PV.yes }
    if ($it.kind -eq 'set') { $p['Пуф у комплекті'] = $PV.yes }
    if ($it.side) { $p['Розташування тумби'] = $(if ($it.side -eq 'L') { $PV.left } else { $PV.right }) }
    if ($it.ldsp) { $p['Корпус (каркас)'] = $PV.ldsp }
    if ($it.telescopic -and $it.drawers) { $p['Тип напрямних'] = 'Телескопічні' }
  }
  if ($it.warranty) { $p['Гарантійний термін'] = "$($it.warranty) " + (Get-Plural $it.warranty @('місяць', 'місяці', 'місяців')) }
  if ($it.country -eq 'ua') { $p['Країна виробник'] = $PV.ua }
  $p['Стан'] = $PV.new
  return $p
}
$ParamUnit = @{ 'Ширина' = 'мм'; 'Глибина' = 'мм'; 'Висота' = 'мм'; 'Діаметр' = 'мм'; 'Ширина рамки' = 'мм'; 'Кількість висувних ящиків' = 'шт.' }

# ======================================================================
#  ORDER (no variants on Synthetic: mix models, goods in stock first) AND WRITE
# ======================================================================
$noPics = @($kept | Where-Object { -not $_.pics.Count })   # a card without photos is useless on Synthetic
$queues = New-Object System.Collections.Generic.List[object]; $qIndex = @{}
foreach ($it in @($kept | Where-Object { $_.pics.Count } | Sort-Object @{ e = { [int]$_.avail }; Descending = $true }, { $_.cat }, { $_.model }, { [double]$_.price })) {
  if (-not $qIndex.ContainsKey($it.model)) { $qIndex[$it.model] = New-Object System.Collections.Generic.List[object]; $queues.Add($qIndex[$it.model]) }
  $qIndex[$it.model].Add($it)
}
$ordered = New-Object System.Collections.Generic.List[object]; $maxLen = ($queues | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
for ($r = 0; $r -lt $maxLen; $r++) { foreach ($q in $queues) { if ($r -lt $q.Count) { $ordered.Add($q[$r]) } } }

# articles: brand prefixes removed; empty or repeated ones get RIKA-<id>
$codeCount = @{}; foreach ($it in $ordered) { $c = ($it.code -replace '(?i)^(LION|DRS|NWS|PFH)-', '' -replace '(?i)/(LION|DRS|NWS|PFH)-', '/').Trim(); $it.vc = $c; if ($c) { $codeCount[$c.ToLower()] = 1 + [int]$codeCount[$c.ToLower()] } }
foreach ($it in $ordered) { if (-not $it.vc -or $codeCount[$it.vc.ToLower()] -gt 1) { $it.vc = "RIKA-$($it.id)" } }

$usedCats = @{}; foreach ($it in $ordered) { $usedCats[$it.cat] = $true }
foreach ($cid in @($usedCats.Keys)) { $p = $(if ($srcCats.Contains($cid)) { $srcCats[$cid].parent } else { '' }); while ($p) { $usedCats[$p] = $true; $p = $(if ($srcCats.Contains($p)) { $srcCats[$p].parent } else { '' }) } }

$outPath = $(if ([IO.Path]::IsPathRooted($OutFile)) { $OutFile } else { [IO.Path]::GetFullPath((Join-Path (Get-Location) $OutFile)) })
$settings = New-Object System.Xml.XmlWriterSettings; $settings.Indent = $true; $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
$w = [System.Xml.XmlWriter]::Create($outPath, $settings)
$w.WriteStartDocument()
$w.WriteStartElement('yml_catalog'); $w.WriteAttributeString('date', (Get-Date).ToUniversalTime().AddHours(2).ToString('yyyy-MM-dd HH:mm', $inv))
Write-SyNamespace $w
$w.WriteStartElement('shop')
$w.WriteElementString('name', 'Rika Mebli'); $w.WriteElementString('company', 'Rika Mebli'); $w.WriteElementString('url', 'https://rika-mebli.prom.ua/')
$w.WriteStartElement('currencies'); $w.WriteStartElement('currency'); $w.WriteAttributeString('id', 'UAH'); $w.WriteAttributeString('rate', '1'); $w.WriteEndElement(); $w.WriteEndElement()
# Synthetic categories (with sy_code); goods of an unknown kind keep their Prom category
$syCodes = New-Object System.Collections.Generic.List[string]; $syFallback = [ordered]@{}; $syUnmapped = New-Object System.Collections.Generic.List[string]
foreach ($it in $ordered) {
  $it.sy = Get-SyOwn $it
  if ($it.sy) { $syCodes.Add($it.sy[0]) }
  else { $syUnmapped.Add("$($it.kind): $($it.code)"); $syFallback[$it.cat] = $(if ($CatUa.ContainsKey($it.cat)) { $CatUa[$it.cat] } else { $srcCats[$it.cat].name }) }
}
Write-SyCategories $w $syCodes.ToArray() $syFallback
$w.WriteStartElement('offers')
$maxName = 0; $other = 0
foreach ($it in $ordered) {
  if ($it.kind -eq 'other') { $other++ }
  $ua = Get-Texts $it 'ua'; $ru = Get-Texts $it 'ru'
  $maxName = [math]::Max($maxName, [math]::Max($ua.name.Length, $ru.name.Length))
  $w.WriteStartElement('offer'); $w.WriteAttributeString('id', $it.id); $w.WriteAttributeString('available', $(if ($it.avail) { 'true' } else { 'false' }))
  $w.WriteElementString('name', $ua.name); $w.WriteElementString('name_ru', $ru.name)
  $w.WriteElementString('vendorCode', $it.vc)
  $w.WriteElementString('price', $it.price.ToString('0.##', $inv))
  if ($it.old -gt $it.price) { $w.WriteElementString('oldprice', $it.old.ToString('0.##', $inv)) }
  $w.WriteElementString('currencyId', 'UAH')
  $w.WriteElementString('categoryId', $(if ($it.sy) { $it.sy[0] } else { $it.cat }))
  Write-SyOffer $w $it.sy $it.warranty $it.old $it.price $(if ($it.avail) { $null } else { 0 })
  foreach ($pic in $it.pics) { $w.WriteElementString('picture', $pic) }
  $w.WriteStartElement('description'); $w.WriteCData((Safe-CData $ua.desc)); $w.WriteEndElement()
  $w.WriteStartElement('description_ru'); $w.WriteCData((Safe-CData $ru.desc)); $w.WriteEndElement()
  $w.WriteElementString('keywords', $ua.kw); $w.WriteElementString('keywords_ru', $ru.kw)
  $params = Get-Params $it
  foreach ($pk in $params.Keys) { $w.WriteStartElement('param'); $w.WriteAttributeString('name', $pk); if ($ParamUnit.ContainsKey($pk)) { $w.WriteAttributeString('unit', $ParamUnit[$pk]) }; $w.WriteString($params[$pk]); $w.WriteEndElement() }
  $w.WriteEndElement()
}
$w.WriteEndElement(); $w.WriteEndElement(); $w.WriteEndElement()
$w.WriteEndDocument(); $w.Flush(); $w.Close()

"written: $outPath"
"source offers: $($offers.Count); VAYS skipped: $skippedVays; own goods: $($items.Count)"
"cards: $($ordered.Count) (in stock: $(@($ordered | Where-Object { $_.avail }).Count)); copies merged: $($items.Count - $kept.Count) in $($merged.Count) cases ($priceConflicts with different prices); titles with 'варіант N': $variantTitles; max title length: $maxName"
"by kind: " + ((@($ordered | Group-Object { $_.kind } | Sort-Object Count -Descending | ForEach-Object { "$($_.Name) $($_.Count)" })) -join ', ')
if ($other) { Write-Warning "$other goods of an unknown type got their own Prom title" }
"Synthetic categories: " + ((@($ordered | Where-Object { $_.sy } | Group-Object { $SyCatName[$_.sy[0]] } | Sort-Object Count -Descending | ForEach-Object { "$($_.Name) $($_.Count)" })) -join ', ')
if ($syUnmapped.Count) { "no Synthetic category (left in their Prom category): " + ($syUnmapped -join '; ') }
if ($noPics.Count) { "skipped, no photos on Prom: " + ((@($noPics | ForEach-Object { "$($_.titleUa) [$($_.code)]" })) -join '; ') }
if ($ReportFile) { $merged | Sort-Object | Set-Content -Path $ReportFile -Encoding utf8 }
