<#
  Synthetic catalogue codes shared by build-prom.ps1 (VAYS) and build-own.ps1 (her tables / mirrors).
  Requirements: https://synthetic.ua/pages/seller/xmlreq
  Codes come from her seller cabinet: my.synthetic.ua > Імпорт та синхронізація > Правила синхронізації.

  On 2026-11-27 most furniture categories had only placeholder product types (quality AVERAGE, the cabinet
  says "Типи товарів для цієї категорії не налаштовані"). Setup requests were sent from her cabinet
  (tickets RA00083100-RA00083110 + shoe cabinets). When Synthetic sets up proper types, put the new PTD codes here.
#>
$SyNs = 'http://base.synthetic.ua/ns/1.0'
$SyVendor = 'Інше'   # the brand list on Synthetic has neither the shop nor the factories; she picks "Інше"

# category code -> name on Synthetic
$SyCatName = @{
  'C762766' = 'Ліжка та каркаси'
  'C326696' = 'Комоди'
  'C526222' = 'Тумби та підставки для взуття'
  'C352215' = 'Дзеркала'
  'C162424' = 'Дзеркала для ванної кімнати'
  'C972614' = 'Косметичні дзеркала'
  'C635050' = 'Туалетні столики'
  'C653316' = 'Приліжкові тумби'
  'C487583' = 'Шафа'
  'C967313' = 'Полиці для будинку'
  'C754109' = 'Підлогові і підвісні передпокої'
  'C067557' = "Комп'ютерні та письмові столи"
  'C964447' = 'Тумби для ванної'
  'C342205' = 'Стільці і табурети'
  'C686196' = 'Барні стільці'
}

# kind of goods -> Synthetic category code, product type code (sy_type)
$Sy = @{
  bed        = @('C762766', 'PTD-00115')   # Ліжко
  komod      = @('C326696', 'PTD-01954')   # Комоди
  tv         = @('C326696', 'PTD-01954')   # no TV-stand category on Synthetic: "Комоди та тумби"
  shoe       = @('C526222', 'PTD-03653')   # Тумби під взуття
  mirror     = @('C352215', 'PTD-03664')   # Інтер'єрні дзеркала
  bathmirror = @('C162424', 'PTD-4262')    # Дзеркала у ванну (set up properly)
  cosmetic   = @('C972614', 'PTD-00390')   # Дзеркала косметичні (set up properly)
  dressing   = @('C635050', 'PTD-4496')    # Туалетні столики
  bedside    = @('C653316', 'PTD-01016')   # Приліжкові тумби
  wardrobe   = @('C487583', 'PTD-00116')   # Шафа
  shelf      = @('C967313', 'PTD-01758')   # Навісні полиці (set up properly; the only type in "Полиці, стелажі")
  hallway    = @('C754109', 'PTD-03652')   # Передпокій
  pc         = @('C067557', 'PTD-03650')   # Комп'ютерні столи
  desk       = @('C067557', 'PTD-03649')   # Письмові столи
  bath       = @('C964447', 'PTD-03666')   # Тумби у ванну
  stool      = @('C342205', 'PTD-01208')   # Табурети (poufs)
  chair      = @('C342205', 'PTD-03442')   # Стільці для кухні ("М'які крісла" has no types at all)
  barstool   = @('C686196', 'PTD-01759')   # Барні стільці
}

# VAYS supplier "Тип товара" -> kind
$SyVaysKind = @{
  'Кровать двуспальная' = 'bed'; 'Кровать односпальная' = 'bed'
  'Комод' = 'komod'; 'Тумба под телевизор' = 'tv'
  'Тумба для обуви' = 'shoe'; 'Полка для обуви' = 'shoe'
  'Зеркало' = 'mirror'; 'Трюмо (туалетный столик)' = 'dressing'; 'Тумба прикроватная' = 'bedside'
  'Антресоль' = 'wardrobe'; 'Пенал' = 'wardrobe'; 'Пенал для книг' = 'wardrobe'; 'Шкаф (комплект)' = 'wardrobe'
  'Шкаф распашной' = 'wardrobe'; 'Шкаф для книг' = 'wardrobe'; 'Комплект мебели в гостиную' = 'wardrobe'
  'Полка навесная' = 'shelf'; 'Полка' = 'shelf'; 'Стеллаж' = 'shelf'
  'Вешалка для прихожей' = 'hallway'; 'Вешалка' = 'hallway'; 'Комплект мебели в прихожую' = 'hallway'
  'Стол компьютерный' = 'pc'; 'Стол письменный' = 'desk'
  'Тумба под стиральную машину' = 'bath'
}
$SyVaysBathCat = '571'   # "Шафи для ванної кімнати": its pencil cases go to bathroom furniture

function Get-SyVays([string]$type, [string]$cat) {
  if ($cat -eq $SyVaysBathCat) { return $Sy.bath }
  if ($SyVaysKind.ContainsKey("$type")) { return $Sy[$SyVaysKind["$type"]] }
  return $null
}

function Get-SyOwn($it) {
  switch ($it.kind) {
    { $_ -in 'table', 'set', 'komod' } { return $Sy.dressing }
    'mirror' { if ($it.sub -eq 'bath') { return $Sy.bathmirror } elseif ($it.sub -eq 'cosmetic') { return $Sy.cosmetic } else { return $Sy.mirror } }
    'shelf' { return $Sy.shelf }
    'pouf' { return $Sy.stool }
    { $_ -in 'chair', 'armchair' } { return $Sy.chair }
    'barstool' { return $Sy.barstool }
  }
  return $null
}

# <yml_catalog xmlns:sy="...">
function Write-SyNamespace($w) { $w.WriteAttributeString('xmlns', 'sy', 'http://www.w3.org/2000/xmlns/', $SyNs) }

# categories = the Synthetic categories in use (id and sy_code are the Synthetic code); goods without a mapping
# keep their own category, without sy_code
function Write-SyCategories($w, [string[]]$codes, $fallback) {
  $w.WriteStartElement('categories')
  foreach ($c in @($codes | Sort-Object -Unique)) {
    $w.WriteStartElement('category'); $w.WriteAttributeString('id', $c); $w.WriteAttributeString('sy_code', $c); $w.WriteString($SyCatName[$c]); $w.WriteEndElement()
  }
  foreach ($k in $fallback.Keys) { $w.WriteStartElement('category'); $w.WriteAttributeString('id', $k); $w.WriteString($fallback[$k]); $w.WriteEndElement() }
  $w.WriteEndElement()
}

# Synthetic-only offer tags
function Write-SyOffer($w, $map, [int]$warranty, [decimal]$oldPrice, [decimal]$price, $stock) {
  if ($map) { $w.WriteElementString('sy_type', $map[1]) }
  $w.WriteElementString('vendor', $SyVendor)
  $w.WriteElementString('state', 'new')
  if ($null -ne $stock) { $w.WriteElementString('stock_quantity', "$stock") }
  if ($warranty -gt 0) { $w.WriteElementString('sy_warranty', "$warranty") }
  if ($oldPrice -gt $price) { $w.WriteElementString('sy_price_old', $oldPrice.ToString('0.##', [Globalization.CultureInfo]::InvariantCulture)) }
}
