$ErrorActionPreference = "Stop"

$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$out = Join-Path $root "SpotFinder-Presentasi-Canva-Editable-Final.pptx"
$tmp = Join-Path $env:TEMP ("spotfinder_pptx_" + [guid]::NewGuid().ToString("N"))

function XmlEscape($value) {
    return [System.Security.SecurityElement]::Escape([string]$value)
}

function Write-Utf8($path, $content) {
    $dir = Split-Path -Parent $path
    if (!(Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    [System.IO.File]::WriteAllText($path, $content, [System.Text.UTF8Encoding]::new($false))
}

function ShapeXml($id, $name, $x, $y, $w, $h, $text, $size, $color, $bold, $fill = $null) {
    $paras = @()
    foreach ($line in ([string]$text -split "`n")) {
        $b = if ($bold) { ' b="1"' } else { "" }
        $paras += @"
<a:p><a:r><a:rPr lang="id-ID" sz="$($size * 100)"$b><a:solidFill><a:srgbClr val="$color"/></a:solidFill></a:rPr><a:t>$(XmlEscape $line)</a:t></a:r></a:p>
"@
    }
    $fillXml = if ($fill) {
        "<a:solidFill><a:srgbClr val=""$fill""/></a:solidFill>"
    } else {
        "<a:solidFill><a:srgbClr val=""FFFFFF""><a:alpha val=""0""/></a:srgbClr></a:solidFill>"
    }
    return @"
<p:sp>
  <p:nvSpPr><p:cNvPr id="$id" name="$(XmlEscape $name)"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="$x" y="$y"/><a:ext cx="$w" cy="$h"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>$fillXml<a:ln><a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:solidFill></a:ln></p:spPr>
  <p:txBody><a:bodyPr wrap="square" anchor="mid"/><a:lstStyle/>$($paras -join "")</p:txBody>
</p:sp>
"@
}

function BulletsXml($startId, $x, $y, $w, $h, $items) {
    $paras = @()
    foreach ($item in $items) {
        $paras += @"
<a:p><a:pPr marL="342900" indent="-228600"><a:buChar char="•"/></a:pPr><a:r><a:rPr lang="id-ID" sz="2000"><a:solidFill><a:srgbClr val="A8B9C2"/></a:solidFill></a:rPr><a:t>$(XmlEscape $item)</a:t></a:r></a:p>
"@
    }
    return @"
<p:sp>
  <p:nvSpPr><p:cNvPr id="$startId" name="Bullet List"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="$x" y="$y"/><a:ext cx="$w" cy="$h"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:solidFill></p:spPr>
  <p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>$($paras -join "")</p:txBody>
</p:sp>
"@
}

function ImageXml($id, $name, $rid, $x, $y, $w, $h) {
    return @"
<p:pic>
  <p:nvPicPr><p:cNvPr id="$id" name="$(XmlEscape $name)"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
  <p:blipFill><a:blip r:embed="$rid"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
  <p:spPr><a:xfrm><a:off x="$x" y="$y"/><a:ext cx="$w" cy="$h"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
</p:pic>
"@
}

function CardXml($id, $x, $y, $w, $h, $title, $body) {
    return @"
<p:sp>
  <p:nvSpPr><p:cNvPr id="$id" name="Card"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="$x" y="$y"/><a:ext cx="$w" cy="$h"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="13242E"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="34515E"/></a:solidFill></a:ln></p:spPr>
  <p:txBody><a:bodyPr lIns="190500" tIns="152400" rIns="190500" bIns="152400"/><a:lstStyle/>
    <a:p><a:r><a:rPr lang="id-ID" sz="2200" b="1"><a:solidFill><a:srgbClr val="EEF7F4"/></a:solidFill></a:rPr><a:t>$(XmlEscape $title)</a:t></a:r></a:p>
    <a:p><a:r><a:rPr lang="id-ID" sz="1500"><a:solidFill><a:srgbClr val="A8B9C2"/></a:solidFill></a:rPr><a:t>$(XmlEscape $body)</a:t></a:r></a:p>
  </p:txBody>
</p:sp>
"@
}

function SlideXml($bodyXml) {
    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="071016"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      $bodyXml
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
"@
}

New-Item -ItemType Directory -Path $tmp -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmp "ppt\slides\_rels") -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmp "ppt\media") -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmp "ppt\_rels") -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmp "_rels") -Force | Out-Null

$media = @{
    "hero.png" = Join-Path $root "assets\img\hero.png"
    "screen1.png" = Join-Path $root "assets\img\screen-1.png"
    "screen2.png" = Join-Path $root "assets\img\screen-2.png"
    "screen3.png" = Join-Path $root "assets\img\screen-3.png"
    "logo.png" = Join-Path $root "assets\img\favicon.png"
}
foreach ($item in $media.GetEnumerator()) {
    Copy-Item $item.Value (Join-Path $tmp ("ppt\media\" + $item.Key)) -Force
}

$slides = @()

$slides += @{
    Body = (ShapeXml 2 "Title" 600000 760000 7200000 1500000 "SPOTFINDER" 58 "EEF7F4" $true) +
           (ShapeXml 3 "Subtitle" 620000 2200000 6400000 1100000 "Sistem Parkir Cerdas Berbasis IoT dan Web" 28 "35D0FF" $true) +
           (ShapeXml 4 "Byline" 620000 3600000 5200000 750000 "By: Zahran Al Syafit`nDurasi presentasi: 10 menit" 20 "A8B9C2" $false) +
           (ImageXml 5 "Hero" "rId2" 7850000 850000 3300000 3900000) +
           (ImageXml 6 "Logo" "rId1" 620000 5550000 520000 520000)
    Rels = @(@("rId1","../media/logo.png"), @("rId2","../media/hero.png"))
}

$slides += @{
    Body = (ShapeXml 2 "Title" 560000 520000 9400000 800000 "Masalah yang Ingin Diselesaikan" 34 "EEF7F4" $true) +
           (BulletsXml 3 720000 1600000 5300000 3300000 @("Pengguna sering membuang waktu mencari slot kosong.", "Pengelola sulit melihat kapasitas parkir secara cepat.", "Transaksi dan riwayat kendaraan belum terpusat.", "Palang dan laporan masih banyak dilakukan secara manual.")) +
           (CardXml 4 6800000 1600000 1900000 1450000 "Waktu" "Pencarian parkir lebih lama.") +
           (CardXml 5 8950000 1600000 1900000 1450000 "Data" "Status dan riwayat belum rapi.") +
           (CardXml 6 6800000 3350000 1900000 1450000 "Akses" "Kontrol masuk-keluar kurang otomatis.") +
           (CardXml 7 8950000 3350000 1900000 1450000 "Laporan" "Rekap operasional lambat.")
    Rels = @()
}

$slides += @{
    Body = (ShapeXml 2 "Title" 560000 520000 9600000 800000 "Tujuan SpotFinder" 34 "EEF7F4" $true) +
           (ShapeXml 3 "Lead" 620000 1400000 9800000 900000 "Membuat sistem parkir yang lebih cepat, tertata, real-time, dan mudah dikelola oleh pengguna maupun admin." 23 "A8B9C2" $false) +
           (CardXml 4 700000 2750000 2350000 1500000 "Real-time Slot" "Slot tersedia, terisi, booking, dan pelanggaran bisa dipantau langsung.") +
           (CardXml 5 3300000 2750000 2350000 1500000 "RFID Access" "Masuk parkir memakai kartu member yang diverifikasi server.") +
           (CardXml 6 5900000 2750000 2350000 1500000 "Cashless Wallet" "Saldo, top-up, booking fee, dan biaya parkir tersimpan digital.") +
           (CardXml 7 8500000 2750000 2350000 1500000 "Admin Dashboard" "Admin mengelola member, tarif, gate, analytics, dan laporan.")
    Rels = @()
}

$slides += @{
    Body = (ShapeXml 2 "Title" 560000 520000 9600000 800000 "Fitur Utama" 34 "EEF7F4" $true) +
           (CardXml 3 620000 1500000 3300000 1250000 "Client Portal" "Login, registrasi, profil, saldo, riwayat, dan chat support.") +
           (CardXml 4 4450000 1500000 3300000 1250000 "Booking Slot" "Pengguna dapat memesan slot kosong dari dashboard.") +
           (CardXml 5 8280000 1500000 3300000 1250000 "Admin Panel" "Kelola RFID, member, top-up, tarif, history, dan analytics.") +
           (CardXml 6 620000 3250000 3300000 1250000 "IoT Detection" "Sensor ultrasonik membaca keberadaan kendaraan per slot.") +
           (CardXml 7 4450000 3250000 3300000 1250000 "Auto Gate" "Servo membuka palang setelah akses dinyatakan valid.") +
           (CardXml 8 8280000 3250000 3300000 1250000 "AI Voice" "Notifikasi suara untuk welcome, goodbye, dan status sistem.")
    Rels = @()
}

$slides += @{
    Body = (ShapeXml 2 "Title" 560000 520000 9600000 800000 "Cara Kerja Sistem" 34 "EEF7F4" $true) +
           (CardXml 3 520000 1600000 2000000 2500000 "1. Login" "Pengguna masuk atau daftar kendaraan.") +
           (CardXml 4 2700000 1600000 2000000 2500000 "2. Cek Slot" "Dashboard menampilkan slot kosong dan rekomendasi.") +
           (CardXml 5 4880000 1600000 2000000 2500000 "3. Scan RFID" "UID kartu dikirim ke server untuk validasi.") +
           (CardXml 6 7060000 1600000 2000000 2500000 "4. Sensor" "ESP32 membaca slot dan memperbarui status.") +
           (CardXml 7 9240000 1600000 2000000 2500000 "5. Keluar" "Biaya dihitung, riwayat tersimpan, struk bisa dicetak.")
    Rels = @()
}

$slides += @{
    Body = (ShapeXml 2 "Title" 560000 430000 9600000 650000 "Tampilan Aplikasi" 32 "EEF7F4" $true) +
           (ImageXml 3 "Screen 1" "rId1" 820000 1300000 2750000 3550000) +
           (ImageXml 4 "Screen 2" "rId2" 4620000 1300000 2750000 3550000) +
           (ImageXml 5 "Screen 3" "rId3" 8420000 1300000 2750000 3550000) +
           (ShapeXml 6 "Caption1" 820000 5050000 2750000 620000 "Smart Dashboard`nSlot, okupansi, aktivitas, statistik." 15 "A8B9C2" $false) +
           (ShapeXml 7 "Caption2" 4620000 5050000 2750000 620000 "Client Portal`nSaldo, booking, riwayat, profil." 15 "A8B9C2" $false) +
           (ShapeXml 8 "Caption3" 8420000 5050000 2750000 620000 "AI Support Chat`nBantuan client dan admin." 15 "A8B9C2" $false)
    Rels = @(@("rId1","../media/screen1.png"), @("rId2","../media/screen2.png"), @("rId3","../media/screen3.png"))
}

$slides += @{
    Body = (ShapeXml 2 "Title" 560000 520000 9600000 800000 "Teknologi yang Digunakan" 34 "EEF7F4" $true) +
           (BulletsXml 3 700000 1550000 5200000 3600000 @("PHP dan MySQL sebagai backend dan database.", "HTML, CSS, JavaScript, jQuery, dan Chart.js untuk dashboard.", "ESP32 sebagai pusat perangkat IoT.", "RFID RC522 untuk identifikasi pengguna.", "Ultrasonik, OLED, LED, buzzer, dan servo untuk prototype parkir.", "ElevenLabs TTS dan fallback browser voice untuk notifikasi suara.")) +
           (CardXml 4 6900000 1700000 4000000 1150000 "API Integration" "Endpoint menghubungkan web app dengan ESP32.") +
           (CardXml 5 6900000 3150000 4000000 1150000 "Database Tables" "users, sensor_status, parking_history, settings, topup_requests.") +
           (CardXml 6 6900000 4600000 4000000 1150000 "Scalable Idea" "Dapat dikembangkan ke banyak slot, QRIS asli, dan plat recognition.")
    Rels = @()
}

$slides += @{
    Body = (ShapeXml 2 "Title" 560000 520000 9600000 800000 "Cara Penggunaan Saat Demo" 34 "EEF7F4" $true) +
           (CardXml 3 650000 1500000 5000000 3700000 "Untuk Client" "1. Buka Client Portal.`n2. Login atau registrasi kendaraan.`n3. Lihat saldo dan slot real-time.`n4. Booking slot kosong.`n5. Gunakan chat support jika perlu.") +
           (CardXml 4 6400000 1500000 5000000 3700000 "Untuk Admin" "1. Login admin.`n2. Lihat dashboard slot dan aktivitas.`n3. Approve top-up atau tambah RFID.`n4. Cek analytics dan history.`n5. Gunakan gate control untuk palang.")
    Rels = @()
}

$slides += @{
    Body = (ShapeXml 2 "Title" 560000 520000 9600000 800000 "Manfaat dan Kelebihan" 34 "EEF7F4" $true) +
           (CardXml 3 700000 1600000 2350000 1500000 "Efisien" "Mengurangi waktu mencari parkir karena status slot terlihat.") +
           (CardXml 4 3300000 1600000 2350000 1500000 "Terdokumentasi" "Masuk, keluar, booking, top-up, dan biaya tercatat otomatis.") +
           (CardXml 5 5900000 1600000 2350000 1500000 "Modern" "Mendukung wallet, chat, voice notification, dan analytics.") +
           (CardXml 6 8500000 1600000 2350000 1500000 "Bisa Dikembangkan" "Dapat ditambah banyak slot, QRIS asli, dan kamera plat nomor.") +
           (ShapeXml 7 "Summary" 1200000 4200000 9700000 950000 "SpotFinder memberi pengalaman parkir yang lebih jelas untuk pengguna dan data operasional yang lebih rapi untuk pengelola." 24 "35D0FF" $true)
    Rels = @()
}

$slides += @{
    Body = (ImageXml 2 "Logo" "rId1" 5480000 850000 620000 620000) +
           (ShapeXml 3 "Title" 1200000 1750000 9800000 1900000 "Parkir lebih pintar, data lebih rapi, pengalaman lebih cepat." 44 "EEF7F4" $true) +
           (ShapeXml 4 "Closing" 2200000 3900000 7800000 900000 "SpotFinder membuktikan bahwa masalah parkir bisa diselesaikan dengan kombinasi web app, database, IoT, dan desain penggunaan yang jelas." 22 "A8B9C2" $false) +
           (ShapeXml 5 "Thanks" 4750000 5350000 2600000 520000 "Terima kasih" 24 "071016" $true "77E08F")
    Rels = @(@("rId1","../media/logo.png"))
}

$contentTypesSlides = ""
$presentationRels = @('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>')
$slideIds = ""

for ($i = 0; $i -lt $slides.Count; $i++) {
    $n = $i + 1
    Write-Utf8 (Join-Path $tmp "ppt\slides\slide$n.xml") (SlideXml $slides[$i].Body)
    $relLines = @('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
    $relLines += '<Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
    foreach ($rel in $slides[$i].Rels) {
        $relLines += "<Relationship Id=""$($rel[0])"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"" Target=""$($rel[1])""/>"
    }
    $relLines += '</Relationships>'
    Write-Utf8 (Join-Path $tmp "ppt\slides\_rels\slide$n.xml.rels") ($relLines -join "`n")
    $contentTypesSlides += "<Override PartName=""/ppt/slides/slide$n.xml"" ContentType=""application/vnd.openxmlformats-officedocument.presentationml.slide+xml""/>"
    $presentationRels += "<Relationship Id=""rId$($n + 1)"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"" Target=""slides/slide$n.xml""/>"
    $slideIds += "<p:sldId id=""$($n + 255)"" r:id=""rId$($n + 1)""/>"
}

Write-Utf8 (Join-Path $tmp "[Content_Types].xml") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  $contentTypesSlides
</Types>
"@

Write-Utf8 (Join-Path $tmp "_rels\.rels") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>
"@

Write-Utf8 (Join-Path $tmp "ppt\presentation.xml") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>$slideIds</p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000" type="wide"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>
"@

Write-Utf8 (Join-Path $tmp "ppt\_rels\presentation.xml.rels") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  $($presentationRels -join "`n  ")
</Relationships>
"@

New-Item -ItemType Directory -Path (Join-Path $tmp "ppt\slideMasters\_rels") -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmp "ppt\slideLayouts\_rels") -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmp "ppt\theme") -Force | Out-Null

Write-Utf8 (Join-Path $tmp "ppt\slideMasters\slideMaster1.xml") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>
"@

Write-Utf8 (Join-Path $tmp "ppt\slideMasters\_rels\slideMaster1.xml.rels") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>
"@

Write-Utf8 (Join-Path $tmp "ppt\slideLayouts\slideLayout1.xml") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>
"@

Write-Utf8 (Join-Path $tmp "ppt\slideLayouts\_rels\slideLayout1.xml.rels") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>
"@

Write-Utf8 (Join-Path $tmp "ppt\theme\theme1.xml") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="SpotFinder">
  <a:themeElements>
    <a:clrScheme name="SpotFinder"><a:dk1><a:srgbClr val="071016"/></a:dk1><a:lt1><a:srgbClr val="EEF7F4"/></a:lt1><a:dk2><a:srgbClr val="13242E"/></a:dk2><a:lt2><a:srgbClr val="A8B9C2"/></a:lt2><a:accent1><a:srgbClr val="35D0FF"/></a:accent1><a:accent2><a:srgbClr val="77E08F"/></a:accent2><a:accent3><a:srgbClr val="F7C948"/></a:accent3><a:accent4><a:srgbClr val="FF7A7A"/></a:accent4><a:accent5><a:srgbClr val="6EA8FE"/></a:accent5><a:accent6><a:srgbClr val="B38CFF"/></a:accent6><a:hlink><a:srgbClr val="35D0FF"/></a:hlink><a:folHlink><a:srgbClr val="77E08F"/></a:folHlink></a:clrScheme>
    <a:fontScheme name="SpotFinder"><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="SpotFinder"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>
"@

if (Test-Path $out) {
    Remove-Item $out -Force
}
$zipOut = Join-Path $env:TEMP ("spotfinder_pptx_" + [guid]::NewGuid().ToString("N") + ".zip")
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::Open($zipOut, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    $base = (Resolve-Path $tmp).Path.TrimEnd('\') + '\'
    Get-ChildItem -Path $tmp -Recurse -File | ForEach-Object {
        $entryName = $_.FullName.Substring($base.Length).Replace('\', '/')
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $entryName, [System.IO.Compression.CompressionLevel]::Optimal) | Out-Null
    }
}
finally {
    $zip.Dispose()
}
[System.IO.File]::Copy($zipOut, $out, $true)
try { Remove-Item $zipOut -Force } catch { }
Remove-Item $tmp -Recurse -Force
Write-Host $out
