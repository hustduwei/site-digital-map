#Requires -Version 5.1
# Sync Excel -> index.html MOCK. Run: powershell -ExecutionPolicy Bypass -File .\sync-from-excel.ps1
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$HtmlPath = Join-Path $Root "index.html"
$CoordsPath = Join-Path $Root "project-coords.json"
$GeoPath = Join-Path $Root "geo-maps.json"

function Find-ExcelFile {
  $files = @(
    [IO.Directory]::GetFiles($Root, "*.xlsx") |
      Where-Object { -not ([IO.Path]::GetFileName($_).StartsWith("~$")) } |
      Sort-Object { [IO.File]::GetLastWriteTime($_) } -Descending
  )
  if (-not $files) { throw "No xlsx in $Root" }
  $prefer = @($files | Where-Object { $_ -match "总览|项目" })
  if ($prefer.Count -gt 0) { return $prefer[0] }
  return $files[0]
}

function Get-SharedStrings([string]$ssPath) {
  [xml]$ss = Get-Content -LiteralPath $ssPath -Encoding UTF8
  $list = New-Object Collections.Generic.List[string]
  foreach ($si in $ss.sst.si) {
    if ($null -ne $si.t) { [void]$list.Add([string]$si.t) }
    else {
      $parts = @(); foreach ($r in $si.r) { $parts += [string]$r.t }
      [void]$list.Add(($parts -join ""))
    }
  }
  return $list
}

function Read-ExcelRows([string]$xlsxPath) {
  $copy = Join-Path $env:TEMP ("gongdi_sync_" + [guid]::NewGuid().ToString("N") + ".xlsx")
  $tmp = Join-Path $env:TEMP ("gongdi_sync_ex_" + [guid]::NewGuid().ToString("N"))
  Copy-Item -LiteralPath $xlsxPath -Destination $copy -Force
  New-Item -ItemType Directory -Path $tmp | Out-Null
  try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [IO.Compression.ZipFile]::ExtractToDirectory($copy, $tmp)
    $strings = Get-SharedStrings (Join-Path $tmp "xl\sharedStrings.xml")
    [xml]$sheet = Get-Content -LiteralPath (Join-Path $tmp "xl\worksheets\sheet1.xml") -Encoding UTF8
    $rows = @{}
    foreach ($row in $sheet.worksheet.sheetData.row) {
      $rnum = [int]$row.r
      $vals = @{}
      foreach ($c in $row.c) {
        if ($c.r -notmatch "^([A-Z]+)(\d+)$") { continue }
        $v = $c.v
        if ($c.t -eq "s") { $v = $strings[[int]$v] }
        $vals[$Matches[1]] = [string]$v
      }
      $rows[$rnum] = $vals
    }
    # 新表头：B项目 C单位 D类型 E地点 F日期 G状态 H对接人 I电话 J经纬度；设备区独立（B种类/C数量）
    # 旧表头兼容：C类型 D地点 E经纬度；设备在 L/M
    $header = if ($rows.ContainsKey(2)) { $rows[2] } else { @{} }
    $hE = (($header["E"] + "").Trim())
    $hJ = (($header["J"] + "").Trim())
    $useNewLayout = ($hE -match "地点|城市") -or ($hJ -match "经纬|坐标|高德")
    $validTypes = @("房建", "厂房", "基建", "公建")
    $projects = @()
    $deviceMap = @{}
    $inDeviceSection = $false
    foreach ($key in ($rows.Keys | Sort-Object)) {
      if ($key -lt 3) { continue }
      $v = $rows[$key]
      $colB = (($v["B"] + "").Trim())
      $colC = (($v["C"] + "").Trim())

      # 设备区：标题行 / 表头行 / 数据行（B=种类 C=数量）
      if ($colB -match "接入设备|种类和数量") { $inDeviceSection = $true; continue }
      if ($colB -eq "种类" -and $colC -match "数量") { $inDeviceSection = $true; continue }
      if ($inDeviceSection) {
        if ($colB -and $colC -match "^\d+") {
          $qty = [int]$Matches[0]
          if (-not $deviceMap.ContainsKey($colB)) { $deviceMap[$colB] = 0 }
          $deviceMap[$colB] += $qty
        }
        continue
      }

      # 旧版并列表设备（L/M）
      $dtype = (($v["L"] + "").Trim())
      $dqtyRaw = (($v["M"] + "").Trim())
      if ($dtype -and $dqtyRaw -match "^\d+" -and $dtype -notmatch "种类") {
        $qty = [int]$Matches[0]
        if (-not $deviceMap.ContainsKey($dtype)) { $deviceMap[$dtype] = 0 }
        $deviceMap[$dtype] += $qty
      }

      $name = $colB
      if (-not $name) { continue }
      if ($name -match "接入设备|种类和数量|^种类$") { continue }

      if ($useNewLayout) {
        $type = (($v["D"] + "").Trim())
        $city = (($v["E"] + "").Trim())
        $coordRaw = (($v["J"] + "").Trim())
        $start = (($v["F"] + "").Trim())
        $status = (($v["G"] + "").Trim())
        $owner = (($v["H"] + "").Trim())
        $phone = (($v["I"] + "").Trim())
        $unit = $colC
      } else {
        $type = $colC
        $city = (($v["D"] + "").Trim())
        $coordRaw = (($v["E"] + "").Trim())
        $start = (($v["F"] + "").Trim())
        $status = (($v["G"] + "").Trim())
        $owner = (($v["H"] + "").Trim())
        $phone = (($v["I"] + "").Trim())
        $unit = ""
      }
      if ($validTypes -notcontains $type) { continue }

      # 状态文案与大屏筛选/图例对齐
      if ($status -match "规划") { $status = "规划对接" }
      elseif ($status -match "持续") { $status = "持续应用中" }
      elseif ($status -match "结束") { $status = "应用结束" }

      $projects += [pscustomobject]@{
        name=$name; type=$type; city=$city; unit=$unit
        coordRaw=$coordRaw
        start=$start; status=$status
        owner=$owner; phone=$phone
      }
    }
    return [pscustomobject]@{ projects = $projects; deviceMap = $deviceMap }
  } finally {
    Remove-Item -LiteralPath $copy -Force -EA SilentlyContinue
    Remove-Item -LiteralPath $tmp -Recurse -Force -EA SilentlyContinue
  }
}

function Escape-Js([string]$s) {
  if ($null -eq $s) { return "" }
  return ($s -replace "\\","\\" -replace '"','\"' -replace "`r","" -replace "`n"," ")
}

function Format-Coord($lng,$lat) {
  $a = [math]::Round([double]$lng,2).ToString([cultureinfo]::InvariantCulture)
  $b = [math]::Round([double]$lat,2).ToString([cultureinfo]::InvariantCulture)
  return "[$a, $b]"
}

function Get-JsonMap($obj) {
  $h = @{}
  $obj.PSObject.Properties | ForEach-Object { $h[$_.Name] = $_.Value }
  return $h
}

function Get-UnitShortName([string]$full) {
  if (-not $full) { return "" }
  if ($full -match "工程总承包") { return "总承包公司" }
  if ($full -match "第一建设") { return "一公司" }
  if ($full -match "第二建设") { return "二公司" }
  if ($full -match "科创") { return "科创公司" }
  if ($full -match "华南") { return "华南公司" }
  if ($full -match "北京") { return "北京公司" }
  if ($full -match "华东") { return "华东公司" }
  if ($full -match "西南") { return "西南公司" }
  if ($full -match "建设发展") { return "建设发展" }
  if ($full -match "铁投|基础设施") { return "铁投公司" }
  if ($full -match "先进技术") { return "先进院" }
  if ($full.Length -gt 8) { return ($full.Substring(0, 8) + [char]0x2026) }
  return $full
}

Write-Host "Root: $Root"
$xlsx = Find-ExcelFile
Write-Host "Excel: $xlsx"
$excelData = Read-ExcelRows $xlsx
$projects = @($excelData.projects)
$deviceMap = @{}
if ($excelData.deviceMap) {
  $excelData.deviceMap.GetEnumerator() | ForEach-Object { $deviceMap[$_.Key] = [int]$_.Value }
}
if ($projects.Count -lt 1) { throw "No project rows" }
Write-Host ("Projects: {0}" -f $projects.Count)
Write-Host ("Device types (Excel): {0}" -f $deviceMap.Count)

$geo = Get-Content -LiteralPath $GeoPath -Raw -Encoding UTF8 | ConvertFrom-Json
$cityProvince = Get-JsonMap $geo.cityProvince
$cityCoord = @{}
$geo.cityCoord.PSObject.Properties | ForEach-Object { $cityCoord[$_.Name] = @([double]$_.Value[0],[double]$_.Value[1]) }
$allProvinces = @($geo.allProvinces)
$provinceCoords = @{}
$geo.provinceCoords.PSObject.Properties | ForEach-Object { $provinceCoords[$_.Name] = @([double]$_.Value[0],[double]$_.Value[1]) }

$coords = @{}
if (Test-Path -LiteralPath $CoordsPath) {
  $obj = Get-Content -LiteralPath $CoordsPath -Raw -Encoding UTF8 | ConvertFrom-Json
  $obj.PSObject.Properties | ForEach-Object { $coords[$_.Name] = @([double]$_.Value[0],[double]$_.Value[1]) }
}

$html = [IO.File]::ReadAllText($HtmlPath, [Text.Encoding]::UTF8)
$progressMap = @{}
foreach ($m in [regex]::Matches($html, '\{ name: "([^"]+)".*?progress: (\d+)')) {
  $progressMap[$m.Groups[1].Value] = [int]$m.Groups[2].Value
}

# 接入设备：KPI 统计全部种类；左下角图表最多 7 项（Top6 + 其他）
# 数量相同时按展示优先级排序；未列入优先级（如 3D打印机）更易归入「其他」
if ($deviceMap.Count -lt 1) { throw "No device rows in Excel (种类/数量)" }
$deviceTypes = $deviceMap.Count
$deviceTotal = ($deviceMap.Values | Measure-Object -Sum).Sum
$deviceChartPriority = @(
  "塔机", "电梯", "安全帽", "无人机", "无人车", "机器狗",
  "地磅", "鹰眼", "远控挖掘机", "远控装载机", "挖掘机"
)
function Get-DeviceChartRank([string]$name) {
  $i = [array]::IndexOf($deviceChartPriority, $name)
  if ($i -ge 0) { return $i }
  return 1000
}
$sortedDevices = @(
  $deviceMap.GetEnumerator() | Sort-Object `
    @{ Expression = { -$_.Value }; Ascending = $true }, `
    @{ Expression = { Get-DeviceChartRank $_.Key }; Ascending = $true }, `
    @{ Expression = { $_.Key }; Ascending = $true }
)
$chartDevices = @()
if ($sortedDevices.Count -le 7) {
  $chartDevices = @($sortedDevices | ForEach-Object { [pscustomobject]@{ name = $_.Key; value = [int]$_.Value } })
} else {
  $top = $sortedDevices | Select-Object -First 6
  $rest = $sortedDevices | Select-Object -Skip 6
  $chartDevices = @($top | ForEach-Object { [pscustomobject]@{ name = $_.Key; value = [int]$_.Value } })
  $otherSum = ($rest | Measure-Object -Property Value -Sum).Sum
  $chartDevices += [pscustomobject]@{ name = "其他"; value = [int]$otherSum }
}
Write-Host ("Device KPI: types={0}, total={1}" -f $deviceTypes, $deviceTotal)
Write-Host ("Device chart ({0}): {1}" -f $chartDevices.Count, (($chartDevices | ForEach-Object { "$($_.name)=$($_.value)" }) -join ", "))

# 项目分布：应用单位覆盖三局二级单位（默认 23 家）
$SECONDARY_UNIT_TOTAL = 23
$unitMap = @{}
foreach ($p in $projects) {
  $u = (($p.unit + "").Trim())
  if (-not $u -or $u -eq "/" -or $u -eq "-" -or $u -eq [string]([char]0x2014)) { continue }
  if (-not $unitMap.ContainsKey($u)) { $unitMap[$u] = 0 }
  $unitMap[$u]++
}
$unitCovered = $unitMap.Count
$unitRows = @(
  $unitMap.GetEnumerator() | Sort-Object `
    @{ Expression = { -$_.Value }; Ascending = $true }, `
    @{ Expression = { $_.Key }; Ascending = $true } |
  ForEach-Object {
    [pscustomobject]@{
      name = (Get-UnitShortName $_.Key)
      full = $_.Key
      value = [int]$_.Value
    }
  }
)
Write-Host ("Unit cover: {0}/{1}" -f $unitCovered, $SECONDARY_UNIT_TOTAL)
Write-Host ("Units ({0}): {1}" -f $unitRows.Count, (($unitRows | ForEach-Object { "$($_.name)=$($_.value)" }) -join ", "))

$byProvince=@{}; $typeCount=@{}; $statusCount=@{}
$missingCoord = New-Object Collections.Generic.List[string]
$newCoords = [ordered]@{}

# 地点为空或「/」时按项目名兜底；表内明显笔误按名称纠正
$cityNameFix = @{
  "胖东来" = "许昌市"
  "重庆石船安置房" = "重庆市"
  "深圳第一职校" = "深圳市"
  "钢丝绳厂" = "武汉市"
  "德阳体育馆" = "德阳市"
}

foreach ($p in $projects) {
  $city = ($p.city + "").Trim()
  if ($cityNameFix.ContainsKey($p.name)) { $city = $cityNameFix[$p.name] }
  elseif (-not $city -or $city -eq "/") {
    if ($p.name -match "深圳") { $city = "深圳市" }
    elseif ($p.name -match "德阳") { $city = "德阳市" }
    elseif ($p.name -match "苏州") { $city = "苏州市" }
    elseif ($p.name -match "武汉|光谷|华科|汉韵|向阳|阳逻|葛店|钢丝|戏曲|华师|楚能|金银湖|九峯|中法同济") { $city = "武汉市" }
  }
  if (-not $cityProvince.ContainsKey($city)) { throw "Unknown city mapping: $city ($($p.name)). Edit geo-maps.json" }
  $prov = [string]$cityProvince[$city]
  $owner = if ($p.owner) { $p.owner } else { [char]0x2014 + "" }; if ($owner -eq "") { $owner = "-" }
  # use em dash as —
  if (-not $p.owner) { $owner = [string]([char]0x2014) }
  $phone = if ($p.phone) { $p.phone } else { [string]([char]0x2014) }
  $start = if (-not $p.start -or $p.start -eq "/") { [string]([char]0x2014) } else { $p.start }
  $status = $p.status
  if (-not $statusCount.ContainsKey($status)) { $statusCount[$status]=0 }
  $statusCount[$status]++
  if (-not $typeCount.ContainsKey($p.type)) { $typeCount[$p.type]=0 }
  $typeCount[$p.type]++

  $lng=$null; $lat=$null
  if ($p.coordRaw -match "^\s*(-?\d+(?:\.\d+)?)\s*[,，]\s*(-?\d+(?:\.\d+)?)\s*$") {
    $lng=[double]$Matches[1]; $lat=[double]$Matches[2]
  } elseif ($coords.ContainsKey($p.name)) {
    $lng=$coords[$p.name][0]; $lat=$coords[$p.name][1]
  } elseif ($cityCoord.ContainsKey($city)) {
    $lng=$cityCoord[$city][0]; $lat=$cityCoord[$city][1]; [void]$missingCoord.Add($p.name)
  } elseif ($provinceCoords.ContainsKey($prov)) {
    $lng=$provinceCoords[$prov][0]; $lat=$provinceCoords[$prov][1]; [void]$missingCoord.Add($p.name)
  } else { $lng=116.4; $lat=39.9; [void]$missingCoord.Add($p.name) }
  $newCoords[$p.name]=@($lng,$lat)

  if ($progressMap.ContainsKey($p.name)) { $progress=$progressMap[$p.name] }
  else {
    if ($status -match "结束") { $progress=100 }
    elseif ($status -match "持续") { $progress=60 }
    else { $progress=15 }
  }

  if ($prov -eq $city -or $city -eq "西藏") {
    if ($city -eq "西藏") { $loc = "西藏自治区" } else { $loc = $prov }
  } else { $loc = "$prov$city" }

  $unit = (($p.unit + "").Trim())
  if (-not $unit -or $unit -eq "/" -or $unit -eq "-") { $unit = [string]([char]0x2014) }

  $item = [pscustomobject]@{ name=$p.name; type=$p.type; city=$city; loc=$loc; progress=$progress; status=$status; start=$start; owner=$owner; phone=$phone; unit=$unit; lng=$lng; lat=$lat; province=$prov }
  if (-not $byProvince.ContainsKey($prov)) { $byProvince[$prov] = New-Object Collections.Generic.List[object] }
  [void]$byProvince[$prov].Add($item)
}

$coordJsonLines = foreach ($k in $newCoords.Keys) {
  "  `"$(Escape-Js $k)`": [{0}, {1}]" -f ([math]::Round($newCoords[$k][0],2).ToString([cultureinfo]::InvariantCulture)), ([math]::Round($newCoords[$k][1],2).ToString([cultureinfo]::InvariantCulture))
}
[IO.File]::WriteAllText($CoordsPath, "{`r`n$($coordJsonLines -join ",`r`n")`r`n}`r`n", [Text.UTF8Encoding]::new($false))

$provinceCount = @($byProvince.Keys).Count
$projectTotal = $projects.Count
$dash = [string]([char]0x2014)

$ppLines = foreach ($name in $allProvinces) {
  $c = if ($byProvince.ContainsKey($name)) { $byProvince[$name].Count } else { 0 }
  "        `"$name`": $c"
}
$bizLines = ($typeCount.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
  "        { name: `"$(Escape-Js $_.Key)`", value: $($_.Value) }"
}) -join ",`r`n"

$sRun = if ($statusCount.ContainsKey("持续应用中")) { [int]$statusCount["持续应用中"] } else { 0 }
$sPlan = if ($statusCount.ContainsKey("规划对接")) { [int]$statusCount["规划对接"] } else { 0 }
$sDone = if ($statusCount.ContainsKey("应用结束")) { [int]$statusCount["应用结束"] } else { 0 }
# fallback match partial
foreach ($k in $statusCount.Keys) {
  if ($k -match "持续") { $sRun = [int]$statusCount[$k] }
  if ($k -match "规划") { $sPlan = [int]$statusCount[$k] }
  if ($k -match "结束") { $sDone = [int]$statusCount[$k] }
}

$pcLines = @(
  "        { name: `"持续应用中`", value: $sRun }",
  "        { name: `"规划对接`", value: $sPlan }",
  "        { name: `"应用结束`", value: $sDone }"
) -join ",`r`n"

$pcCoordLines = ($provinceCoords.GetEnumerator() | ForEach-Object {
  "        `"$($_.Key)`": [{0}, {1}]" -f $_.Value[0].ToString([cultureinfo]::InvariantCulture), $_.Value[1].ToString([cultureinfo]::InvariantCulture)
}) -join ",`r`n"

$cityBlocks = foreach ($prov in ($byProvince.Keys | Sort-Object)) {
  $itemLines = foreach ($it in $byProvince[$prov]) {
    "          { name: `"$(Escape-Js $it.name)`", type: `"$(Escape-Js $it.type)`", city: `"$(Escape-Js $it.city)`", loc: `"$(Escape-Js $it.loc)`", progress: $($it.progress), status: `"$(Escape-Js $it.status)`", start: `"$(Escape-Js $it.start)`", owner: `"$(Escape-Js $it.owner)`", phone: `"$(Escape-Js $it.phone)`", unit: `"$(Escape-Js $it.unit)`", coord: $(Format-Coord $it.lng $it.lat) }"
  }
  "        `"$prov`": [`r`n$($itemLines -join ",`r`n")`r`n        ]"
}
$cityProjectsBlock = $cityBlocks -join ",`r`n"

$mock = @"
    /* === MOCK_DATA_START === */
    const MOCK = {
      kpi: {
        projectTotal: $projectTotal,
        provinceCount: $provinceCount,
        deviceTypes: $deviceTypes,
        deviceTotal: $deviceTotal
      },
      weatherTemp: 26,
      bizTypes: [
$bizLines
      ],
      devices: [
$(($chartDevices | ForEach-Object { "        { name: `"$(Escape-Js $_.name)`", value: $($_.value) }" }) -join ",`r`n")
      ],
      unitCover: {
        covered: $unitCovered,
        total: $SECONDARY_UNIT_TOTAL,
        units: [
$(($unitRows | ForEach-Object { "          { name: `"$(Escape-Js $_.name)`", full: `"$(Escape-Js $_.full)`", value: $($_.value) }" }) -join ",`r`n")
        ]
      },
      trend: {
        months: ["25-01","25-10","25-11","25-12","26-01","26-02","26-03","26-04","26-06","规划"],
        values: [2, 1, 1, 1, 2, 1, 2, 1, 1, 9]
      },
      provinceProjects: {
$($ppLines -join ",`r`n")
      },
      progressCats: [
$pcLines
      ],
      provinceCoords: {
$pcCoordLines
      },
      cityProjects: {
$cityProjectsBlock
      }
    };
    /* === MOCK_DATA_END === */
"@

if ($html -match "(?s)/\* === MOCK_DATA_START === \*/.*?/\* === MOCK_DATA_END === \*/") {
  $html2 = [regex]::Replace($html, "(?s)/\* === MOCK_DATA_START === \*/.*?/\* === MOCK_DATA_END === \*/", [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $mock.TrimEnd() })
} else {
  throw "MOCK markers not found in index.html"
}

# 同步 KPI 初始 data-target（与 MOCK.kpi 一致）
$kpiTargets = @($projectTotal, $provinceCount, $deviceTypes, $deviceTotal)
$kpiIdx = 0
$html2 = [regex]::Replace($html2, 'class="kpi-value" data-target="\d+"', {
  param($m)
  if ($kpiIdx -ge $kpiTargets.Count) { return $m.Value }
  $t = $kpiTargets[$kpiIdx]
  $script:kpiIdx++
  "class=`"kpi-value`" data-target=`"$t`""
})

[IO.File]::WriteAllText($HtmlPath, $html2, [Text.UTF8Encoding]::new($false))
Write-Host "Updated index.html"
Write-Host "Updated project-coords.json"
if ($missingCoord.Count -gt 0) {
  Write-Host "Approx coords (edit project-coords.json):"
  $missingCoord | ForEach-Object { Write-Host "  - $_" }
}
Write-Host "Done."