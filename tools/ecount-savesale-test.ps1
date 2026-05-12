# ECOUNT API 검증 — OAPILogin + Sale/SaveSale (로그인 호스트 = SaveSale 호스트 필수)
#
# 1) 로그인만 성공 확인 (Data.Code=00, SESSION_ID 존재):
#    .\ecount-savesale-test.ps1 -LoginOnly -UserId '...' -UserPw '...' -ApiKey '...' [-ApiOrigin https://sboapiCD.ecount.com]
#
# 2) 판매저장까지 (이카운트에 있는 품목/창고/거래처 코드로 채울 것):
#    .\ecount-savesale-test.ps1 -UserId '...' -UserPw '...' -ApiKey '...' -WhCd '유' -ProdCd '...' -ProdDes '...' -CustCd '...'

param(
  [string] $ComCode = "89221",
  [string] $UserId = "",
  [string] $UserPw = "",
  [string] $ApiKey = "",
  [string] $Lang = "ko-KR",
  [string] $WhCd = "",
  [string] $ProdCd = "",
  [string] $ProdDes = "",
  [string] $CustCd = "",
  [string] $CustDes = "",
  [string] $EmpCd = "",
  [string] $IoType = "",
  [string] $Remark = "",
  [Parameter(Mandatory = $false)][AllowEmptyString()] [string] $DocNoUser = "",  # 비우면 예제처럼 DOC_NO ""
  [string] $Qty = "1",
  [string] $Price = "",
  [string] $IoDate = "",  # default: today yyyyMMdd
  [string] $ApiOrigin = "",  # 권장: https://sboapi{ZONE}.ecount.com (로그인 성공한 호스트와 동일)
  [string] $SessionId = "",  # 붙여넣기 세션 — 반드시 -ApiOrigin 과 짝을 맞출 것
  [switch] $LoginOnly,
  [switch] $FullEcounExampleBody  # 이카운트 개발자센터 예제와 동일한 BulkDatas 키 전체 포함
)

$ErrorActionPreference = "Stop"
if (-not $IoDate) { $IoDate = Get-Date -Format "yyyyMMdd" }
$docNo = ""
if ($DocNoUser.Trim()) {
  $docNo = $DocNoUser.Trim()
} else {
  $docNo = "WMS-PS1-" + (Get-Date -Format "yyyyMMddHHmmss")
}

$zone = ""
$origin = ""

if ($SessionId) {
  if ($ApiOrigin) {
    $origin = $ApiOrigin.TrimEnd('/')
  } else {
    $zoneResp = Invoke-RestMethod -Uri "https://oapi.ecount.com/OAPI/V2/Zone" -Method POST `
      -ContentType "application/json; charset=utf-8" `
      -Body (@{ COM_CODE = $ComCode } | ConvertTo-Json -Compress)
    if ($zoneResp.ZONE) { $zone = [string]$zoneResp.ZONE }
    elseif ($zoneResp.Data -and $zoneResp.Data.ZONE) { $zone = [string]$zoneResp.Data.ZONE }
    $zone = $zone.Trim()
    if (-not $zone) { throw "ZONE lookup failed (need COM_CODE for SaveSale host)." }
    $origin = "https://oapi$zone.ecount.com"
  }
  $sessionId = $SessionId.Trim()
  Write-Host "Skip login; using provided SESSION_ID (length=$($sessionId.Length)), origin=$origin"
} else {
  if (-not $UserId -or -not $UserPw -or -not $ApiKey) {
    throw "Either pass -SessionId, or pass -UserId, -UserPw, -ApiKey for OAPILogin."
  }
  $zoneResp = Invoke-RestMethod -Uri "https://oapi.ecount.com/OAPI/V2/Zone" -Method POST `
    -ContentType "application/json; charset=utf-8" `
    -Body (@{ COM_CODE = $ComCode } | ConvertTo-Json -Compress)

  if ($zoneResp.ZONE) { $zone = [string]$zoneResp.ZONE }
  elseif ($zoneResp.Data -and $zoneResp.Data.ZONE) { $zone = [string]$zoneResp.Data.ZONE }
  $zone = $zone.Trim()
  if (-not $zone) {
    Write-Host ($zoneResp | ConvertTo-Json -Depth 6)
    throw "ZONE lookup failed. Check COM_CODE."
  }
  Write-Host "ZONE=$zone"

  if ($ApiOrigin) {
    $origin = $ApiOrigin.TrimEnd('/')
  } else {
    $origin = "https://oapi$zone.ecount.com"
  }

  $loginObj = @{
    COM_CODE      = $ComCode
    USER_ID       = $UserId
    USER_PW       = $UserPw
    API_CERT_KEY  = $ApiKey
    LAN_TYPE      = $Lang
    ZONE          = $zone
    DOMAIN        = ""
  }
  $loginBody = $loginObj | ConvertTo-Json -Compress

  function Get-SessionIdFromLogin($o) {
    if ($null -eq $o) { return "" }
    if ($o.SESSION_ID) { return [string]$o.SESSION_ID }
    if ($o.session_id) { return [string]$o.session_id }
    if ($o.Data -and $o.Data.SESSION_ID) { return [string]$o.Data.SESSION_ID }
    if ($o.Data -and $o.Data.session_id) { return [string]$o.Data.session_id }
    # sboapi 계열: SESSION_ID 가 Data.Datas 안에 올 수 있음
    if ($o.Data -and $o.Data.Datas) {
      $d = $o.Data.Datas
      if ($d.SESSION_ID) { return [string]$d.SESSION_ID }
      if ($d.session_id) { return [string]$d.session_id }
    }
    return ""
  }

  $sessionId = ""
  $loginOriginsToTry = @($origin)
  if (-not $ApiOrigin) {
    $alt = if ($origin -match 'oapi') { "https://sboapi$zone.ecount.com" } else { "https://oapi$zone.ecount.com" }
    if (-not ($loginOriginsToTry -contains $alt)) { $loginOriginsToTry += $alt }
  }

  foreach ($tryOrigin in $loginOriginsToTry) {
    $loginUrl = "$tryOrigin/OAPI/V2/OAPILogin"
    Write-Host "OAPILogin try: $loginUrl"
    try {
      $loginRaw = Invoke-WebRequest -Uri $loginUrl -Method POST `
        -ContentType "application/json; charset=utf-8" `
        -Body ([System.Text.Encoding]::UTF8.GetBytes($loginBody)) -UseBasicParsing
    } catch {
      Write-Host "  request failed: $($_.Exception.Message)"
      continue
    }
    Write-Host "  HTTP $($loginRaw.StatusCode)"
    $loginData = $loginRaw.Content | ConvertFrom-Json
    $sessionId = Get-SessionIdFromLogin $loginData
    if ($sessionId) {
      $origin = $tryOrigin
      Write-Host "SESSION_ID ok (length=$($sessionId.Length)), using origin=$origin"
      break
    }
    $utf8 = try {
      [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::GetEncoding(28591).GetBytes($loginRaw.Content))
    } catch {
      $loginRaw.Content
    }
    Write-Host ($utf8.Substring(0, [Math]::Min(1200, $utf8.Length)))
  }

  if (-not $sessionId) {
    throw "No SESSION_ID from any login URL. Fix API key (test/production), USER_ID/PW, COM_CODE."
  }
}

if ($LoginOnly) {
  Write-Host ""
  Write-Host "=== LOGIN_API_SUCCESS ==="
  Write-Host "HOST: $origin"
  Write-Host "SESSION_ID(first-32): $($sessionId.Substring(0, [Math]::Min(32, $sessionId.Length)))..."
  Write-Host "(exit 0)"
  exit 0
}

if (-not $WhCd.Trim()) { throw "SaveSale 검증에는 -WhCd(이카운트 창고코드)가 필요합니다." }
if (-not $ProdCd.Trim()) { throw "SaveSale 검증에는 -ProdCd(이카운트 품목코드)가 필요합니다." }
if (-not $ProdDes.Trim()) { throw "SaveSale 검증에는 -ProdDes(품목명)가 필요합니다. API 필수 검증 때문입니다." }

function New-EcountSaveSaleBulkDatasFullTemplate {
  param(
    [string] $IoDateVal,
    [string] $WhCdVal,
    [string] $CustVal,
    [string] $CustDesVal,
    [string] $EmpVal,
    [string] $IoTypeVal,
    [string] $DocNoVal,
    [string] $RemarkVal,
    [string] $ProdCdVal,
    [string] $ProdDesVal,
    [string] $QtyVal,
    [string] $PriceVal
  )

  # 이카운트 Sale/SaveSale 예제 Parameter 전체 필드 (값 없으면 빈 문자열)
  return @{
    IO_DATE           = $IoDateVal
    UPLOAD_SER_NO     = ""
    CUST              = $CustVal
    CUST_DES          = $CustDesVal
    EMP_CD            = $EmpVal
    WH_CD             = $WhCdVal
    IO_TYPE           = $IoTypeVal
    EXCHANGE_TYPE     = ""
    EXCHANGE_RATE     = ""
    SITE              = ""
    PJT_CD            = ""
    DOC_NO            = $DocNoVal
    TTL_CTT           = ""
    U_MEMO1           = ""
    U_MEMO2           = ""
    U_MEMO3           = ""
    U_MEMO4           = ""
    U_MEMO5           = ""
    ADD_TXT_01_T      = ""
    ADD_TXT_02_T      = ""
    ADD_TXT_03_T      = ""
    ADD_TXT_04_T      = ""
    ADD_TXT_05_T      = ""
    ADD_TXT_06_T      = ""
    ADD_TXT_07_T      = ""
    ADD_TXT_08_T      = ""
    ADD_TXT_09_T      = ""
    ADD_TXT_10_T      = ""
    ADD_NUM_01_T      = ""
    ADD_NUM_02_T      = ""
    ADD_NUM_03_T      = ""
    ADD_NUM_04_T      = ""
    ADD_NUM_05_T      = ""
    ADD_CD_01_T       = ""
    ADD_CD_02_T       = ""
    ADD_CD_03_T       = ""
    ADD_DATE_01_T     = ""
    ADD_DATE_02_T     = ""
    ADD_DATE_03_T     = ""
    U_TXT1            = ""
    ADD_LTXT_01_T     = ""
    ADD_LTXT_02_T     = ""
    ADD_LTXT_03_T     = ""
    PROD_CD           = $ProdCdVal
    PROD_DES          = $ProdDesVal
    SIZE_DES          = ""
    UQTY              = ""
    QTY               = $QtyVal
    PRICE             = $PriceVal
    USER_PRICE_VAT    = ""
    SUPPLY_AMT        = ""
    SUPPLY_AMT_F      = ""
    VAT_AMT           = ""
    REMARKS           = $RemarkVal
    ITEM_CD           = ""
    P_REMARKS1        = ""
    P_REMARKS2        = ""
    P_REMARKS3        = ""
    ADD_TXT_01        = ""
    ADD_TXT_02        = ""
    ADD_TXT_03        = ""
    ADD_TXT_04        = ""
    ADD_TXT_05        = ""
    ADD_TXT_06        = ""
    REL_DATE          = ""
    REL_NO            = ""
    MAKE_FLAG         = ""
    CUST_AMT          = ""
    P_AMT1            = ""
    P_AMT2            = ""
    ADD_NUM_01        = ""
    ADD_NUM_02        = ""
    ADD_NUM_03        = ""
    ADD_CD_01         = ""
    ADD_CD_02         = ""
    ADD_CD_03         = ""
    ADD_CD_NM_01      = ""
    ADD_CD_NM_02      = ""
    ADD_CD_NM_03      = ""
    ADD_CDNM_01       = ""
    ADD_CDNM_02       = ""
    ADD_CDNM_03       = ""
    ADD_DATE_01       = ""
    ADD_DATE_02       = ""
    ADD_DATE_03       = ""
  }
}

$enc = [System.Uri]::EscapeDataString($sessionId)
$saveUrl = "$origin/OAPI/V2/Sale/SaveSale?SESSION_ID=$enc"

$useDoc = if ($DocNoUser.Trim()) { $DocNoUser.Trim() } else { "" }
$useRemark = if ($Remark.Trim()) { $Remark.Trim() } else { "WMS PowerShell SaveSale (full template)" }

$bulk = if ($FullEcounExampleBody) {
  New-EcountSaveSaleBulkDatasFullTemplate -IoDateVal $IoDate -WhCdVal $WhCd -CustVal $CustCd -CustDesVal $CustDes `
    -EmpVal $EmpCd -IoTypeVal $IoType -DocNoVal $useDoc -RemarkVal $useRemark -ProdCdVal $ProdCd -ProdDesVal $ProdDes `
    -QtyVal $Qty -PriceVal $Price
} else {
  @{
    IO_DATE         = $IoDate
    UPLOAD_SER_NO   = ""
    CUST            = $CustCd
    CUST_DES        = $CustDes
    EMP_CD          = $EmpCd
    WH_CD           = $WhCd
    IO_TYPE         = $IoType
    EXCHANGE_TYPE   = ""
    EXCHANGE_RATE   = ""
    REMARKS         = $useRemark
    DOC_NO          = $(if ($useDoc) { $useDoc } else { $docNo })
    PROD_CD         = $ProdCd
    PROD_DES        = $ProdDes
    QTY             = $Qty
    PRICE           = $(if ($Price -ne "") { $Price } else { "0" })
    SUPPLY_AMT      = "0"
    VAT_AMT         = "0"
    REMARKS1        = ""
    REMARKS2        = ""
    REMARKS3        = ""
  }
}

$saleBodyObj = @{ SaleList = @(@{ BulkDatas = $bulk }) }

$saleJson = $saleBodyObj | ConvertTo-Json -Depth 12 -Compress
Write-Host "SaveSale POST (IO_DATE=$IoDate, DOC_NO=$(if ($useDoc) { $useDoc } else { '(auto/minimal)' }), FullTemplate=$FullEcounExampleBody)"
$saveRaw = Invoke-WebRequest -Uri $saveUrl -Method POST `
  -ContentType "application/json; charset=utf-8" `
  -Body ([System.Text.Encoding]::UTF8.GetBytes($saleJson)) -UseBasicParsing

Write-Host "SaveSale HTTP $($saveRaw.StatusCode)"
$ct = $saveRaw.Content
if ($ct.Length -eq 0) {
  Write-Host "(empty JSON body — 거의 항상 실패·권한·URL 불일치 케이스)"
  exit 2
}

Write-Host "--- SaveSale response (truncated) ---"
Write-Host $ct.Substring(0, [Math]::Min(5000, $ct.Length))

$saleResp = $ct | ConvertFrom-Json
$st = [string]$saleResp.Status

# 이카운트 V2 응답: Status 문자열 "200", Data.SuccessCnt / FailCnt
$okSale = $false
if ($st -eq "200" -and $saleResp.Data) {
  $scRaw = $saleResp.Data.SuccessCnt
  $fcRaw = $saleResp.Data.FailCnt
  $sc = 0; $fc = 0
  [void][int]::TryParse("$scRaw", [ref]$sc)
  [void][int]::TryParse("$fcRaw", [ref]$fc)
  if ($fc -eq 0 -and $sc -gt 0) { $okSale = $true }
}

Write-Host ""
if ($okSale) {
  Write-Host "[SUCCESS]"
  Write-Host ($saleResp | ConvertTo-Json -Depth 12)
  exit 0
}

Write-Host "[FAIL]"
$errMsg = ""
if ($saleResp.Error -and $saleResp.Error.Message) { $errMsg = [string]$saleResp.Error.Message }
Write-Host "Status=$st  Error.Message=$errMsg"
Write-Host ($saleResp | ConvertTo-Json -Depth 8)

exit 1
