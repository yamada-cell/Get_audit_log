<#
M365_Audit_Full_5_7.ps1
- Microsoft 365 監査ログ（UTCで前日 00:00〜当日 00:00、テスト時は適宜拡張）
- 生ログ + 拡張ログ（ディレクトリ/ファイル名/サイズ）
- 4操作（FileDownloaded / FileUploaded / FileSyncDownloaded / FileSyncUploaded）のユーザー別集計
- 閾値超過時に Teams Workflows Webhook 通知（テキスト）
- 取得CSV/集計CSVをTeamsのファイルタブへ自動アップロード（PnP.PowerShell）
- Exchange Online: App-only（Thumbprint）、PnP.PowerShell: PFX
- Windows PowerShell 5.1 互換（?. / ?? を使用しない）
#>

#========================
# Section 0: 初期化
#========================
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

#========================
# Section 1: パラメータ
#========================
# Exchange Online 認証（App-only / Thumbprint）
$ExoOrganization   = "NihonTechnoSystems.onmicrosoft.com"
$ExoAppId          = "793d5df8-82bd-4c30-8258-55f7b6b98ded"
$ExoCertThumbprint = "50D2CB25C9FD695B2CDE6F81421FBE4AD57D4337"

# PnP.PowerShell 認証（PFX）
$PnPTenant      = "NihonTechnoSystems.onmicrosoft.com"
$PnPClientId    = "0caabba2-7d10-47ab-9b39-06fd62b0ee2a"
$PnPPfxPath     = "C:\Secure\pnp_app.pfx"
$PnPPfxPassword = "aabbCCDD5678"

# Teams（SharePointサイト）保存先
$SiteUrl    = "https://nihontechnosystems.sharepoint.com/sites/msteams_5ff01b"
$FolderPath = "Shared Documents/M365通知テスト用/ログ保存用フォルダ3"

# 出力・閾値・Webhook
$OutputDir       = "C:\Users\山田辰徳\Desktop\log\audit_log"
#$ThresholdCount  = 50
#$ThresholdSizeMB = 10
#$WebhookUrl      = "https://default694c9923d2f040fe9f48b75b4fc88e.4a.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/37198d9ec32f44d0a0952952c57ec5f9/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=F041yno7ZL2MFsesMdk51_1P9M3TYtLAmdx2g6PAVq8"

# 期間（UTC） 前日 00:00 〜 当日 00:00
$StartDateUtc = (Get-Date).Date.AddDays(-1).ToUniversalTime()
$EndDateUtc   = (Get-Date).Date.ToUniversalTime()

# --- Environment overrides (存在すれば上書き) ---
if ($env:EXO_ORG)             { $ExoOrganization   = $env:EXO_ORG }
if ($env:EXO_APP_ID)          { $ExoAppId          = $env:EXO_APP_ID }
if ($env:EXO_CERT_THUMBPRINT) { $ExoCertThumbprint = $env:EXO_CERT_THUMBPRINT }

if ($env:PNP_TENANT)          { $PnPTenant         = $env:PNP_TENANT }
if ($env:PNP_CLIENT_ID)       { $PnPClientId       = $env:PNP_CLIENT_ID }
if ($env:PNP_PFX_PATH)        { $PnPPfxPath        = $env:PNP_PFX_PATH }
if ($env:PNP_PFX_PASSWORD)    { $PnPPfxPassword    = $env:PNP_PFX_PASSWORD }

if ($env:SITE_URL)            { $SiteUrl           = $env:SITE_URL }
if ($env:FOLDER_PATH)         { $FolderPath        = $env:FOLDER_PATH }

if ($env:THRESHOLD_COUNT)     { $ThresholdCount    = [int]$env:THRESHOLD_COUNT }
if ($env:THRESHOLD_SIZE_MB)   { $ThresholdSizeMB   = [int]$env:THRESHOLD_SIZE_MB }
if ($env:WEBHOOK_URL)         { $WebhookUrl        = $env:WEBHOOK_URL }


#========================
# Section 1.5: ユーティリティ
#========================
# UTF-8 BOMでCSV出力（Excel向け）
function Export-CsvUtf8BOM {
    param(
        [Parameter(Mandatory)][object]$InputObject,
        [Parameter(Mandatory)][string]$Path
    )
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    $csvLines = $InputObject | ConvertTo-Csv -NoTypeInformation
    $csvText  = $csvLines -join "`r`n"
    [System.IO.File]::WriteAllText($Path, $csvText, $utf8Bom)
}

# Operation 名を4操作に正規化
function Normalize-Operation([string]$op) {
    if ([string]::IsNullOrWhiteSpace($op)) { return $null }
    if ($op -eq 'FileDownloaded') { return 'FileDownloaded' }
    if ($op -eq 'FileUploaded')   { return 'FileUploaded' }
    if ($op -like 'FileSyncDownloaded*') { return 'FileSyncDownloaded' }
    if ($op -like 'FileSyncUploaded*')   { return 'FileSyncUploaded' }
    return $null
}

# 任意のオブジェクト/Hashtable/JSON文字列から安全にプロパティを取得
function Get-Prop {
    param(
        [Parameter(Mandatory)][object]$Obj,
        [Parameter(Mandatory)][string]$Name
    )
    if ($null -eq $Obj) { return $null }

    # 文字列なら JSON とみなして展開を試みる
    if ($Obj -is [string]) {
        try { $Obj = $Obj | ConvertFrom-Json -ErrorAction Stop } catch { return $null }
    }

    # Hashtable
    if ($Obj -is [hashtable]) {
        if ($Obj.ContainsKey($Name)) { return $Obj[$Name] } else { return $null }
    }

    # PSCustomObject 等
    $p = $Obj.PSObject.Properties[$Name]
    if ($p) { return $p.Value } else { return $null }
}


# 監査レコードのトップレベル（$rec）のプロパティを安全取得（StrictMode対応）
function Get-TopProp {
    param(
        [Parameter(Mandatory)][object]$Obj,
        [Parameter(Mandatory)][string]$Name
    )
    if ($null -eq $Obj) { return $null }
    $p = $Obj.PSObject.Properties[$Name]
    if ($p) { return $p.Value } else { return $null }
}


# AuditData からファイル名/ディレクトリ/サイズをできるだけ抽出（null/Hashtable/PSCustomObject/文字列に対応）
function Parse-FileInfo {
    param(
        [Parameter(Mandatory)][object]$AuditData
    )

    # 候補を安全に取得
    $objId   = Get-Prop $AuditData 'ObjectId'
    $siteUrl = Get-Prop $AuditData 'SiteUrl'
    $srcRel  = Get-Prop $AuditData 'SourceRelativeUrl'
    $srcName = Get-Prop $AuditData 'SourceFileName'
    $tgtRel  = Get-Prop $AuditData 'TargetRelativeUrl'
    $itemUrl = Get-Prop $AuditData 'ItemUrl'

    # サイズ候補（優先: SourceFileSize → FileSize → Size）
    $size = 0
    foreach ($key in @('FileSyncBytesCommitted','FileSizeBytes','SourceFileSize','FileSize','Size')) {
        $val = Get-Prop $AuditData $key
        if ($val) { $size = $val; break }
    }

    # パス候補
    $pathCandidates = @()
    foreach ($p in @($objId, $itemUrl, $srcRel, $tgtRel)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$p)) { $pathCandidates += [string]$p }
    }

    # ファイル名
    $fileName = $null
    if ($srcName) {
        $fileName = [string]$srcName
    } elseif ($pathCandidates.Count -gt 0) {
        try { $fileName = [System.IO.Path]::GetFileName($pathCandidates[0]) } catch { $fileName = $null }
    }

    # ディレクトリ
    $dir = $null
    $firstPath = $pathCandidates | Select-Object -First 1
    if ($firstPath) {
        try { $dir = Split-Path -Path $firstPath -Parent -ErrorAction SilentlyContinue } catch { $dir = $null }
    }
    if ($siteUrl -and $dir -and ($dir -notlike 'http*')) {
        $dir = ($siteUrl.TrimEnd('/') + '/' + $dir.TrimStart('/')).Replace('//','/')
    } elseif ($siteUrl -and -not $dir) {
        $dir = $siteUrl
    }

    # 戻り値（常に PSCustomObject）
    return [PSCustomObject]@{
    FileDirectory = [string]$dir
    FileName      = [string]$fileName
    FileSizeBytes = [long]$size
    }
''
}


#========================
# Section 2: 接続
#========================
# 出力フォルダ
if (-not (Test-Path $OutputDir)) { New-Item -Path $OutputDir -ItemType Directory | Out-Null }

# Exchange Online
#try {
#    Connect-ExchangeOnline -AppId $ExoAppId -CertificateThumbprint $ExoCertThumbprint -Organization $ExoOrganization -ShowBanner:$false
#} catch {
#    Write-Error "Exchange Online接続に失敗: $($_.Exception.Message)"
#    exit 1
#}

# Exchange Online（ファイル方式へ統一。Thumbprintはフォールバック）
try {
   # 事前ガード（StrictMode向け）
   if (-not $ExoAppId)        { throw "EXO AppId が未設定です。Secrets(EXO_APP_ID)またはスクリプト定数を設定してください。" }
   if (-not $ExoOrganization) { throw "EXO Organization が未設定です。Secrets(EXO_ORG)またはスクリプト定数を設定してください。" }

   if ($env:EXO_PFX_PATH -and $env:EXO_PFX_PASSWORD) {
     # Secretsから復元したPFXファイルを直接指定（推奨）
     Connect-ExchangeOnline `
       -AppId $ExoAppId `
       -Organization $ExoOrganization `
       -CertificateFilePath $env:EXO_PFX_PATH `
       -CertificatePassword (ConvertTo-SecureString $env:EXO_PFX_PASSWORD -AsPlainText -Force) `
       -ShowBanner:$false
   }
   elseif ($ExoCertThumbprint) {
     # 既存のThumbprint方式（Windowsランナーでのみ可）
     Connect-ExchangeOnline `
       -AppId $ExoAppId `
       -CertificateThumbprint $ExoCertThumbprint `
       -Organization $ExoOrganization `
       -ShowBanner:$false
   }
   else {
     throw "EXOの証明書指定が不足しています（EXO_PFX_PATH/EXO_PFX_PASSWORD または Thumbprint）。"
   }
 } catch {
   Write-Error "Exchange Online接続に失敗: $($_.Exception.Message)"
   exit 1
 }


# PnP.PowerShell
try {
    Connect-PnPOnline -Url $SiteUrl -ClientId $PnPClientId -Tenant $PnPTenant -CertificatePath $PnPPfxPath -CertificatePassword (ConvertTo-SecureString $PnPPfxPassword -AsPlainText -Force)
} catch {
    Write-Error "PnP.PowerShell接続に失敗: $($_.Exception.Message)"
    Disconnect-ExchangeOnline -Confirm:$false
    exit 1
}

#========================
# Section 3: 監査ログ取得（UTC）
#========================
$sessionId  = [guid]::NewGuid().ToString()
$allRecords = @()
$max        = 5000

do {
    $batch = Search-UnifiedAuditLog `
        -StartDate $StartDateUtc `
        -EndDate   $EndDateUtc `
        -SessionId $sessionId `
        -SessionCommand ReturnLargeSet `
        -ResultSize $max

    # PS5.1 互換の null 安全カウント
    if ($null -eq $batch) { $batchCount = 0 } else { $batchCount = $batch.Count }

    if ($batchCount -gt 0) { $allRecords += $batch }
    Write-Host ("Fetched batch: {0}" -f $batchCount)
}
while ($batchCount -eq $max)

Write-Host ("Total records: {0}" -f ($allRecords.Count))

$stamp = (Get-Date -Format "yyyyMMdd_HHmmss")

# 生ログ（BOM付）
$auditCsvPath = Join-Path $env:TEMP "AuditLog_$stamp.csv"
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$csvText = $allRecords | ConvertTo-Csv -NoTypeInformation | Out-String
[System.IO.File]::WriteAllText($auditCsvPath, $csvText, $utf8Bom)

#========================
# Section 4: 拡張ログ生成（AuditData解析） - 方式A（安定版）
#========================
$parsed = @()

foreach ($rec in $allRecords) {
    if (-not $rec) { continue }

    # Operation（上位 → JSON フォールバック）
    $opValue = $null
    if ($rec.PSObject.Properties.Match('Operation').Count -gt 0) {
        $opValue = $rec.Operation
    }
    if (-not $opValue -and $rec.PSObject.Properties.Match('AuditData').Count -gt 0 -and $rec.AuditData) {
        try { $opValue = (ConvertFrom-Json $rec.AuditData -ErrorAction Stop).Operation } catch { $opValue = $null }
    }

    $normOp = Normalize-Operation -op $opValue
    if (-not $normOp) { continue }  # 対象4操作以外はスキップ

    # AuditData を展開（失敗しても例外にしない）
    $data = $null
    try { $data = $rec.AuditData | ConvertFrom-Json -ErrorAction Stop } catch { $data = $null }

    # ファイル情報（Parse-FileInfo は null/Hashtable/PSCustomObject/文字列に対応）
    $fileInfo = Parse-FileInfo -AuditData $data

# ...（前略）Operation 取得と $data, $fileInfo の生成まで同じ
# UserId（上位優先 → JSON → UserIds）※StrictMode対応
$uid = $null
if ($rec.PSObject.Properties.Match('UserId').Count -gt 0 -and $rec.UserId) {
    $uid = $rec.UserId
} elseif ($data) {
    $uid = Get-Prop $data 'UserId'
}
if (-not $uid -and $rec.PSObject.Properties.Match('UserIds').Count -gt 0 -and $rec.UserIds) {
    $uid = ($rec.UserIds -join ',')
}

# トップレベルの他プロパティも安全取得（推奨）
$creationDate = Get-TopProp $rec 'CreationDate'
$workload     = Get-TopProp $rec 'Workload'
$clientIP     = Get-TopProp $rec 'ClientIP'
$resultStatus = Get-TopProp $rec 'ResultStatus'
$recordType   = Get-TopProp $rec 'RecordType'

$parsed += [PSCustomObject]@{
    TimeGenerated  = $creationDate
    Workload       = $workload
    UserId         = $uid
    Operation      = $normOp
    RawOperation   = $opValue
    FileDirectory  = $fileInfo.FileDirectory
    FileName       = $fileInfo.FileName
    FileSizeBytes  = $fileInfo.FileSizeBytes
    ClientIP       = $clientIP
    ResultStatus   = $resultStatus
    RecordType     = $recordType
}

if ($null -eq $fileInfo) {
    $fileInfo = [PSCustomObject]@{
        FileDirectory  = ''
        FileName       = ''
        FileSizeBytes  = 0
    }
}

}

# 拡張ログCSV（BOM付）
$enrichedCsvPath = Join-Path $env:TEMP "AuditLog_Enriched_$stamp.csv"
if ($parsed.Count -gt 0) {
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$csvText = $parsed | ConvertTo-Csv -NoTypeInformation | Out-String
[System.IO.File]::WriteAllText($enrichedCsvPath, $csvText, $utf8Bom)
}

#========================
# Section 5: Teamsへアップロード（生ログ / 拡張ログ）
#========================
# 生ログ
try {
    Add-PnPFile -Path $auditCsvPath -Folder $FolderPath
} catch {
    Write-Warning "生ログのアップロードに失敗: $($_.Exception.Message)"
}

#拡張ログ
try {
        Add-PnPFile -Path $enrichedCsvPath -Folder $FolderPath
    } catch {
        Write-Warning "拡張ログのアップロードに失敗: $($_.Exception.Message)"
    }

#========================
# Section 6: ユーザー×操作の集計
#========================

$opSummary = @()

if ($parsed.Count -gt 0) {
    $opSummary = $parsed |
        Group-Object -Property UserId, Operation |
        ForEach-Object {
            $user  = $_.Name.Split(',')[0].Trim()
            $op    = $_.Name.Split(',')[1].Trim()
            $items = $_.Group

            $sumBytes = ($items | Measure-Object -Property FileSizeBytes -Sum).Sum
            if (-not $sumBytes) { $sumBytes = 0 }

            [PSCustomObject]@{
                UserId      = $user
                Operation   = $op
                Count       = $items.Count
                TotalSizeMB = [Math]::Round(($sumBytes / 1MB), 2)
            }
        } | Sort-Object UserId, Operation
}

# 集計CSVを保存＆アップロード

$opSummaryCsvPath = Join-Path $env:TEMP "OpSummary_$stamp.csv"
if ($opSummary.Count -gt 0) {
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$csvText = $opSummary | ConvertTo-Csv -NoTypeInformation | Out-String
[System.IO.File]::WriteAllText($opSummaryCsvPath, $csvText, $utf8Bom)
}

    # 保険で再接続
    try {
        Connect-PnPOnline -Url $SiteUrl -ClientId $PnPClientId -Tenant $PnPTenant -CertificatePath $PnPPfxPath -CertificatePassword (ConvertTo-SecureString $PnPPfxPassword -AsPlainText -Force)
    } catch {
        Write-Warning "PnP再接続に失敗: $($_.Exception.Message)"
    }
    try {
        Add-PnPFile -Path $opSummaryCsvPath -Folder $FolderPath
    } catch {
        Write-Warning "集計ファイルのアップロードに失敗: $($_.Exception.Message)"
    }



#========================
# Section 7: 閾値チェック & Teams通知
#========================
$exceeded = @()
if (@($opSummary).Count -gt 0) {
    $exceeded = @(
        $opSummary | Where-Object {
            $_.Count -ge $ThresholdCount -or $_.TotalSizeMB -ge $ThresholdSizeMB
        }
    )
}

if (@($exceeded).Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($WebhookUrl)) {
    $lines = $exceeded | ForEach-Object {
        "{0} ({1}) : 件数={2}, サイズ={3}MB" -f $_.UserId, $_.Operation, $_.Count, $_.TotalSizeMB
    }
    $messageDateLocal = $StartDateUtc.ToLocalTime()
    $message = "⚠️ 監査集計の閾値超過を検知しました（{0:yyyy-MM-dd}）`n{1}" -f $messageDateLocal, ($lines -join "`n")
  

# ===== Table（v1.5）で罫線付きの一覧を作成 =====
# 1) ヘッダー行を作成
$tableRows = @()
$tableRows += @{
  type  = "TableRow"
  cells = @(
    @{ type="TableCell"; items=@(@{ type="TextBlock"; text="UserId";   weight="Bolder"; wrap=$true }) },
    @{ type="TableCell"; items=@(@{ type="TextBlock"; text="Category"; weight="Bolder"; wrap=$true }) },
    @{ type="TableCell"; items=@(@{ type="TextBlock"; text="Count";    weight="Bolder" }) },
    @{ type="TableCell"; items=@(@{ type="TextBlock"; text="TotalMB";  weight="Bolder" }) }
  )
}

# 2) exceeded の各行を TableRow に変換
foreach ($item in $exceeded) {
  $tableRows += @{
    type  = "TableRow"
    cells = @(
      @{ type="TableCell"; items=@(@{ type="TextBlock"; text="$($item.UserId)";   wrap=$true }) },
      @{ type="TableCell"; items=@(@{ type="TextBlock"; text="$($item.Operation)"            }) },
      @{ type="TableCell"; items=@(@{ type="TextBlock"; text="$($item.Count)"                }) },
      @{ type="TableCell"; items=@(@{ type="TextBlock"; text="$([string]$item.TotalSizeMB)"  }) }
    )
  }
}

# 3) 罫線付き Table を作成
$table = @{
  type              = "Table"
  firstRowAsHeaders = $true
  showGridLines     = $true      # 罫線表示
  gridStyle         = "accent"   # 線色（"default"/"emphasis"/"accent" 等）
  columns           = @(
    @{ width = 3 }, @{ width = 2 }, @{ width = 1 }, @{ width = 1 }
  )
  rows = $tableRows
}

# 4) カード本体（version は 1.5）
$card = @{
  type    = "AdaptiveCard"
  version = "1.5"
  body    = @(
    @{ type="TextBlock"; text="OneDrive/SharePoint ファイル操作アラート（しきい値超過）"; weight="Bolder"; size="Medium" },
    @{ type="TextBlock"; text="期間: $($StartDateUtc.ToString('yyyy-MM-dd HH:mm')) ～ $($EndDateUtc.ToString('yyyy-MM-dd HH:mm')) UTC`n条件: 件数>=$ThresholdCount または 合計MB>=$ThresholdSizeMB"; wrap=$true },
    $table
  )
  # 任意: CSV へのリンクボタン（必要なら事前に $detailUrl / $aggUrl を用意）
  actions = @()

  msteams = @{
    width = "Full"   # フル幅指定（チャネルは広く、チャットは吹き出し最大幅に）
  }
}


# 5) Workflows Webhook へ送るペイロード（既存の形式を踏襲）
$payload = @{
  type = "message"
  attachments = @(@{
    contentType = "application/vnd.microsoft.card.adaptive"
    content     = $card
  })
} | ConvertTo-Json -Depth 50 -Compress

try {
  Invoke-RestMethod -Uri $WebhookUrl -Method Post -ContentType 'application/json' -Body $payload | Out-Null
} catch {
  Write-Warning "Webhook通知に失敗: $($_.Exception.Message)"
}

}

#========================
# Section 8: 後処理
#========================
Remove-Item $auditCsvPath, $enrichedCsvPath, $opSummaryCsvPath -ErrorAction SilentlyContinue

Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Host "完了: $stamp"
