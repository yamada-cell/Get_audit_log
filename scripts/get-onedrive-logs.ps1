# 認証情報の取得
$clientId = $env:CLIENT_ID
$clientSecret = $env:CLIENT_SECRET
$tenantId = $env:TENANT_ID
$webhookUrl = $env:TEAMS_WEBHOOK_URL

# トークン取得
$tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}
$accessToken = $tokenResponse.access_token
$headers = @{ Authorization = "Bearer $accessToken" }

# 監査ログ取得（FileDownloaded / FileUploaded）
# ※ 実際には Office 365 Management Activity API を使う必要があります
# ここでは仮の構造で進めます
$startDate = (Get-Date).AddHours(-1).ToString("yyyy-MM-ddTHH:mm:ssZ")
$endDate = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")

# 仮のログデータ取得（実際は Graph API または Management API を使用）
$response = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/auditLogs/signIns" -Headers $headers

# ユーザーごとの集計
$userStats = @{}

foreach ($log in $response.value) {
    if ($log.activityDisplayName -match "FileDownloaded|FileUploaded") {
        $user = $log.userPrincipalName
        $fileSizeMB = [math]::Round(($log.resourceSize / 1MB), 2)

        if (-not $userStats.ContainsKey($user)) {
            $userStats[$user] = @{
                FileCount = 0
                TotalSizeMB = 0
            }
        }

        $userStats[$user].FileCount++
        $userStats[$user].TotalSizeMB += $fileSizeMB
    }
}

# 閾値設定
$thresholdCount = 100
$thresholdSizeMB = 100

# Teams通知（閾値超過ユーザーのみ）
foreach ($user in $userStats.Keys) {
    $stats = $userStats[$user]
    if ($stats.FileCount -gt $thresholdCount -or $stats.TotalSizeMB -gt $thresholdSizeMB) {
        $message = @{
            "@type" = "MessageCard"
            "@context" = "http://schema.org/extensions"
            "summary" = "OneDriveファイル操作アラート"
            "themeColor" = "FF0000"
            "title" = "OneDriveファイル操作アラート"
            "text" = "ユーザー：$user`nファイル数：$($stats.FileCount) 件`n合計サイズ：$($stats.TotalSizeMB) MB"
        }
        Invoke-RestMethod -Uri $webhookUrl -Method Post -Body (ConvertTo-Json $message -Depth 3) -ContentType 'application/json'
    }
}

# CSV出力（オプション）
$userStats.GetEnumerator() | ForEach-Object {
    [PSCustomObject]@{
        User          = $_.Key
        FileCount     = $_.Value.FileCount
        TotalSizeMB   = $_.Value.TotalSizeMB
    }
} | Export-Csv -Path "onedrive-user-log.csv" -NoTypeInformation
