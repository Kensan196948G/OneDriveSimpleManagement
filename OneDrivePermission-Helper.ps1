# Microsoft Graph APIパーミッション取得スクリプト
# 作成日: 2025/3/13
# 
# このスクリプトはOneDriveスクリプト用のAPI権限を取得します：
# 1. 必要なAPI権限の確認と取得
# 2. テナントID自動検出
# 3. グローバル管理者承認プロセスのサポート

# モジュールの確認とインストール
Write-Host "Microsoft Graph APIパーミッション取得サポートを開始します..." -ForegroundColor Cyan
Write-Host

# Microsoft Graphモジュールの確認
$requiredModules = @("Microsoft.Graph")
$modulesToInstall = @()

foreach ($module in $requiredModules) {
    if (-not (Get-Module -Name $module -ListAvailable)) {
        $modulesToInstall += $module
    }
}

if ($modulesToInstall.Count -gt 0) {
    Write-Host "必要なモジュール ($($modulesToInstall -join ', ')) をインストールします..." -ForegroundColor Yellow
    try {
        foreach ($module in $modulesToInstall) {
            Install-Module -Name $module -Force -Scope CurrentUser
        }
        Write-Host "必要なモジュールのインストールが完了しました" -ForegroundColor Green
    }
    catch {
        Write-Host "モジュールのインストール中にエラーが発生しました: $_" -ForegroundColor Red
        Write-Host "別途、Microsoft-Graph-Update-Helper.ps1を実行してモジュールをインストールしてください" -ForegroundColor Yellow
    }
}
else {
    Write-Host "必要なモジュールはすでにインストールされています" -ForegroundColor Green
}

# OneDriveスクリプトのクライアントID
$clientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e" # Graph PowerShellのClientID

# テナントID取得
function Get-AzureTenantId {
    try {
        Connect-MgGraph -Scopes "User.Read" -ErrorAction Stop
        $context = Get-MgContext
        if ($context -and $context.TenantId) {
            return $context.TenantId
        }
    }
    catch {
        Write-Host "テナントID取得中にエラーが発生しました: $_" -ForegroundColor Red
    }
    return $null
}

# 必要なスコープの確認
function Test-GraphScopes {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        
        [Parameter(Mandatory = $false)]
        [string[]]$RequiredScopes = @("User.Read.All", "Directory.Read.All", "Sites.Read.All", "Sites.ReadWrite.All")
    )
    
    try {
        # スコープを指定して接続を試みる
        Connect-MgGraph -Scopes $RequiredScopes -ErrorAction Stop
        $context = Get-MgContext
        
        if ($context) {
            # 持っているスコープを確認
            $grantedScopes = $context.Scopes
            
            # 不足しているスコープを確認
            $missingScopes = @()
            foreach ($scope in $RequiredScopes) {
                if (-not $grantedScopes.Contains($scope)) {
                    $missingScopes += $scope
                }
            }
            
            if ($missingScopes.Count -gt 0) {
                return @{
                    Success = $false
                    MissingScopes = $missingScopes
                    GrantedScopes = $grantedScopes
                }
            }
            else {
                return @{
                    Success = $true
                    MissingScopes = @()
                    GrantedScopes = $grantedScopes
                }
            }
        }
    }
    catch {
        return @{
            Success = $false
            Error = $_
            MissingScopes = $RequiredScopes
            GrantedScopes = @()
        }
    }
}

# 管理者同意URL生成
function Get-AdminConsentUrl {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        
        [Parameter(Mandatory = $true)]
        [string]$ClientId
    )
    
    return "https://login.microsoftonline.com/$TenantId/adminconsent?client_id=$ClientId"
}

# ユーザー情報取得テスト
function Test-UserInfo {
    try {
        $currentUser = Get-MgUser -UserId (Get-MgContext).Account -Property DisplayName,Mail,UserType -ErrorAction Stop
        if ($currentUser) {
            Write-Host "ユーザー情報取得成功:" -ForegroundColor Green
            Write-Host "- 名前: $($currentUser.DisplayName)"
            Write-Host "- メール: $($currentUser.Mail)"
            Write-Host "- 種別: $($currentUser.UserType)"
            return $true
        }
    }
    catch {
        Write-Host "ユーザー情報取得中にエラーが発生しました: $_" -ForegroundColor Red
    }
    return $false
}

# OneDrive情報取得テスト
function Test-OneDriveInfo {
    try {
        $drive = Get-MgUserDrive -UserId (Get-MgContext).Account -ErrorAction Stop
        if ($drive) {
            $totalGB = [math]::Round($drive.Quota.Total / 1GB, 2)
            $usedGB = [math]::Round($drive.Quota.Used / 1GB, 2)
            $usagePercent = [math]::Round(($drive.Quota.Used / $drive.Quota.Total) * 100, 2)
            
            Write-Host "OneDrive情報取得成功:" -ForegroundColor Green
            Write-Host "- 総容量: $totalGB GB"
            Write-Host "- 使用量: $usedGB GB"
            Write-Host "- 使用率: $usagePercent %"
            return $true
        }
    }
    catch {
        Write-Host "OneDrive情報取得中にエラーが発生しました: $_" -ForegroundColor Red
    }
    return $false
}

# OneDriveスクリプトの更新
function Update-OneDriveScript {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        
        [Parameter(Mandatory = $false)]
        [string]$ScriptPath = ".\OneDriveSyncManagement.ps1"
    )
    
    if (-not (Test-Path $ScriptPath)) {
        Write-Host "スクリプトファイルが見つかりません: $ScriptPath" -ForegroundColor Red
        return $false
    }
    
    try {
        $scriptContent = Get-Content $ScriptPath -Raw
        
        # テナントIDを更新
        $updatedContent = $scriptContent -replace '(\$tenantId\s*=\s*)["'']your-tenant-id["'']', "`$1`"$TenantId`""
        
        # 更新内容をファイルに書き込み
        Set-Content -Path $ScriptPath -Value $updatedContent
        
        Write-Host "スクリプトファイルのテナントIDを更新しました: $ScriptPath" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "スクリプトファイル更新中にエラーが発生しました: $_" -ForegroundColor Red
        return $false
    }
}

# メイン処理
Write-Host "====== OneDrive API権限取得プロセス ======" -ForegroundColor Cyan

# テナントID取得
Write-Host "1. テナントIDを取得しています..." -ForegroundColor Yellow
$tenantId = Get-AzureTenantId
if ($tenantId) {
    Write-Host "テナントID: $tenantId" -ForegroundColor Green
    
    # スクリプトのテナントID自動更新
    $updateScript = Read-Host "OneDriveスクリプトのテナントIDを自動更新しますか？ (Y/N)"
    if ($updateScript -eq "Y") {
        $scriptPath = Read-Host "スクリプトのパス (デフォルト: .\OneDriveSyncManagement.ps1)"
        if ([string]::IsNullOrEmpty($scriptPath)) {
            $scriptPath = ".\OneDriveSyncManagement.ps1"
        }
        
        Update-OneDriveScript -TenantId $tenantId -ScriptPath $scriptPath
    }
    
    # 権限確認
    Write-Host "2. API権限を確認しています..." -ForegroundColor Yellow
    $scopeResult = Test-GraphScopes -TenantId $tenantId
    
    if ($scopeResult.Success) {
        Write-Host "すべての必要な権限が付与されています！" -ForegroundColor Green
        
        # ユーザー情報取得テスト
        Write-Host "3. ユーザー情報取得テスト..." -ForegroundColor Yellow
        $userTest = Test-UserInfo
        
        # OneDrive情報取得テスト
        Write-Host "4. OneDrive情報取得テスト..." -ForegroundColor Yellow
        $oneDriveTest = Test-OneDriveInfo
        
        if ($userTest -and $oneDriveTest) {
            Write-Host "すべてのテストが成功しました！OneDriveスクリプトを実行する準備ができています。" -ForegroundColor Green
        }
        else {
            Write-Host "一部のテストに失敗しました。グローバル管理者の承認が必要かもしれません。" -ForegroundColor Yellow
            $adminConsentUrl = Get-AdminConsentUrl -TenantId $tenantId -ClientId $clientId
            Write-Host "以下のURLをグローバル管理者に共有して承認を依頼してください:" -ForegroundColor Yellow
            Write-Host $adminConsentUrl -ForegroundColor Cyan
        }
    }
    else {
        Write-Host "一部の権限が不足しています。グローバル管理者の承認が必要です。" -ForegroundColor Yellow
        
        if ($scopeResult.MissingScopes.Count -gt 0) {
            Write-Host "不足している権限:" -ForegroundColor Yellow
            foreach ($scope in $scopeResult.MissingScopes) {
                Write-Host "- $scope" -ForegroundColor White
            }
        }
        
        $adminConsentUrl = Get-AdminConsentUrl -TenantId $tenantId -ClientId $clientId
        Write-Host "以下のURLをグローバル管理者に共有して承認を依頼してください:" -ForegroundColor Yellow
        Write-Host $adminConsentUrl -ForegroundColor Cyan
        
        # ブラウザで開くオプション
        $openBrowser = Read-Host "ブラウザでこのURLを開きますか？ (Y/N)"
        if ($openBrowser -eq "Y") {
            Start-Process $adminConsentUrl
        }
    }
}
else {
    Write-Host "テナントIDの取得に失敗しました。ログインして再度お試しください。" -ForegroundColor Red
}

Write-Host "処理が完了しました。Enterキーを押して終了..." -ForegroundColor Cyan
Read-Host
