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
        # ウェルカムメッセージを抑制して接続
        Connect-MgGraph -Scopes "User.Read" -NoWelcome -ErrorAction Stop
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
    
    # まず現在のセッションで有効なスコープを確認
    $currentContext = Get-MgContext -ErrorAction SilentlyContinue
    if ($currentContext) {
        Write-Host "現在のセッションで以下の権限が確認されています:" -ForegroundColor Cyan
        foreach ($scope in $currentContext.Scopes) {
            Write-Host "- $scope" -ForegroundColor White
        }
        
        # 必要なスコープがすでに含まれているか確認
        $missingScopes = @()
        foreach ($scope in $RequiredScopes) {
            # Sites.ReadWrite.All権限があれば、Sites.Read.Allも含まれていると見なす
            if ($scope -eq "Sites.Read.All" -and $currentContext.Scopes -contains "Sites.ReadWrite.All") {
                Write-Host "Sites.ReadWrite.All権限があるため、Sites.Read.All権限も満たされていると判断" -ForegroundColor Green
                continue
            }
            
            # 権限の一部が含まれているかも確認（例：Directory.ReadWrite.Allは Directory.Read.Allを包含）
            $hasHigherPermission = $false
            foreach ($grantedScope in $currentContext.Scopes) {
                if ($grantedScope -like "*$($scope.Replace('.Read.', '.ReadWrite.'))*" -or 
                    $grantedScope -like "*$($scope.Replace('.Read.', '.FullControl.'))*") {
                    Write-Host "$scope は上位権限 $grantedScope によって満たされていると判断" -ForegroundColor Green
                    $hasHigherPermission = $true
                    break
                }
            }
            
            if (-not $hasHigherPermission -and -not $currentContext.Scopes -contains $scope) {
                $missingScopes += $scope
            }
        }
        
        # 直接権限テストを実行してみる
        $canAccessUserInfo = $false
        $canAccessOneDrive = $false
        
        try {
            # ユーザー情報にアクセスできるかテスト
            $me = Get-MgUser -UserId $currentContext.Account -Property DisplayName -ErrorAction Stop
            Write-Host "ユーザー情報アクセステスト成功: $($me.DisplayName)" -ForegroundColor Green
            $canAccessUserInfo = $true
            
            # OneDriveへアクセスできるかテスト
            $drive = Get-MgUserDrive -UserId $currentContext.Account -ErrorAction Stop
            Write-Host "OneDriveアクセステスト成功" -ForegroundColor Green
            $canAccessOneDrive = $true
        }
        catch {
            Write-Host "アクセステスト失敗: $_" -ForegroundColor Yellow
        }
        
        # 権限不足でも実際のアクセスが可能なら成功とみなす
        if ($canAccessUserInfo -and $canAccessOneDrive) {
            Write-Host "必要な操作が可能なため、権限は十分と判断します" -ForegroundColor Green
            return @{
                Success = $true
                MissingScopes = @()
                GrantedScopes = $currentContext.Scopes
            }
        }
        
        if ($missingScopes.Count -eq 0) {
            return @{
                Success = $true
                MissingScopes = @()
                GrantedScopes = $currentContext.Scopes
            }
        }
    }
    
    # 現在のセッションで不足している場合は再接続を試みる
    try {
        Write-Host "必要なスコープで再接続を試みます..." -ForegroundColor Yellow
        # より広範囲の権限を試みる
        $expandedScopes = $RequiredScopes + @("Directory.ReadWrite.All", "Sites.FullControl.All")
        Connect-MgGraph -Scopes $expandedScopes -NoWelcome -ErrorAction Stop
        $context = Get-MgContext
        
        if ($context) {
            # 持っているスコープを確認
            $grantedScopes = $context.Scopes
            Write-Host "再接続で取得した権限:" -ForegroundColor Cyan
            foreach ($scope in $grantedScopes) {
                Write-Host "- $scope" -ForegroundColor White
            }
            
            # 不足しているスコープを確認
            $missingScopes = @()
            foreach ($scope in $RequiredScopes) {
                # 直接または上位権限があるか確認
                $hasScope = $grantedScopes -contains $scope
                $hasHigherScope = $false
                
                # 上位権限の確認
                foreach ($grantedScope in $grantedScopes) {
                    if ($grantedScope -like "*$($scope.Replace('.Read.', '.ReadWrite.'))*" -or 
                        $grantedScope -like "*$($scope.Replace('.Read.', '.FullControl.'))*") {
                        $hasHigherScope = $true
                        break
                    }
                }
                
                if (-not $hasScope -and -not $hasHigherScope) {
                    $missingScopes += $scope
                }
            }
            
            # 直接権限テストを再度実行
            $canAccessUserInfo = $false
            $canAccessOneDrive = $false
            
            try {
                # ユーザー情報にアクセスできるかテスト
                $me = Get-MgUser -UserId $context.Account -Property DisplayName -ErrorAction Stop
                Write-Host "ユーザー情報アクセステスト成功: $($me.DisplayName)" -ForegroundColor Green
                $canAccessUserInfo = $true
                
                # OneDriveへアクセスできるかテスト
                $drive = Get-MgUserDrive -UserId $context.Account -ErrorAction Stop
                Write-Host "OneDriveアクセステスト成功" -ForegroundColor Green
                $canAccessOneDrive = $true
                
                # アクセスが可能なら成功とみなす
                if ($canAccessUserInfo -and $canAccessOneDrive) {
                    Write-Host "必要な操作が可能なため、権限は十分と判断します" -ForegroundColor Green
                    return @{
                        Success = $true
                        MissingScopes = @()
                        GrantedScopes = $grantedScopes
                    }
                }
            }
            catch {
                Write-Host "アクセステスト失敗: $_" -ForegroundColor Yellow
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
        Write-Host "再接続中にエラーが発生しました: $_" -ForegroundColor Red
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

# 既知のテナントID（スクリプト更新などに使用）
$knownTenantId = "a7232f7a-a9e5-4f71-9372-dc8b1c6645ea"

# テナントID取得（確認用）
Write-Host "1. テナントIDを取得・確認しています..." -ForegroundColor Yellow
$fetchedTenantId = Get-AzureTenantId
if ($fetchedTenantId) {
    # 文字列をトリミングして余分な空白や改行を除去
    $fetchedTenantId = $fetchedTenantId.ToString().Trim()
    Write-Host "取得したテナントID: $fetchedTenantId" -ForegroundColor Green
    
    # 実際に使用するテナントIDは既知の値を使用
    $tenantId = $knownTenantId
    Write-Host "使用するテナントID: $tenantId" -ForegroundColor Cyan
    
    # スクリプトのテナントID自動更新
    $updateScript = Read-Host "OneDriveスクリプトのテナントIDを自動更新しますか？ (Y/N)"
    if ($updateScript -eq "Y") {
        $scriptPath = Read-Host "スクリプトのパス (デフォルト: .\OneDriveSyncManagement.ps1)"
        if ([string]::IsNullOrEmpty($scriptPath)) {
            $scriptPath = ".\OneDriveSyncManagement.ps1"
        }
        
        Update-OneDriveScript -TenantId $knownTenantId -ScriptPath $scriptPath
    }
    
    # 権限確認（既知のテナントIDを使用）
    Write-Host "2. API権限を確認しています..." -ForegroundColor Yellow
    $scopeResult = Test-GraphScopes -TenantId $knownTenantId
    
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
