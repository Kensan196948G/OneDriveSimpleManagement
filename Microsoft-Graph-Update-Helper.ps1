# Microsoft Graphモジュール更新＆権限サポートスクリプト
# 作成日: 2025/3/13
# 
# このスクリプトはOneDriveスクリプト実行前の準備として以下を行います：
# 1. Microsoft Graphモジュールの最新版へのアップデート
# 2. 必要なAPI権限の確認
# 3. 管理者権限での実行手順の表示

#region 管理者権限の確認
# 管理者権限で実行されているか確認
function Test-AdminRights {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

# 管理者権限で実行されていない場合は警告を表示
if (-not (Test-AdminRights)) {
    Write-Host "警告: このスクリプトは管理者権限で実行されていません。" -ForegroundColor Yellow
    Write-Host "一部の操作が失敗する可能性があります。" -ForegroundColor Yellow
    Write-Host
    Write-Host "管理者権限で実行するには、以下の手順を実行してください:" -ForegroundColor Cyan
    Write-Host "1. Windowsメニューから「PowerShell」または「Windows PowerShell」を右クリック"
    Write-Host "2. 「管理者として実行」を選択"
    Write-Host "3. 表示されたPowerShellウィンドウで次のコマンドを実行:"
    Write-Host "   cd '$PWD'" -ForegroundColor Green
    Write-Host "   & '.\Microsoft-Graph-Update-Helper.ps1'" -ForegroundColor Green
    Write-Host
    $continue = Read-Host "続行しますか？ (Y/N)"
    if ($continue -ne "Y") {
        Write-Host "スクリプトを終了します。"
        exit
    }
    Write-Host "一般ユーザー権限で続行します。一部の機能が制限される場合があります。" -ForegroundColor Yellow
    Write-Host
}
else {
    Write-Host "管理者権限で実行中です - OK" -ForegroundColor Green
    Write-Host
}
#endregion

#region Microsoft Graph モジュール更新
Write-Host "====== Microsoft Graph モジュール更新 ======" -ForegroundColor Cyan

# 現在のGraph モジュールバージョンを確認
$currentModule = Get-Module -Name Microsoft.Graph -ListAvailable
if ($currentModule) {
    Write-Host "現在のMicrosoft.Graphモジュールバージョン: " -NoNewline
    Write-Host "$($currentModule.Version)" -ForegroundColor Yellow
}
else {
    Write-Host "Microsoft.Graphモジュールがインストールされていません" -ForegroundColor Red
}

Write-Host "Microsoft.Graph モジュールを最新バージョンにアップデートしています..." -ForegroundColor Cyan
Write-Host "(これには数分かかる場合があります)" -ForegroundColor Gray

try {
    # NuGetプロバイダーの確認とインストール
    $nugetProvider = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
    if (-not $nugetProvider -or $nugetProvider.Version -lt [Version]::new(2, 8, 5, 201)) {
        Write-Host "NuGetプロバイダーをインストールしています..." -ForegroundColor Yellow
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser
    }

    # PSGalleryが信頼済みリポジトリとして設定されているか確認
    $psGallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($psGallery -and $psGallery.InstallationPolicy -ne "Trusted") {
        Write-Host "PSGalleryを信頼済みリポジトリとして設定しています..." -ForegroundColor Yellow
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    }

    # まず既存のモジュールをアンインストール（クリーンインストール）
    if ($currentModule) {
        Write-Host "既存のMicrosoft.Graphモジュールをアンインストールしています..." -ForegroundColor Yellow
        Uninstall-Module -Name Microsoft.Graph -AllVersions -Force -ErrorAction SilentlyContinue
        # 個別のサブモジュールも削除
        Get-Module -Name Microsoft.Graph.* -ListAvailable | ForEach-Object {
            Write-Host "サブモジュール $($_.Name) をアンインストールしています..." -ForegroundColor Yellow
            Uninstall-Module -Name $_.Name -AllVersions -Force -ErrorAction SilentlyContinue
        }
    }

    # 新しいモジュールをインストール
    Write-Host "Microsoft.Graphモジュールの最新バージョンをインストールしています..." -ForegroundColor Yellow
    Install-Module -Name Microsoft.Graph -Force -AllowClobber -Scope CurrentUser -SkipPublisherCheck

    # インストール後のバージョンを確認
    $updatedModule = Get-Module -Name Microsoft.Graph -ListAvailable
    Write-Host "アップデート後のMicrosoft.Graphモジュールバージョン: " -NoNewline
    Write-Host "$($updatedModule.Version)" -ForegroundColor Green
    Write-Host "Microsoft.Graphモジュールが正常に更新されました" -ForegroundColor Green
}
catch {
    Write-Host "エラーが発生しました: $_" -ForegroundColor Red
    Write-Host "代替方法として、管理者権限でPowerShellを開き、以下のコマンドを実行してください:" -ForegroundColor Yellow
    Write-Host "Install-Module -Name Microsoft.Graph -Force -AllowClobber -Scope CurrentUser" -ForegroundColor Green
}
#endregion

#region APIパーミッション設定
Write-Host "====== API パーミッション設定ヘルパー ======" -ForegroundColor Cyan
Write-Host "OneDriveスクリプトの実行には、Microsoft Graph APIの適切なパーミッションが必要です。" -ForegroundColor White

# 必要なパーミッションを表示
Write-Host "必要なパーミッション:" -ForegroundColor Yellow
Write-Host "- User.Read.All (ユーザー情報の読み取り)" -ForegroundColor White
Write-Host "- Directory.Read.All (ディレクトリ情報の読み取り)" -ForegroundColor White
Write-Host "- Sites.Read.All (サイト情報の読み取り - OneDrive基本情報取得用)" -ForegroundColor White
Write-Host "- Sites.ReadWrite.All (サイト情報の読み書き - OneDrive詳細情報取得用)" -ForegroundColor White

Write-Host
Write-Host "グローバル管理者への承認依頼手順:" -ForegroundColor Cyan
Write-Host "1. OneDriveスクリプト実行時に表示されるメッセージに従い、承認ページを開きます"
Write-Host "   ※承認ページのURLの例: https://login.microsoftonline.com/[テナントID]/adminconsent?client_id=[クライアントID]"
Write-Host "2. グローバル管理者アカウントでログインします"
Write-Host "3. 「組織の代理として同意する」ボタンをクリックします"
Write-Host "4. 承認が完了したら、スクリプトを再度実行します"
Write-Host

# テナントIDの設定方法を説明
Write-Host "テナントIDの設定方法:" -ForegroundColor Cyan
Write-Host "1. OneDriveスクリプト内のテナントID設定を実際の値に変更します"
Write-Host "   変更箇所: `$tenantId = `"your-tenant-id`"  <- ここを実際のテナントIDに変更"
Write-Host "   テナントIDはAzure管理ポータルで確認できます"
Write-Host
Write-Host "2. テナントIDが不明な場合、以下のコマンドで現在のテナントIDを取得できます:"
Write-Host "   Connect-MgGraph -Scopes User.Read; (Get-MgContext).TenantId"
Write-Host

# スコープ確認用関数
function Test-GraphPermissions {
    param (
        [Parameter(Mandatory = $false)]
        [string[]]$RequiredScopes = @("User.Read.All", "Directory.Read.All", "Sites.Read.All")
    )

    try {
        # Microsoft Graphへ接続
        Connect-MgGraph -Scopes $RequiredScopes -ErrorAction Stop
        
        # コンテキストの取得
        $context = Get-MgContext
        
        if ($context) {
            Write-Host "Microsoft Graph APIに接続しました" -ForegroundColor Green
            Write-Host "現在のスコープ:" -ForegroundColor Cyan
            foreach ($scope in $context.Scopes) {
                Write-Host "- $scope" -ForegroundColor White
            }
            
            # 必要なスコープがすべて含まれているか確認
            $missingScopes = @()
            foreach ($requiredScope in $RequiredScopes) {
                if (-not $context.Scopes.Contains($requiredScope)) {
                    $missingScopes += $requiredScope
                }
            }
            
            if ($missingScopes.Count -gt 0) {
                Write-Host "警告: 以下の必要なスコープが不足しています:" -ForegroundColor Yellow
                foreach ($missingScope in $missingScopes) {
                    Write-Host "- $missingScope" -ForegroundColor Yellow
                }
                Write-Host "管理者承認が必要な可能性があります。上記の「グローバル管理者への承認依頼手順」を参照してください。" -ForegroundColor Yellow
            }
            else {
                Write-Host "すべての必要なパーミッションがあります！スクリプトの実行準備が整いました。" -ForegroundColor Green
            }
        }
        else {
            Write-Host "Microsoft Graph APIへの接続に失敗しました" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "エラーが発生しました: $_" -ForegroundColor Red
    }
}

# パーミッションテスト実行（オプション）
$runPermTest = Read-Host "Microsoft Graph APIのパーミッションをテストしますか？ (Y/N)"
if ($runPermTest -eq "Y") {
    Test-GraphPermissions
}
#endregion

#region OneDriveスクリプト実行手順
Write-Host "====== OneDriveスクリプト実行手順 ======" -ForegroundColor Cyan
Write-Host "OneDriveスクリプトを管理者権限で実行するには、以下の手順に従ってください:" -ForegroundColor White

Write-Host "方法1: PowerShell管理者モードで実行" -ForegroundColor Yellow
Write-Host "1. Windowsメニューから「PowerShell」または「Windows PowerShell」を右クリック"
Write-Host "2. 「管理者として実行」を選択"
Write-Host "3. 表示されたPowerShellウィンドウで次のコマンドを実行:"
Write-Host "   cd '$PWD'" -ForegroundColor Green
Write-Host "   & '.\OneDriveSyncManagement.ps1'" -ForegroundColor Green
Write-Host

Write-Host "方法2: 現在のPowerShellセッションから管理者として実行" -ForegroundColor Yellow
Write-Host "以下のコマンドを実行すると、管理者権限の新しいウィンドウでスクリプトが起動します:"
Write-Host "   Start-Process PowerShell -Verb RunAs -ArgumentList `"-ExecutionPolicy Bypass -File `"`"$PWD\OneDriveSyncManagement.ps1`"`"`"" -ForegroundColor Green
Write-Host

Write-Host "OneDriveスクリプトで推奨する実行オプション:" -ForegroundColor Yellow
Write-Host "- 実行ポリシーのバイパス: -ExecutionPolicy Bypass"
Write-Host "- 詳細ログ出力: -Verbose"
Write-Host "- 例: " -NoNewline
Write-Host "& '.\OneDriveSyncManagement.ps1' -Verbose" -ForegroundColor Green
Write-Host

# 便利なエイリアスを作成（オプション）
$createAlias = Read-Host "OneDriveスクリプト実行用のエイリアスを作成しますか？ (Y/N)"
if ($createAlias -eq "Y") {
    $psProfilePath = $PROFILE
    $aliasContent = @"

# OneDriveスクリプト実行用エイリアス - $(Get-Date -Format "yyyy/MM/dd")追加
function Run-OneDriveSync { & '$PWD\OneDriveSyncManagement.ps1' @args }
New-Alias -Name onedrive -Value Run-OneDriveSync
"@

    try {
        # プロファイルディレクトリが存在しない場合は作成
        $profileDir = Split-Path -Path $psProfilePath -Parent
        if (-not (Test-Path -Path $profileDir)) {
            New-Item -Path $profileDir -ItemType Directory -Force | Out-Null
        }

        # プロファイルファイルが存在しない場合は作成
        if (-not (Test-Path -Path $psProfilePath)) {
            New-Item -Path $psProfilePath -ItemType File -Force | Out-Null
        }

        # エイリアスを追加
        Add-Content -Path $psProfilePath -Value $aliasContent
        Write-Host "エイリアスが追加されました。新しいPowerShellウィンドウで 'onedrive' コマンドでスクリプトを実行できます。" -ForegroundColor Green
    }
    catch {
        Write-Host "エイリアス作成中にエラーが発生しました: $_" -ForegroundColor Red
    }
}
#endregion

Write-Host "===============================" -ForegroundColor Cyan
Write-Host "準備作業が完了しました！" -ForegroundColor Green
Write-Host "OneDriveスクリプトを実行する前に、上記の手順に従って権限設定と管理者実行を行ってください。" -ForegroundColor White
Write-Host "===============================" -ForegroundColor Cyan

# スクリプト終了
Write-Host "Enterキーを押すと終了します..."
Read-Host
