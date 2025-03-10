# OneDrive 同期状況管理スクリプト
# 説明: Microsoft 365/Office 365のユーザーのOneDrive利用状況を収集し、レポートとして出力します
# 対応レポート形式: CSV, HTML（インタラクティブUIつき）, テキストログ
# 本スクリプトは次の3つのスクリプトを統合したものです:
# - miraiAllUserInfoComplete_final.ps1
# - miraiAllUserInfoComplete.ps1
# - miraiAllUserInfoComplete_updated.ps1
# バージョン: 1.0.0

#region パラメータ設定
param (
    [string]$OutputDir = "$(Get-Location)"
)

# Microsoft Graph の API パーミッション設定
$tenantId = "your-tenant-id"
$clientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"  # Graph PowerShellのClientID
$requiredScopes = @("User.Read.All", "Directory.Read.All", "Sites.Read.All")

# 実行開始時刻を記録
$executionTime = Get-Date
Write-Output "実行開始時刻: $($executionTime.ToString('yyyy/MM/dd HH:mm:ss'))"

# 日付ベースのフォルダ名を作成 (OneDriveCheck.YYYYMMDD)
$dateFolderName = "OneDriveCheck." + $executionTime.ToString("yyyyMMdd")
$dateFolderPath = Join-Path -Path $OutputDir -ChildPath $dateFolderName

# 出力ディレクトリが存在しない場合は作成
if (-not (Test-Path -Path $dateFolderPath)) {
    New-Item -Path $dateFolderPath -ItemType Directory | Out-Null
    Write-Output "出力用フォルダを作成しました: $dateFolderPath"
}
#endregion

#region Microsoft Graph接続
# Microsoft Graphモジュールのインストール確認と実施
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Host "Microsoft.Graph モジュールが未インストールのためインストールします..."
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}

# Microsoft Graph に接続（基本情報取得のための最小スコープ）
try {
    Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All"
    Write-Output "Microsoft Graph APIに接続しました（初期スコープ: User.Read.All, Directory.Read.All）"
} 
catch {
    Write-Error "Microsoft Graph APIへの接続に失敗しました: $_"
    Exit 1
}

# ログインユーザーのUPN（メールアドレス）を自動取得
$context = Get-MgContext
if (-not $context) {
    Write-Error "Microsoft Graph接続のコンテキストが取得できませんでした。"
    Exit 1
}
$UserUPN = $context.Account
Write-Output "ログインユーザー: $UserUPN"

# ログイン済ユーザー情報を取得
$currentUser = Get-MgUser -UserId $UserUPN -Property DisplayName,Mail,onPremisesSamAccountName,AccountEnabled,onPremisesLastSyncDateTime,UserType
if (-not $currentUser) {
    Write-Error "ログインユーザーの情報が取得できませんでした。"
    Exit 1
}
Write-Output "ユーザー情報取得成功: $($currentUser.DisplayName)"
#endregion

#region ユーザー種別の判定
# ユーザーの役割を取得（グローバル管理者かどうかを確認）
$roles = Get-MgDirectoryRole -All
$globalAdminRole = $roles | Where-Object { $_.DisplayName -eq "Global Administrator" -or $_.DisplayName -eq "Company Administrator" }
$isGlobalAdmin = $false
$roleMembers = @() # 初期化して常に配列として扱えるようにする

if ($globalAdminRole) {
    Write-Output "グローバル管理者ロール検索結果: $($globalAdminRole.DisplayName)"
    Write-Output "グローバル管理者ロールID: $($globalAdminRole.Id)"
    $roleMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $globalAdminRole.Id
    Write-Output "グローバル管理者ロールメンバー取得完了: $($roleMembers.Count) 件"
    $isGlobalAdmin = $roleMembers.Id -contains $currentUser.Id
    Write-Output "現在のユーザーがグローバル管理者か: $isGlobalAdmin"
}

# ユーザー種別の判定
if ($currentUser.UserType -eq "Guest") {
    $userType = "Guest"
    Write-Output "ユーザー種別: ゲスト"
} elseif ($isGlobalAdmin) {
    $userType = "Administrator"
    Write-Output "ユーザー種別: グローバル管理者"
} else {
    $userType = "Member"
    Write-Output "ユーザー種別: 一般メンバー"
}

# 一般ユーザーなら管理者の承認をリクエスト
if ($userType -eq "Member") {
    Write-Host "一般ユーザーのため、API パーミッションの付与が必要です。"
    $adminConsent = Read-Host "グローバル管理者の承認を得るためのページを開きますか？ (Y/N)"
    if ($adminConsent -eq 'Y') {
        Start-Process "https://login.microsoftonline.com/$tenantId/adminconsent?client_id=$clientId"
        Write-Host "グローバル管理者の承認を得てから、もう一度スクリプトを実行してください。"
        Exit
    }
}

# グローバル管理者かどうかを判定（権限のチェック）
$isAdmin = ($context.Scopes -contains "Directory.ReadWrite.All") # Directory.ReadWrite.All権限があるかどうか
#endregion

#region 追加スコープで再接続
# グローバル管理者または承認済みユーザーなら、Microsoft Graph にフル接続
try {
    Connect-MgGraph -Scopes $requiredScopes
    Write-Output "Microsoft Graph APIに再接続しました（フルスコープ: $($requiredScopes -join ', ')）"
} 
catch {
    Write-Warning "追加スコープでの接続に失敗しました。制限付きの機能で続行します: $_"
}
#endregion

#region OneDriveデータ取得
# 出力用のユーザーリスト
$userList = @()

Write-Output "OneDriveユーザー情報取得を開始します..."
if ($isAdmin) {
    # グローバル管理者の場合、すべてのユーザー情報を取得
    $allUsers = Get-MgUser -All -Property DisplayName,Mail,onPremisesSamAccountName,AccountEnabled,onPremisesLastSyncDateTime,UserType
    $userCount = $allUsers.Count
    $processedCount = 0
    
    Write-Output "全ユーザー数: $userCount 人"
    foreach ($user in $allUsers) {
        $processedCount++
        $percentComplete = [math]::Round(($processedCount / $userCount) * 100, 0)
        Write-Progress -Activity "ユーザー情報を取得中..." -Status "$processedCount / $userCount ユーザー ($percentComplete%)" -PercentComplete $percentComplete
        
        # ユーザー種別の判定
        Write-Output "処理中: $($user.DisplayName) (ID: $($user.Id), UserType: $($user.UserType))"
        
        # まずUserTypeを基準に判定
        if ($user.UserType -eq "Guest") {
            $userTypeValue = "Guest"
            Write-Output "  ゲストユーザーを特定: $($user.DisplayName)"
        } else {
            $userTypeValue = "Member"
            Write-Output "  通常のメンバーユーザー: $($user.DisplayName)"
        }
        
        # グローバル管理者の場合（優先的に設定）
        $isUserAdmin = $false
        if ($roleMembers.Count -gt 0) {
            foreach ($member in $roleMembers) {
                # 明示的にロールメンバーと比較（大文字小文字を区別しない）
                if ($member.Id -eq $user.Id) {
                    $isUserAdmin = $true
                    $userTypeValue = "Administrator"
                    Write-Output "  グローバル管理者を特定: $($user.DisplayName)"
                    break
                }
            }
        }
        
        try {
            $drive = Get-MgUserDrive -UserId $user.UserPrincipalName -ErrorAction Stop
            $totalGB = [math]::Round($drive.Quota.Total / 1GB, 2)
            $usedGB = [math]::Round($drive.Quota.Used / 1GB, 2)
            $remainingGB = [math]::Round(($drive.Quota.Total - $drive.Quota.Used) / 1GB, 2)
            $usagePercent = [math]::Round(($drive.Quota.Used / $drive.Quota.Total) * 100, 2)
            Write-Output "  OneDrive情報: 総容量=$totalGB GB, 使用量=$usedGB GB, 使用率=$usagePercent%"
            
            # OneDrive状態の詳細な判定
            if ($usagePercent -ge 90) {
                $oneDriveStatus = "使用率高（90%以上）"
            } elseif ($usagePercent -ge 70) {
                $oneDriveStatus = "使用率中（70%以上90%未満）"
            } elseif ($usagePercent -gt 0) {
                $oneDriveStatus = "使用率低（0%超70%未満）"
            } elseif ($usagePercent -eq 0) {
                $oneDriveStatus = "未使用（設定済）"
            } else {
                $oneDriveStatus = "設定済"
            }
        } catch {
            Write-Output "  OneDrive情報取得エラー: $_"
            $totalGB = "取得不可"
            $usedGB = "取得不可"
            $remainingGB = "取得不可"
            $usagePercent = "取得不可"
            $oneDriveStatus = "未設定"
        }
        
        $userList += [PSCustomObject]@{
            "ユーザー名"       = $user.DisplayName
            "メールアドレス"   = $user.Mail
            "ログインユーザー名" = if($user.onPremisesSamAccountName){$user.onPremisesSamAccountName}else{"同期なし"}
            "ユーザー種別"   = $userTypeValue
            "アカウント状態"   = if($user.AccountEnabled){"有効"}else{"無効"}
            "最終同期日時"   = if($user.onPremisesLastSyncDateTime){$user.onPremisesLastSyncDateTime}else{"同期情報なし"}
            "総容量(GB)"   = $totalGB
            "使用容量(GB)"   = $usedGB
            "残り容量(GB)"   = $remainingGB
            "使用率(%)"     = $usagePercent
            "OneDrive状態"  = $oneDriveStatus
        }
    }
    Write-Progress -Activity "ユーザー情報を取得中..." -Completed
} else {
    # 一般ユーザーまたはゲストの場合、自分自身の情報のみ取得
    try {
        Write-Output "自分自身のOneDrive情報を取得します..."
        
        # ユーザー種別の判定（自分自身）
        # UserTypeを基準に判定
        if ($currentUser.UserType -eq "Guest") {
            $userTypeValue = "Guest"
            Write-Output "自分はゲストユーザーです"
        } else {
            $userTypeValue = "Member"
            Write-Output "自分は通常のメンバーユーザーです"
        }
        
        # グローバル管理者かどうかを判定（優先的に設定）
        $isUserAdmin = $false
        if ($roleMembers.Count -gt 0) {
            Write-Output "ロールメンバー数: $($roleMembers.Count)"
            foreach ($member in $roleMembers) {
                Write-Output "比較: ロールID=$($member.Id), 自分のID=$($currentUser.Id)"
                if ($member.Id -eq $currentUser.Id) {
                    $isUserAdmin = $true
                    $userTypeValue = "Administrator"
                    Write-Output "自分はグローバル管理者権限を持っています"
                    break
                }
            }
        }
        Write-Output "最終判定：自分のユーザー種別: $userTypeValue"
        
        $myDrive = Get-MgUserDrive -UserId $UserUPN -ErrorAction Stop
        $totalGB = [math]::Round($myDrive.Quota.Total / 1GB, 2)
        $usedGB = [math]::Round($myDrive.Quota.Used / 1GB, 2)
        $remainingGB = [math]::Round(($myDrive.Quota.Total - $myDrive.Quota.Used) / 1GB, 2)
        $usagePercent = [math]::Round(($myDrive.Quota.Used / $myDrive.Quota.Total)*100, 2)
        Write-Output "自分のOneDrive情報: 総容量=$totalGB GB, 使用量=$usedGB GB, 使用率=$usagePercent%"
        
        # OneDrive状態の詳細な判定
        if ($usagePercent -ge 90) {
            $oneDriveStatus = "使用率高（90%以上）"
        } elseif ($usagePercent -ge 70) {
            $oneDriveStatus = "使用率中（70%以上90%未満）"
        } elseif ($usagePercent -gt 0) {
            $oneDriveStatus = "使用率低（0%超70%未満）"
        } elseif ($usagePercent -eq 0) {
            $oneDriveStatus = "未使用（設定済）"
        } else {
            $oneDriveStatus = "設定済（詳細不明）"
        }
    } catch {
        Write-Output "OneDrive情報取得エラー: $_"
        $totalGB = "取得不可"
        $usedGB = "取得不可"
        $remainingGB = "取得不可"
        $usagePercent = "取得不可"
        $oneDriveStatus = "未設定"
    }
    
    $userList += [PSCustomObject]@{
        "ユーザー名"       = $currentUser.DisplayName
        "メールアドレス"   = $currentUser.Mail
        "ログインユーザー名" = if($currentUser.onPremisesSamAccountName){$currentUser.onPremisesSamAccountName}else{"同期なし"}
        "ユーザー種別"   = $userTypeValue
        "アカウント状態"   = if($currentUser.AccountEnabled){"有効"}else{"無効"}
        "最終同期日時"   = if($currentUser.onPremisesLastSyncDateTime){$currentUser.onPremisesLastSyncDateTime}else{"同期情報なし"}
        "総容量(GB)"   = $totalGB
        "使用容量(GB)"   = $usedGB
        "残り容量(GB)"   = $remainingGB
        "使用率(%)"     = $usagePercent
        "OneDrive状態"  = $oneDriveStatus
    }
}
#endregion

#region レポート作成
# タイムスタンプ
$timestamp = Get-Date -Format "yyyyMMddHHmmss"

# 出力ファイル名の設定
$csvFile = "OneDriveCheck.$timestamp.csv"
$logFile = "OneDriveCheck.$timestamp.txt"
$htmlFile = "OneDriveCheck.$timestamp.html"
$jsFile = "OneDriveCheck.$timestamp.js"

# 出力パスの設定（日付フォルダに変更）
$csvPath = (Join-Path -Path $dateFolderPath -ChildPath $csvFile)
$logPath = (Join-Path -Path $dateFolderPath -ChildPath $logFile)
$htmlPath = (Join-Path -Path $dateFolderPath -ChildPath $htmlFile)
$jsPath = (Join-Path -Path $dateFolderPath -ChildPath $jsFile)

# JavaScriptファイル名のみ（HTMLからの相対パス用）
$jsFileNameOnly = $jsFile

# CSV出力（文字化け対策済み）
try {
    # PowerShell Core (バージョン 6.0以上)の場合
    if ($PSVersionTable.PSVersion.Major -ge 6) {
        $userList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8BOM
    }
    # PowerShell 5.1以下の場合
    else {
        $userList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        # BOMを追加して文字化け対策
        $content = [System.IO.File]::ReadAllText($csvPath)
        [System.IO.File]::WriteAllText($csvPath, $content, [System.Text.Encoding]::UTF8)
    }
    Write-Output "CSVファイルを作成しました: $csvPath"
    
    # CSVファイルをExcelで開き、列幅の調整とフィルターの適用を行う
    try {
        Write-Output "Excelでファイルを開いて列幅の調整とフィルターの適用を行います..."
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $workbook = $excel.Workbooks.Open($csvPath)
        $worksheet = $workbook.Worksheets.Item(1)
        
        # 列幅の自動調整
        $usedRange = $worksheet.UsedRange
        $usedRange.Columns.AutoFit() | Out-Null
        
        # フィルターの適用
        $usedRange.AutoFilter() | Out-Null
        
        # ウィンドウを最前面に表示
        $excel.ActiveWindow.WindowState = -4143 # xlMaximized
        
        # 変更を保存
        $workbook.Save()
        
        Write-Output "Excelでの処理が完了しました。"
    }
    catch {
        Write-Warning "Excelでの処理中にエラーが発生しました: $_"
        Write-Warning "CSVファイルは正常に作成されましたが、Excel処理はスキップされました。"
    }
}
catch {
    Write-Error "CSVファイルの作成中にエラーが発生しました: $_"
}

# ログ出力
$userList | Format-Table -AutoSize | Out-File -FilePath $logPath -Encoding UTF8
Write-Output "ログファイルを作成しました: $logPath"

# JavaScript ファイルの生成 (文字化け対策済み)
$jsContent = @"
// OneDriveCheck データ操作用 JavaScript

// グローバル変数
let currentPage = 1;
let rowsPerPage = 10; // デフォルトの1ページあたりの行数
let filteredRows = []; // フィルタリングされた行を保持する配列

// テーブルを検索する関数（インクリメンタル検索対応）
function searchTable() {
    var input = document.getElementById('searchInput').value.toLowerCase();
    var table = document.getElementById('userTable');
    var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    filteredRows = [];
    
    for (var i = 0; i < rows.length; i++) {
        var found = false;
        var cells = rows[i].getElementsByTagName('td');
        var rowData = {};
        
        for (var j = 0; j < cells.length; j++) {
            var cellText = cells[j].textContent || cells[j].innerText;
            // 列のヘッダー名を取得
            var headerText = table.getElementsByTagName('thead')[0].getElementsByTagName('th')[j].textContent;
            rowData[headerText] = cellText;
            
            if (cellText.toLowerCase().indexOf(input) > -1) {
                found = true;
            }
        }
        
        if (found) {
            filteredRows.push({row: rows[i], data: rowData});
        }
    }
    
    // 検索候補の表示
    showSearchSuggestions(input);
    
    // 検索結果が空の場合は検索候補を非表示
    if (filteredRows.length === 0 && input.length > 0) {
        document.getElementById('searchSuggestions').innerHTML = '<div class="suggestion-item">検索結果がありません</div>';
        document.getElementById('searchSuggestions').style.display = 'block';
    }
    
    // ページングの更新
    currentPage = 1;
    updatePagination();
}

// 検索候補を表示する関数
function showSearchSuggestions(input) {
    var suggestionsDiv = document.getElementById('searchSuggestions');
    suggestionsDiv.innerHTML = '';
    
    if (input.length < 1) {
        suggestionsDiv.style.display = 'none';
        return;
    }
    
    // 一致する値を収集（重複なし）
    var matches = new Set();
    filteredRows.forEach(item => {
        Object.values(item.data).forEach(value => {
            if (value.toLowerCase().indexOf(input.toLowerCase()) > -1) {
                matches.add(value);
            }
        });
    });
    
    // 最大5件まで表示（より見やすく）
    var count = 0;
    matches.forEach(match => {
        if (count < 5) {
            var div = document.createElement('div');
            div.className = 'suggestion-item';
            div.textContent = match;
            div.onclick = function() {
                document.getElementById('searchInput').value = match;
                searchTable();
                suggestionsDiv.style.display = 'none';
            };
            suggestionsDiv.appendChild(div);
            count++;
        }
    });
    
    if (count > 0) {
        suggestionsDiv.style.display = 'block';
    } else if (input.length > 0) {
        // 検索結果がない場合のメッセージ
        var noResults = document.createElement('div');
        noResults.className = 'suggestion-item no-results';
        noResults.textContent = '検索結果がありません';
        suggestionsDiv.appendChild(noResults);
        suggestionsDiv.style.display = 'block';
    } else {
        suggestionsDiv.style.display = 'none';
    }
}

// 列フィルターを作成する関数
function createColumnFilters() {
    var table = document.getElementById('userTable');
    var headers = table.getElementsByTagName('thead')[0].getElementsByTagName('th');
    var filterRow = document.createElement('tr');
    filterRow.className = 'filter-row';
    
    for (var i = 0; i < headers.length; i++) {
        var cell = document.createElement('th');
        var select = document.createElement('select');
        select.className = 'column-filter';
        select.setAttribute('data-column', i);
        
        // デフォルトのオプション
        var defaultOption = document.createElement('option');
        defaultOption.value = '';
        defaultOption.textContent = 'すべて';
        select.appendChild(defaultOption);
        
        // 列の一意の値を取得
        var uniqueValues = new Set();
        var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
        for (var j = 0; j < rows.length; j++) {
            var cellValue = rows[j].getElementsByTagName('td')[i].textContent;
            uniqueValues.add(cellValue);
        }
        
        // 一意の値をソートしてオプションとして追加
        Array.from(uniqueValues).sort().forEach(value => {
            var option = document.createElement('option');
            option.value = value;
            option.textContent = value;
            select.appendChild(option);
        });
        
        // 変更イベントリスナーを追加
        select.addEventListener('change', applyColumnFilters);
        
        cell.appendChild(select);
        filterRow.appendChild(cell);
    }
    
    // フィルター行をテーブルヘッダーに追加
    table.getElementsByTagName('thead')[0].appendChild(filterRow);
}

// 列フィルターを適用する関数
function applyColumnFilters() {
    var table = document.getElementById('userTable');
    var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    var filters = document.getElementsByClassName('column-filter');
    filteredRows = [];
    
    // 各行に対してフィルターを適用
    for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        var cells = row.getElementsByTagName('td');
        var rowData = {};
        var includeRow = true;
        
        // 各フィルターをチェック
        for (var j = 0; j < filters.length; j++) {
            var filter = filters[j];
            var columnIndex = parseInt(filter.getAttribute('data-column'));
            var filterValue = filter.value;
            
            // 列のヘッダー名を取得
            var headerText = table.getElementsByTagName('thead')[0].getElementsByTagName('th')[columnIndex].textContent;
            var cellValue = cells[columnIndex].textContent;
            rowData[headerText] = cellValue;
            
            // フィルター値が設定されていて、セルの値と一致しない場合は行を除外
            if (filterValue && cellValue !== filterValue) {
                includeRow = false;
                break;
            }
        }
        
        if (includeRow) {
            filteredRows.push({row: row, data: rowData});
        }
    }
    
    // 検索フィールドの値も考慮
    var searchInput = document.getElementById('searchInput').value.toLowerCase();
    if (searchInput) {
        filteredRows = filteredRows.filter(item => {
            return Object.values(item.data).some(value => 
                value.toLowerCase().indexOf(searchInput) > -1
            );
        });
    }
    
    // ページングの更新
    currentPage = 1;
    updatePagination();
}

// ページングを更新する関数
function updatePagination() {
    var table = document.getElementById('userTable');
    var tbody = table.getElementsByTagName('tbody')[0];
    var rows = tbody.getElementsByTagName('tr');
    
    // すべての行を非表示にする
    for (var i = 0; i < rows.length; i++) {
        rows[i].style.display = 'none';
    }
    
    // フィルタリングされた行のみを表示
    var startIndex = (currentPage - 1) * rowsPerPage;
    var endIndex = Math.min(startIndex + rowsPerPage, filteredRows.length);
    
    for (var i = startIndex; i < endIndex; i++) {
        filteredRows[i].row.style.display = '';
    }
    
    // ページネーションコントロールを更新
    updatePaginationControls();
}

// ページネーションコントロールを更新する関数
function updatePaginationControls() {
    var paginationDiv = document.getElementById('pagination');
    paginationDiv.innerHTML = '';
    
    var totalPages = Math.ceil(filteredRows.length / rowsPerPage);
    
    // 「前へ」ボタン
    var prevButton = document.createElement('button');
    prevButton.innerHTML = '<span class="button-icon">◀</span>前へ';
    prevButton.disabled = currentPage === 1;
    prevButton.addEventListener('click', function() {
        if (currentPage > 1) {
            currentPage--;
            updatePagination();
        }
    });
    paginationDiv.appendChild(prevButton);
    
    // ページ番号
    var pageInfo = document.createElement('span');
    pageInfo.className = 'page-info';
    pageInfo.textContent = currentPage + ' / ' + (totalPages || 1) + ' ページ';
    paginationDiv.appendChild(pageInfo);
    
    // 「次へ」ボタン
    var nextButton = document.createElement('button');
    nextButton.innerHTML = '次へ<span class="button-icon">▶</span>';
    nextButton.disabled = currentPage === totalPages || totalPages === 0;
    nextButton.addEventListener('click', function() {
        if (currentPage < totalPages) {
            currentPage++;
            updatePagination();
        }
    });
    paginationDiv.appendChild(nextButton);
    
    // 1ページあたりの行数を選択
    var rowsPerPageDiv = document.createElement('div');
    rowsPerPageDiv.className = 'rows-per-page';
    
    var rowsPerPageLabel = document.createElement('span');
    rowsPerPageLabel.textContent = '表示件数: ';
    rowsPerPageDiv.appendChild(rowsPerPageLabel);
    
    var rowsPerPageSelect = document.createElement('select');
    [10, 20, 50, 100].forEach(function(value) {
        var option = document.createElement('option');
        option.value = value;
        option.textContent = value + '件';
        if (value === rowsPerPage) {
            option.selected = true;
        }
        rowsPerPageSelect.appendChild(option);
    });
    
    rowsPerPageSelect.addEventListener('change', function() {
        rowsPerPage = parseInt(this.value);
        currentPage = 1;
        updatePagination();
    });
    
    rowsPerPageDiv.appendChild(rowsPerPageSelect);
    paginationDiv.appendChild(rowsPerPageDiv);
    
    // 総件数表示
    var totalItems = document.createElement('span');
    totalItems.className = 'total-items';
    totalItems.textContent = '全 ' + filteredRows.length + ' 件';
    paginationDiv.appendChild(totalItems);
}

// 検索入力フィールドからフォーカスが外れたときに検索候補を非表示にする
function hideSearchSuggestions() {
    // 少し遅延させて、候補をクリックする時間を確保
    setTimeout(function() {
        document.getElementById('searchSuggestions').style.display = 'none';
    }, 200);
}

// CSVとしてエクスポートする関数 (文字化け対策済み)
function exportTableToCSV() {
    var table = document.getElementById('userTable');
    var headerRow = table.getElementsByTagName('thead')[0].getElementsByTagName('tr')[0]; // ヘッダー行（1行目）のみ
    var bodyRows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    var csv = [];
    
    // ヘッダー行を処理
    var headerCols = headerRow.getElementsByTagName('th');
    var headerData = [];
    for (var i = 0; i < headerCols.length; i++) {
        var data = headerCols[i].innerText.replace(/(\r\n|\n|\r)/gm, ' ').replace(/"/g, '""');
        headerData.push('"' + data + '"');
    }
    csv.push(headerData.join(','));
    
    // データ行を処理（フィルター行は除外）
    for (var i = 0; i < bodyRows.length; i++) {
        var row = [], cols = bodyRows[i].getElementsByTagName('td');
        for (var j = 0; j < cols.length; j++) {
            // セル内のテキストから改行や引用符を適切に処理
            var data = cols[j].innerText.replace(/(\r\n|\n|\r)/gm, ' ').replace(/"/g, '""');
            row.push('"' + data.trim() + '"'); // 余分な空白を削除
        }
        csv.push(row.join(','));
    }
    
    // CSVファイルのダウンロード（UTF-8 BOM付きで文字化け対策）
    var csvContent = '\uFEFF' + csv.join('\n'); // BOMを追加
    var csvFile = new Blob([csvContent], {type: 'text/csv;charset=utf-8'});
    var downloadLink = document.createElement('a');
    downloadLink.download = 'OneDriveCheck_Export.csv';
    downloadLink.href = window.URL.createObjectURL(csvFile);
    downloadLink.style.display = 'none';
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
}

// 印刷機能
function printTable() {
    window.print();
}

// 表の行に色を付ける
function colorizeRows() {
    var table = document.getElementById('userTable');
    var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');

    for (var i = 0; i < rows.length; i++) {
        var oneDriveStatusCell = rows[i].querySelector('td:nth-child(11)'); // OneDrive状態のセル
        var usageCell = rows[i].querySelector('td:nth-child(10)'); // 使用率のセル
        var userTypeCell = rows[i].querySelector('td:nth-child(4)');        // ユーザー種別のセル

        // ユーザー種別によるスタイル適用
        if (userTypeCell) {
            // 適切なアイコンを追加
            var userType = userTypeCell.textContent.trim();
            if (userType === "Administrator") {
                userTypeCell.innerHTML = '<span style="color: #d32f2f;">👑 ' + userType + '</span>';
            } else if (userType === "Guest") {
                userTypeCell.innerHTML = '<span style="color: #7B1FA2;">👤 ' + userType + '</span>';
            } else if (userType === "Member") {
                userTypeCell.innerHTML = '<span style="color: #0288D1;">👥 ' + userType + '</span>';
            }
        }

        // OneDrive状態による行のスタイル適用
        if (oneDriveStatusCell) {
            var status = oneDriveStatusCell.textContent.trim();
            if (status === "未設定") {
                rows[i].classList.add('not-configured');
            } else if (status.includes("90%以上") || usageCell && parseFloat(usageCell.textContent) >= 90) {
                rows[i].classList.add('danger');
            } else if (status.includes("70%以上") || usageCell && parseFloat(usageCell.textContent) >= 70) {
                rows[i].classList.add('warning');
            } else {
                rows[i].classList.add('normal');
            }
        }
        
        // アカウント状態によっても色分け
        var accountStatus = rows[i].querySelector('td:nth-child(5)').textContent; // アカウント状態のセル
        if (accountStatus.includes('無効')) {
            rows[i].classList.add('disabled');
        }
    }
}

// ページロード時に実行
window.onload = function() {
    colorizeRows();
    createColumnFilters();
    
    // 検索イベントリスナーを設定
    document.getElementById('searchInput').addEventListener('keyup', function(e) {
        // リアルタイムで検索を実行（インクリメンタル検索）
        searchTable();
    });
    document.getElementById('searchInput').addEventListener('blur', hideSearchSuggestions);
    
    // エクスポートボタンにイベントリスナーを設定
    document.getElementById('exportBtn').addEventListener('click', exportTableToCSV);
    
    // 印刷ボタンにイベントリスナーを設定
    document.getElementById('printBtn').addEventListener('click', printTable);
    
    // 初期ページングの設定
    var table = document.getElementById('userTable');
    var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    for (var i = 0; i < rows.length; i++) {
        var cells = rows[i].getElementsByTagName('td');
        var rowData = {};
        
        for (var j = 0; j < cells.length; j++) {
            var headerText = table.getElementsByTagName('thead')[0].getElementsByTagName('th')[j].textContent;
            rowData[headerText] = cells[j].textContent;
        }
        
        filteredRows.push({row: rows[i], data: rowData});
    }
    
    updatePagination();
};
"@

# JavaScript ファイルを出力
$jsContent | Out-File -FilePath $jsPath -Encoding UTF8
Write-Output "JavaScriptファイルを作成しました: $jsPath"

# 実行日時とユーザー情報を取得
$executionDateFormatted = $executionTime.ToString("yyyy/MM/dd HH:mm:ss")
$executorName = $currentUser.DisplayName
$userType = if($currentUser.UserType){$currentUser.UserType}else{"未定義"}

# HTML ファイルの生成
$htmlContent = @"
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <title>OneDrive 利用状況レポート</title>
    <script src="./$jsFileNameOnly"></script>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }
        .container {
            background-color: white;
            padding: 20px;
            border-radius: 5px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }
        .header {
            background-color: #0078d4;
            color: white;
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 5px;
            display: flex;
            align-items: center;
        }
        .header-icon {
            font-size: 24px;
            margin-right: 10px;
        }
        h1 {
            margin: 0;
            font-size: 24px;
        }
        .info-section {
            background-color: #f0f0f0;
            padding: 10px;
            margin-bottom: 20px;
            border-radius: 5px;
            font-size: 14px;
        }
        .info-label {
            font-weight: bold;
            margin-right: 5px;
        }
        .toolbar {
            margin-bottom: 20px;
            display: flex;
            gap: 10px;
            align-items: center;
            position: relative;
        }
        #searchInput {
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
            flex-grow: 1;
        }
        #searchSuggestions {
            position: absolute;
            top: 100%;
            left: 0;
            width: 100%;
            max-height: 200px;
            overflow-y: auto;
            background-color: white;
            border: 1px solid #ddd;
            border-radius: 0 0 4px 4px;
            z-index: 1000;
            display: none;
        }
        .suggestion-item {
            padding: 8px;
            border-bottom: 1px solid #eee;
            cursor: pointer;
        }
        .suggestion-item:hover {
            background-color: #f0f0f0;
        }
        .suggestion-item.no-results {
            color: #999;
            font-style: italic;
            cursor: default;
        }
        button {
            padding: 8px 12px;
            background-color: #0078d4;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            display: flex;
            align-items: center;
        }
        button:hover {
            background-color: #106ebe;
        }
        button:disabled {
            background-color: #cccccc;
            cursor: not-allowed;
        }
        .button-icon {
            margin-right: 5px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
        }
        th, td {
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }
        th {
            background-color: #f2f2f2;
            font-weight: bold;
        }
        .filter-row th {
            padding: 5px;
        }
        .column-filter {
            width: 100%;
            padding: 5px;
            border: 1px solid #ddd;
            border-radius: 4px;
        }
        tr.danger {
            background-color: #ffebee;
        }
        tr.warning {
            background-color: #fff8e1;
        }
        tr.normal {
            background-color: #f1f8e9;
        }
        tr.not-configured {
            background-color: #e0e0e0;
            font-style: italic;
        }
        tr.disabled {
            color: #999;
            font-style: italic;
        }
        .status-icon {
            margin-right: 5px;
        }
        #pagination {
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 10px;
            flex-wrap: wrap;
            margin-bottom: 20px;
        }
        .page-info {
            margin: 0 10px;
        }
        .rows-per-page {
            margin-left: 20px;
            display: flex;
            align-items: center;
        }
        .total-items {
            margin-left: 15px;
        }
        
        @media print {
            .toolbar, button, #pagination, .filter-row {
                display: none;
            }
            body {
                background-color: white;
                margin: 0;
            }
            .container {
                box-shadow: none;
                padding: 0;
            }
            .header {
                background-color: black !important;
                color: white !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
            th {
                background-color: #f2f2f2 !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
            tr.danger {
                background-color: #ffebee !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
            tr.warning {
                background-color: #fff8e1 !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
            tr.normal {
                background-color: #f1f8e9 !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="header-icon">📊</div>
            <h1>OneDrive 利用状況レポート</h1>
        </div>
        
        <div class="info-section">
            <p><span class="info-label">実行日時:</span> $executionDateFormatted</p>
            <p><span class="info-label">実行者:</span> $executorName</p>
            <p><span class="info-label">実行者の種別:</span> $userType</p>
            <p><span class="info-label">実行モード:</span> $(if($isAdmin){"管理者モード (Administrator)"}else{"ユーザーモード"})</p>
            <p><span class="info-label">出力フォルダ:</span> $dateFolderPath</p>
        </div>
        
        <div class="toolbar">
            <input type="text" id="searchInput" placeholder="検索...">
            <div id="searchSuggestions"></div>
            <button id="exportBtn"><span class="button-icon">📥</span>CSVエクスポート</button>
            <button id="printBtn"><span class="button-icon">🖨️</span>印刷</button>
        </div>
        
        <div id="pagination"></div>

        <table id="userTable">
            <thead>
                <tr>
                    <th>ユーザー名</th>
                    <th>メールアドレス</th>
                    <th>ログインユーザー名</th>
                    <th>ユーザー種別</th>
                    <th>アカウント状態</th>
                    <th>最終同期日時</th>
                    <th>総容量(GB)</th>
                    <th>使用容量(GB)</th>
                    <th>残り容量(GB)</th>
                    <th>使用率(%)</th>
                    <th>OneDrive状態</th>
                </tr>
            </thead>
            <tbody>
"@

# HTML テーブル本体の作成
foreach ($user in $userList) {
    # アカウント状態に応じたアイコンを設定
    $statusIcon = if ($user.'アカウント状態' -eq "有効") { "✅" } else { "❌" }
    
    # 行を追加
    $htmlContent += @"
                <tr>
                    <td>$($user.'ユーザー名')</td>
                    <td>$($user.'メールアドレス')</td>
                    <td>$($user.'ログインユーザー名')</td>
                    <td>$($user.'ユーザー種別')</td>
                    <td><span class="status-icon">$statusIcon</span>$($user.'アカウント状態')</td>
                    <td>$($user.'最終同期日時')</td>
                    <td>$($user.'総容量(GB)')</td>
                    <td>$($user.'使用容量(GB)')</td>
                    <td>$($user.'残り容量(GB)')</td>
                    <td>$($user.'使用率(%)')</td>
                    <td>$($user.'OneDrive状態')</td>
                </tr>
"@
}

# HTML 終了部分
$htmlContent += @"
            </tbody>
            <tfoot>
                <tr>
                    <td colspan="11" style="text-align: center; font-style: italic;">※ ユーザー種別: 👑 Administrator = グローバル管理者、👥 Member = 一般ユーザー、👤 Guest = ゲストユーザー</td>
                </tr>
            </tfoot>
        </table>
        
        <div class="info-section">
            <p><span class="info-label">色の凡例:</span></p>
            <p>🟢 緑色の行: 使用率が70%未満のユーザー</p>
            <p>⚫ グレー色の行: OneDriveを設定利用していないユーザー</p>
            <p>⚪ 「未設定」: OneDriveが設定されていないことを示します</p>
            <p>🟡 黄色の行: 使用率が70%以上90%未満のユーザー</p>
            <p>🔴 赤色の行: 使用率が90%以上のユーザー</p>
            <p>⚪ グレーの行: 無効なアカウント</p>
        </div>
    </div>
</body>
</html>
"@

# HTML ファイルを出力 (文字化け対策済み)
$htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8
Write-Output "HTMLファイルを作成しました: $htmlPath"

# 出力ディレクトリを開く
try {
    Start-Process -FilePath "explorer.exe" -ArgumentList $dateFolderPath
} catch {
    Write-Warning "フォルダを開けませんでした: $_"
}

# スクリプト終了待機
Write-Output "処理が完了しました。出力ディレクトリ: $dateFolderPath"
Read-Host "Enterキーを押すと終了します"
#endregion