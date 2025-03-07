# 📊 OneDrive利用状況レポートツール

## 📋 目次
- [概要](#概要)
- [機能説明](#機能説明)
- [必要条件](#必要条件)
- [インストール方法](#インストール方法)
- [使用方法](#使用方法)
- [出力ファイルの説明](#出力ファイルの説明)
- [トラブルシューティング](#トラブルシューティング)
- [GitHub管理方法](#github管理方法)
- [更新履歴](#更新履歴)

## 📝 概要

**OneDrive利用状況レポートツール**は、組織内のOneDriveの利用状況を簡単に把握するためのPowerShellスクリプトです。Microsoft Graph APIを使用して、ユーザーごとのOneDriveの容量使用状況を取得し、見やすいHTMLレポートとして出力します。

> 💡 **ポイント**: 管理者権限で実行すると、組織内の全ユーザーの情報を取得できます。一般ユーザー権限では、自分自身の情報のみが表示されます。

## ✨ 機能説明

### 主な機能

- 🔍 **ユーザー情報の取得**
  - ユーザー名、メールアドレス、ログインユーザー名
  - アカウント状態（有効/無効）
  - 最終同期日時

- 💾 **OneDrive容量情報の取得**
  - 総容量(GB)
  - 使用容量(GB)
  - 残り容量(GB)
  - 使用率(%)

- 📊 **インタラクティブなHTMLレポート**
  - インクリメンタル検索機能（文字入力で検索候補がリアルタイムで表示）
  - 各項目ごとにプルダウンでフィルタリングが可能
  - 10名、20名、50名、100名単位で表示件数を選択可能なページング機能
  - 前へ/次へボタンでページ間を移動できるナビゲーション
  - 使用率に応じた行の色分け表示
  - CSVエクスポート機能
  - 印刷機能

## 🔧 必要条件

- Windows PowerShell 5.1以上 または PowerShell Core 6.0以上
- Microsoft Graph PowerShellモジュール
- Microsoft 365管理者アカウント（全ユーザー情報取得の場合）
- インターネット接続

## 📥 インストール方法

### 1. スクリプトファイルのダウンロード

```powershell
# GitHubからクローン（推奨）
git clone https://github.com/Kensan196948G/OneDriveSimpleManagement.git

# または直接ダウンロード
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Kensan196948G/OneDriveSimpleManagement/main/miraiAllUserInfoComplete.ps1" -OutFile "miraiAllUserInfoComplete.ps1"
```

### 2. Microsoft Graph PowerShellモジュールのインストール

```powershell
# 管理者権限でPowerShellを実行
Install-Module Microsoft.Graph -Scope CurrentUser -Force
```

> ⚠️ **注意**: スクリプトは自動的にモジュールの存在を確認し、必要に応じてインストールを試みますが、事前にインストールしておくことをお勧めします。

## 🚀 使用方法

### 基本的な使用方法

1. PowerShellを起動します
2. スクリプトがあるディレクトリに移動します
3. スクリプトを実行します

```powershell
# 基本実行（カレントディレクトリに出力）
.\miraiAllUserInfoComplete.ps1

# 出力先を指定して実行
.\miraiAllUserInfoComplete.ps1 -OutputDir "C:\Reports"
```

### 出力先の指定について

`-OutputDir` パラメータを使用すると、指定したディレクトリに出力ファイルが生成されます：

> 💡 **重要**: 指定したフォルダ（例：C:\Reports）内に、自動的に日付ベースのフォルダ（例：OneDriveCheck.20250307）が生成され、その中にすべての出力ファイル（CSV、HTML、JS、テキスト）が収納・配置されます。指定したフォルダが存在しない場合は自動的に作成されます。


### 実行の流れ

1. Microsoft Graphへの接続とサインイン
   - 初回実行時はブラウザが開き、Microsoft 365アカウントでのサインインが求められます
   - 必要な権限（User.Read.All, Directory.Read.All, Sites.Read.All）に同意する必要があります

2. ユーザー情報とOneDrive情報の取得
   - 管理者権限の場合：全ユーザーの情報を取得
   - 一般ユーザー権限の場合：自分自身の情報のみ取得

3. レポートの生成
   - 出力先に日付ベースのフォルダ（OneDriveCheck.YYYYMMDD）が作成されます
   - CSV、テキスト、HTML、JavaScriptファイルが生成されます
   - 自動的にExcelでCSVファイルが開かれます（可能な場合）
   - 出力フォルダが自動的に開かれます

## 📂 出力ファイルの説明

スクリプトは以下のファイルを生成します：

| ファイル | 説明 |
|---------|------|
| `OneDriveCheck.YYYYMMDDHHmmss.csv` | ユーザー情報とOneDrive使用状況のCSVデータ |
| `OneDriveCheck.YYYYMMDDHHmmss.txt` | テキスト形式のレポート（コンソール出力用） |
| `OneDriveCheck.YYYYMMDDHHmmss.html` | インタラクティブなHTMLレポート |
| `OneDriveCheck.YYYYMMDDHHmmss.js` | HTMLレポートの機能を提供するJavaScriptファイル |

### HTMLレポートの使い方

HTMLレポートには以下の機能があります：

- **検索機能** 🔍
  - 検索ボックスに文字を入力すると、リアルタイムで検索結果が表示されます
  - 検索候補がドロップダウンで表示されます

- **フィルタリング機能** 🔄
  - 各列のヘッダー下にあるプルダウンメニューから値を選択してフィルタリングできます
  - 複数の列を組み合わせたフィルタリングも可能です

- **ページング機能** 📄
  - デフォルトでは10件ずつ表示されます
  - 表示件数は10、20、50、100件から選択可能です
  - 「前へ」「次へ」ボタンでページ間を移動できます

- **エクスポート機能** 📤
  - 「CSVエクスポート」ボタンをクリックすると、現在の表示内容をCSVファイルとしてダウンロードできます

- **印刷機能** 🖨️
  - 「印刷」ボタンをクリックすると、ブラウザの印刷ダイアログが開きます
  - フィルターやページングコントロールは印刷されません

## ❓ トラブルシューティング

### よくある問題と解決方法

| 問題 | 解決方法 |
|------|---------|
| Microsoft Graphへの接続エラー | インターネット接続を確認し、Microsoft 365アカウントの資格情報が正しいことを確認してください。 |
| 「権限が不足しています」エラー | 管理者に必要な権限（User.Read.All, Directory.Read.All, Sites.Read.All）の付与を依頼してください。 |
| OneDrive情報が「取得不可」と表示される | ユーザーにOneDriveが割り当てられていないか、アクセス権限の問題が考えられます。 |
| CSVファイルが文字化けする | スクリプトはUTF-8 BOM付きでエクスポートしていますが、古いExcelバージョンでは問題が発生する場合があります。その場合はHTMLレポートからCSVエクスポート機能を使用してください。 |

### ログの確認

エラーが発生した場合は、PowerShellコンソールに表示されるエラーメッセージを確認してください。詳細なエラー情報が表示されます。

## 🔄 GitHub管理方法

このツールはGitHubでの管理を前提としています。以下は、GitHubでの管理方法の例です。

### リポジトリの初期設定

```bash
# 新しいリポジトリを作成
mkdir OneDriveSimpleManagement
cd OneDriveSimpleManagement
git init

# スクリプトファイルをリポジトリに追加
copy /path/to/miraiAllUserInfoComplete.ps1 .
copy /path/to/OneDriveCheck_ドキュメント.md .

# .gitignoreファイルの作成
echo "# 出力ディレクトリを無視" > .gitignore
echo "OneDriveCheck.*/" >> .gitignore
echo "# 一時ファイルを無視" >> .gitignore
echo "*.tmp" >> .gitignore
echo "*.log" >> .gitignore

# 初回コミット
git add .
git commit -m "初回コミット: OneDrive利用状況レポートツールの追加"

# GitHubリポジトリとの連携
git remote add origin https://github.com/Kensan196948G/OneDriveSimpleManagement.git
git push -u origin main
```

### 更新の管理

```bash
# 変更を確認
git status

# 変更をステージング
git add miraiAllUserInfoComplete.ps1
git add OneDriveCheck_ドキュメント.md

# 変更をコミット
git commit -m "機能追加: CSVエクスポート機能の改善"

# 変更をGitHubにプッシュ
git push origin main
```

### ブランチを使った開発

```bash
# 新機能用のブランチを作成
git checkout -b feature/improved-filtering

# 変更を加えてコミット
git add miraiAllUserInfoComplete.ps1
git commit -m "フィルタリング機能の強化"

# メインブランチにマージ
git checkout main
git merge feature/improved-filtering

# GitHubにプッシュ
git push origin main
```

## 📜 更新履歴

### バージョン 1.0.0 (2025-03-07)
- 初回リリース
- 基本的なOneDrive使用状況レポート機能
- HTMLレポート出力機能

### バージョン 1.1.0 (2025-03-07)
- インクリメンタル検索機能の追加
- 各項目ごとのプルダウンフィルタリング機能の追加
- ページング機能の改善（10名、20名、50名、100名単位で表示件数を選択可能）
- 前へ/次へボタンでページ間を移動できるナビゲーション機能の追加

### バージョン 1.1.1 (2025-03-07)
- CSVエクスポート機能の改善（フィルター行の除外、データの整形）
- ドキュメントの追加

---

## 👥 貢献者

- 開発者: Kensan196948G
- 貢献者: [Contributor Names]

## 📄 ライセンス

このプロジェクトは[MITライセンス](LICENSE)の下で公開されています。

---

*最終更新日: 2025年3月7日*