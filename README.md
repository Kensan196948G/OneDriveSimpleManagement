# 📊 OneDrive利用状況レポートツール

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## 概要

**OneDrive利用状況レポートツール**は、Microsoft 365環境でのOneDriveの使用状況を簡単に把握するためのPowerShellスクリプトです。Microsoft Graph APIを使用して、ユーザーごとのOneDriveの容量使用状況を取得し、インタラクティブなHTMLレポートとして出力します。

![OneDriveレポートサンプル](https://via.placeholder.com/800x400?text=OneDrive+レポート+サンプル)

## 主な機能

- 🔍 **ユーザー情報の取得**
  - ユーザー名、メールアドレス、ログインユーザー名など
  - アカウント状態（有効/無効）の確認

- 💾 **OneDrive容量情報の取得**
  - 総容量、使用容量、残り容量、使用率の表示
  - 使用率に応じた色分け表示

- 📊 **高機能なHTMLレポート**
  - インクリメンタル検索機能
  - 各項目ごとのフィルタリング
  - 表示件数を選択可能なページング機能
  - CSVエクスポート機能
  - 印刷機能

## クイックスタート

### 基本実行（カレントディレクトリに出力）
```powershell
# 1. リポジトリをクローン
git clone https://github.com/Kensan196948G/OneDriveSimpleManagement.git
cd OneDriveSimpleManagement

# 2. スクリプトを実行
.\miraiAllUserInfoComplete.ps1
```

### 直接ダウンロードして実行
```powershell
# スクリプトを直接ダウンロード
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Kensan196948G/OneDriveSimpleManagement/main/miraiAllUserInfoComplete.ps1" -OutFile "miraiAllUserInfoComplete.ps1"

# スクリプトを実行
.\miraiAllUserInfoComplete.ps1
```

### 出力先を指定して実行
```powershell
.\miraiAllUserInfoComplete.ps1 -OutputDir "C:\Reports"
```
> **注**: 指定したフォルダ内に日付ベースのフォルダが自動生成され、その中にすべての出力ファイルが収納されます。

## 詳細ドキュメント

詳細な使用方法、機能説明、トラブルシューティングについては、[OneDriveCheck_ドキュメント.md](OneDriveCheck_ドキュメント.md)を参照してください。

## 必要条件

- Windows PowerShell 5.1以上 または PowerShell Core 6.0以上
- Microsoft Graph PowerShellモジュール
- Microsoft 365アカウント（管理者権限推奨）

## スクリーンショット

### レポート画面
![レポート画面](https://via.placeholder.com/400x200?text=レポート画面)

### フィルタリング機能
![フィルタリング機能](https://via.placeholder.com/400x200?text=フィルタリング機能)

### 検索機能
![検索機能](https://via.placeholder.com/400x200?text=検索機能)

## ライセンス

このプロジェクトは[MITライセンス](LICENSE)の下で公開されています。

## 貢献

バグ報告や機能リクエストは、GitHubのIssueで受け付けています。プルリクエストも歓迎します！

---

*最終更新日: 2025年3月7日*