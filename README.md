# 📊 OneDrive利用状況レポートツール

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## 概要

**OneDrive利用状況レポートツール**は、Microsoft 365環境でのOneDriveの使用状況を簡単に把握するためのPowerShellスクリプトです。Microsoft Graph APIを使用して、ユーザーごとのOneDriveの容量使用状況を取得し、インタラクティブなHTMLレポートとして出力します。

![OneDriveレポートサンプル](https://via.placeholder.com/800x400?text=OneDrive+レポート+サンプル)

## 利用可能なスクリプトファイルと役割

本リポジトリでは以下のPowerShellスクリプトファイルを提供しています：

| スクリプト名 | 役割 | 実行優先度 |
|-------------|------|-----------|
| **OneDriveSyncManagement.ps1** | メインスクリプト：OneDrive利用状況の収集とレポート生成 | 3️⃣ |
| **Microsoft-Graph-Update-Helper.ps1** | Microsoft Graphモジュール更新用ヘルパー | 1️⃣ |
| **OneDrivePermission-Helper.ps1** | API権限取得用ヘルパー | 2️⃣ |
| OneDriveSimpleManagement.ps1 | 過去バージョン（参考用） | - |
| miraiAllUserInfoComplete_integrated.ps1 | 過去バージョン（参考用） | - |

> 注意：実際に使用するのは上部3つのスクリプトです。残りは参照用として保持されています。

## 実行手順と実行順序

以下の順序でスクリプトを実行することを推奨します：

### 1️⃣ Microsoft Graphモジュール更新（初回または更新時）
```powershell
# 管理者権限のPowerShellで実行
.\Microsoft-Graph-Update-Helper.ps1
```
- **目的**: Microsoft Graphモジュールを最新バージョンに更新し、実行環境を整備
- **実行頻度**: 初回実行時、またはエラーが発生した場合
- **管理者権限**: 必須（モジュールのインストール・更新に必要）

### 2️⃣ API権限の確認と取得
```powershell
# 一般ユーザー権限でも実行可能（管理者権限推奨）
.\OneDrivePermission-Helper.ps1
```
- **目的**: テナントIDの確認・更新とAPI権限の取得・確認
- **実行頻度**: 初回実行時、またはAPI権限エラーが発生した場合
- **管理者権限**: 推奨（グローバル管理者の場合、権限承認が容易）

### 3️⃣ OneDrive利用状況レポートの生成
```powershell
# 管理者権限で実行（推奨）
Start-Process PowerShell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File `"$PWD\OneDriveSyncManagement.ps1`""

# または直接実行
.\OneDriveSyncManagement.ps1
```
- **目的**: OneDriveの利用状況データ収集とレポート生成
- **実行頻度**: 利用状況の確認が必要な時
- **管理者権限**: 推奨（全ユーザーの情報を取得する場合は必須）
- **出力**: CSV、HTML、テキストログ形式のレポート

## 詳細実行ガイド

### セットアップと初回実行

1. **環境準備（初回のみ）**
   ```powershell
   # 1. 管理者権限でPowerShellを開く
   # 2. 作業ディレクトリに移動
   cd "パス\OneDriveSimpleManagement"
   
   # 3. Microsoft Graphモジュールの更新
   .\Microsoft-Graph-Update-Helper.ps1
   
   # 4. API権限の確認と取得
   .\OneDrivePermission-Helper.ps1
   # （プロンプトでテナントIDの自動更新に「Y」と回答）
   ```

2. **レポート生成（定期的に実行）**
   ```powershell
   # 管理者権限で実行（すべてのユーザー情報を取得）
   Start-Process PowerShell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File `"$PWD\OneDriveSyncManagement.ps1`""
   
   # 出力先を指定して実行
   .\OneDriveSyncManagement.ps1 -OutputDir "C:\Reports"
   ```

### エラー発生時のトラブルシューティング

1. **Microsoft Graph認証エラー**
   ```powershell
   # モジュールを更新
   .\Microsoft-Graph-Update-Helper.ps1
   ```

2. **API権限エラー**
   ```powershell
   # 権限確認・取得
   .\OneDrivePermission-Helper.ps1
   ```

3. **テナントID関連エラー**
   ```powershell
   # テナントIDを確認・更新
   .\OneDrivePermission-Helper.ps1
   # （プロンプトでテナントIDの自動更新に「Y」と回答）
   ```
