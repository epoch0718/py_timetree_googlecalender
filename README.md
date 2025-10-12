# TimeTree to Google Calendar Scraper

## 概要

このプロジェクトは、[TimeTree](https://timetreeapp.com/) のカレンダーから予定を自動的にスクレイピングし、指定した Google Calendar に同期するためのツールです。
Playwright を使用して TimeTree から情報を取得し、Google Apps Script (GAS) を介して Google Calendar にイベントを登録します。
一連の処理は GitHub Actions によって定期的に自動実行されます。

## 主な機能

- **TimeTree スクレイピング**: Playwright を利用して、TimeTree のウェブサイトにログインし、カレンダーのイベント情報を取得します。
- **Google Calendar 同期**: 取得したイベント情報を Google Apps Script で作成したウェブアプリに送信し、Google Calendar にイベントとして登録します。
- **自動化**: GitHub Actions を使用して、毎日定刻にスクレイピングと同期処理を自動で実行します。

## 必要なもの

- TimeTree アカウント
- Google アカウント
- Python 3.11 以降

## セットアップ手順

### 1. リポジトリのクローン

```bash
git clone https://github.com/your-username/py_timetree_scraper.git
cd py_timetree_scraper
```

### 2. Google Apps Script (GAS) のデプロイ

1.  [Google Apps Script](https://script.google.com/home) にアクセスし、新しいプロジェクトを作成します。
2.  `gas.gs` ファイルの内容をコピーし、GAS エディタに貼り付けます。
3.  プロジェクトを保存し、右上の「デプロイ」>「新しいデプロイ」を選択します。
4.  「種類の選択」で「ウェブアプリ」を選択します。
5.  以下の設定を行います。
    - **説明**: (任意) `TimeTree Sync` など
    - **次のユーザーとして実行**: 自分
    - **アクセスできるユーザー**: 全員
6.  「デプロイ」をクリックします。初回デプロイ時には、カレンダーへのアクセス許可を求められるので承認してください。
7.  表示された**ウェブアプリ URL** をコピーしておきます。これは後で使います。

### 3. Python 環境のセットアップ

```bash
# 仮想環境の作成
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
# source .venv/bin/activate

# 依存ライブラリのインストール
pip install -r requirements.txt

# Playwright のブラウザと依存関係をインストール
playwright install --with-deps
```

### 4. 環境変数の設定

`.env_sample` ファイルをコピーして `.env` という名前のファイルを作成します。

```bash
cp .env_sample .env
```

作成した `.env` ファイルを編集し、以下の情報を設定します。

```dotenv
TIMETREE_EMAIL=YOUR_TIMETREE_EMAIL
TIMETREE_PASSWORD=YOUR_TIMETREE_PASSWORD
TIMETREE_CALENDAR_URL=YOUR_TIMETREE_CALENDAR_URL
GAS_WEBAPP_URL=YOUR_GAS_WEBAPP_URL
```

- `TIMETREE_EMAIL`: TimeTree のログインメールアドレス
- `TIMETREE_PASSWORD`: TimeTree のログインパスワード
- `TIMETREE_CALENDAR_URL`: 同期したい TimeTree カレンダーの URL
- `GAS_WEBAPP_URL`: 手順2で取得した GAS のウェブアプリ URL

### 5. GitHub Secrets の設定 (自動実行に必要)

このリポジトリを GitHub にプッシュした後、Actions を正しく動作させるために、以下の情報をリポジトリの Secrets に登録する必要があります。

1.  リポジトリの「Settings」>「Secrets and variables」>「Actions」に移動します。
2.  「New repository secret」をクリックし、以下の4つの Secret を登録します。
    - `TIMETREE_EMAIL`: TimeTree のログインメールアドレス
    - `TIMETREE_PASSWORD`: TimeTree のログインパスワード
    - `TIMETREE_CALENDAR_URL`: 同期したい TimeTree カレンダーの URL
    - `GAS_WEBAPP_URL`: 手順2で取得した GAS のウェブアプリ URL

## 使い方

### 手動実行

ローカル環境でセットアップが完了していれば、以下のコマンドでスクリプトを手動実行できます。

```bash
python timetree_scraper.py
```

### 自動実行

`.github/workflows/main.yml` に定義されたスケジュール (`cron: "0 0 * * *"`、毎日UTCの0時) に基づいて、GitHub Actions が自動的にスクリプトを実行します。
また、リポジトリの「Actions」タブから `workflow_dispatch` を使って手動でワークフローをトリガーすることも可能です。

## 処理の流れ

1.  **GitHub Actions**: スケジュールされた時刻になると、ワークフローが開始されます。
2.  **`timetree_scraper.py`**:
    - Playwright を起動し、環境変数に設定された認証情報を使って TimeTree にログインします。
    - 環境変数 `TIMETREE_CALENDAR_URL` で指定されたカレンダーページから当月のイベント情報（タイトル、日付、時間）をスクレイピングします。
3.  **POST to GAS**:
    - 取得したイベント情報を JSON 形式にまとめ、環境変数 `GAS_WEBAPP_URL` に POST リクエストを送信します。
4.  **`gas.gs` (ウェブアプリ)**:
    - POST リクエストを受け取ります。
    - JSON データを解析し、イベントごとに Google Calendar API を呼び出して、デフォルトカレンダーにイベントを作成します。
    - 終日イベントと時間指定イベントの両方に対応しています。

## ライセンス

This project is licensed under the MIT License.
