# 勤務表自動入力システム

Outlookの送信済みメール（始業時報告・終業時報告）を読み取り、勤務表Excelへ自動入力するPythonスクリプトです。

## 機能

- **始業時報告メール** → 始業時間・勤務形態（在宅/出社）を取得
- **終業時報告メール** → 終業時間を取得
- 取得したデータを勤務表Excelの対応セルへ自動入力
- **自分の送信メールのみ**処理（他の人の同形式メールは除外）
- Excelファイルはダイアログで選択（パスは自動保存）
- プレビューモード対応（確認してから書き込み）

## 動作環境

| 項目 | 内容 |
|------|------|
| OS | Windows 10 / 11 |
| Python | 3.8 以上 |
| Outlook | Microsoft Outlook（デスクトップ版） |

## セットアップ（初回のみ）

### 方法① `setup.bat` を実行（簡単）

```
setup.bat をダブルクリック
```

ライブラリのインストールと `config.ini` の作成が自動で行われます。

### 方法② 手動

```bash
# ① リポジトリをクローン
git clone https://github.com/murakijun/kintai-auto-input.git
cd kintai-auto-input

# ② ライブラリをインストール
pip install -r requirements.txt

# ③ 設定ファイルを作成
copy config.ini.example config.ini
```

## 設定

`config.ini` をテキストエディタで開いて編集します。

```ini
[user]
sender_email = your-email@example.com   # ★ 自分のメールアドレスを設定

[target]
year  = 2026
month = 4                               # ★ 処理対象月を毎月変更

[excel]
path =                                  # 空白 → ダイアログで選択（自動保存）

[options]
preview_only = false                    # true で書き込まず確認だけ
```

> **注意:** `config.ini` には個人情報（メールアドレス・ファイルパス）が含まれるため `.gitignore` で管理外にしています。

## 使い方

```bash
# Outlookを起動した状態で実行
python 勤務表自動入力.py
```

初回はExcelファイルの選択ダイアログが表示されます。
選択したパスは `config.ini` に自動保存されるため、次回以降は不要です。

### 月が変わったとき

`config.ini` の `month` を変更するだけ：

```ini
[target]
month = 5   ← 変更
```

## 対応メール形式

### 始業時報告（件名に「始業時報告」を含む）

```
始業時間：09:00
在宅/出社勤務：在宅勤務
```

### 終業時報告（件名に「終業時報告」を含む）

```
始業時間：09:00
終業時間：17:30
```

## Excel書き込み先

| 列 | 内容 |
|----|------|
| C列 | 始業時間 |
| D列 | 終業時間 |
| P列 | 備考（勤務形態） |

## ファイル構成

```
kintai-auto-input/
├── 勤務表自動入力.py    メインスクリプト
├── config.ini.example   設定ファイルのテンプレート
├── requirements.txt     必要ライブラリ一覧
├── setup.bat            Windowsセットアップスクリプト
├── 設計書.md            システム設計書
└── README.md            この説明書
```

## トラブルシューティング

| 症状 | 対処 |
|------|------|
| `pywin32 がない` | `pip install pywin32` を実行 |
| メールが見つからない | Outlookが起動しているか確認。`config.ini` の `year/month` を確認 |
| Excelが開けない | 別のアプリでExcelを開いていないか確認 |
| 他の人のメールも処理される | `config.ini` の `sender_email` を設定する |
