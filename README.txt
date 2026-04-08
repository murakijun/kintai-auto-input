========================================
  勤務表自動入力スクリプト v2.0
  使い方
========================================

【初回セットアップ（1回だけ）】

  ① Pythonライブラリをインストール
       pip install pywin32 openpyxl

  ② config.ini を開いて自分のメールアドレスを設定
       sender_email = yourname@example.com

  以上で準備完了！


【毎月の使い方】

  ① config.ini を開いて対象月を変更
       year  = 2026
       month = 5    ← ここを変える

  ② Outlookを起動した状態でスクリプトを実行
       python 勤務表自動入力.py

  ③ 初回はファイル選択ダイアログが開くので
     勤務表Excelファイルを選択する
     （2回目以降は自動でそのファイルを使用）

  ④ 結果を確認して完了！


【まず確認だけしたい場合】

  config.ini の以下を変更してから実行:
    preview_only = true
  → Excelに書き込まず、変更内容だけ表示します
  → 問題なければ false に戻して再実行


【ファイル説明】

  勤務表自動入力.py  メインスクリプト（編集不要）
  config.ini         設定ファイル（★ここだけ編集）
  設計書.md          システムの詳細設計書
  README.txt         この説明書


【config.ini の設定項目】

  [user]
  sender_email = yourname@example.com
    自分のOutlookメールアドレス。
    他の人が送る同じフォーマットのメールを除外するために使用。
    ※ 未設定でも Outlook から自動取得を試みます。

  [target]
  year  = 2026    処理対象年
  month = 4       処理対象月

  [excel]
  path =
    勤務表ExcelファイルのパスZ
    空のままにするとダイアログが開きます。
    一度選択すると自動保存されます。

  [outlook]
  folder_type = 5
    5 = 送信済みアイテム（通常はこちら）
    6 = 受信トレイ

  [options]
  preview_only = false
    true  = 確認だけ（書き込まない）
    false = 実際に書き込む


【処理するメールの形式】

  件名「始業時報告(在宅)」などを含むメール：
    始業時間：09:00
    在宅/出社勤務：在宅勤務
    → C列（始業時間）と P列（勤務形態）に入力

  件名「終業時報告」を含むメール：
    終業時間：17:30
    → D列（終業時間）に入力


【トラブルシューティング】

  「pywin32がインストールされていません」
    → pip install pywin32

  「メールが見つかりません」
    → Outlookが起動しているか確認
    → config.ini の year / month を確認

  「Excelファイルが自動検出されない」
    → ダイアログでファイルを選択（自動保存されます）
    → または config.ini の path に直接記入
