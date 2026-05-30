# トランスクリプト議事録メーカー

トランスクリプト文章を貼り付け、`.txt` ファイルを選択、またはドラッグ＆ドロップして、Excel形式の議事録（`.xlsx`）を作成するPythonプログラムです。

## セットアップ

```bash
python -m venv .venv
source .venv/bin/activate
# 任意: ドラッグ＆ドロップ機能やテスト用依存を入れる場合
pip install -r requirements.txt
```

> Excel生成は標準ライブラリのみで動作します。ドラッグ＆ドロップ機能には `tkinterdnd2` を使います。未インストールでも、ファイル選択と貼り付け入力は利用できます。

## GUIで使う

```bash
python app.py
```

1. トランスクリプト文章を貼り付けるか、`.txt` ファイルをドラッグ＆ドロップします。
2. 出力先の `.xlsx` ファイルを選択します。
3. **Excel議事録を作成** をクリックします。

## コマンドラインで使う

```bash
python minutes_maker.py transcript.txt minutes.xlsx
```

## 自動抽出する内容

- 会議名
- 開催日
- 参加者
- 議題
- 要約
- 決定事項
- アクションアイテム（担当・内容・期限）

抽出はルールベースのため、生成後にExcelを開いて内容を確認・調整してください。
