# Extract Leaders

部員名簿から役職者を抽出する。

## 使い方
1. `main.py`と同じ階層に、`member_list`というディレクトリを作る。
2. 1で作ったディレクトリの中に、識別番号付きの部員名簿を全団体入れる。
3. `main.py`と同じ階層に、「識別番号,団体名,所属,格」の順で団体情報を記述したファイルを`group_list.xlsx`というファイル名で
4. `main.py`を引数なしで実行する。
5. 抽出された役職者の一覧が`output_file.xlsx`に出力される。

`group_list.xlsx`は、共有フォルダにあります。詳しい場所は聞いてください。または作成してください。

## 製作者
Sou Tamura([tmsou0209](https://github.com/tmsou0209))

## ライセンス
MITライセンス