---
title: "ADOXの説明"
permalink: /adox/
last_modified_at: 2022-11-17T10:00:00+09:00
toc: true
---

## ADOXのオブジェクト

|オブジェクト|プロパティ|メソッド|
|---|---|---|
|Catalog|ActiveConnection|Create<br/>GetObjectOwner<br/>SetObjectOwner|
|Column|Name<br/>Type<br/>Attributes<br/>DefinedSize<br/>NumeriScale<br/>Precision<br/>ParentCatalog<br/>RelatedColumn<br/>SortOrder<br/>Properties||
|Group|Name<br/>Users<br/>Properties|GetPermission<br/>SetPermission|
|Index|Name<br/>Columns<br/>Unique<br/>PrimaryKeys<br/>IndexNulls<br/>Clustered<br/>Properties||
|Key|Name<br/>Type<br/>Columns<br/>RelatedTable<br/>DeleteRule><br/>UpdateRule||
|Procedure|Name<br/>Command<br/>DateCreated<br/>DateModified||
|Table|Name<br/>Type<br/>Columns<br/>Indexes<br/>Keys<br/>ParentCatalog<br/>DateCreated<br/>DateModified<br/>Properties||
|User|Name<br/>Groups<br/>Properties|ChangePassword<br/>GetPermission<br/>SetPermission|
|View|Name<br/>Command<br/>DateCreated<br/>DateModified||


## Catalogオブジェクト


### プロパティとメソッド

|プロパティ|説明|
|---|---|
|ActiveConnection|接続文字列またはConnectionオブジェクト|

|メソッド|説明|
|---|---|
|Create|新しいカタログを作成する|
|GetObjectOwner|オブジェクトの所有者を取得する|
|SetObjectOwner|オブジェクトの所有者を設定する|


### Createメソッドの使用例

(その1) Connectionオブジェクトを引数にして呼び出す

```vb
Dim cat              'Catalogオブジェクト
Dim con              'Connectionオブジェクト

Set cat = CreateObject("ADOX.Catalog")
Set con = CreateObject("ADODB.Connection")

con.ConnectionString = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=test.mdb;"
cat.Create con 'データベースを作成

```

(その2) 接続文字列を引数にして呼び出す

```vb
Dim cat              'Catalogオブジェクト

Set cat = CreateObject("ADOX.Catalog")

cat.Create "Provider=Microsoft.JET.OLEDB.4.0;Data Source=test.mdb;" 'データベースを作成

```

後者の例ではConnectionオブジェクトが表に登場してませんが、裏では作成されており、Createメソッドの戻り値として取得できます。

```vb
Dim cat              'Catalogオブジェクト
Dim con              'Connectionオブジェクト

Set cat = CreateObject("ADOX.Catalog")
Set con = cat.Create("Provider=Microsoft.JET.OLEDB.4.0;Data Source=test.mdb;") 'データベースを作成

MsgBox TypeName(con) 'Connectionと表示される
con.Close '接続を閉じる
```

## ADOXのコレクション

|コレクション|プロパティ|メソッド|
|---|---|---|
|Columns|Item<br/>Count|Append<br/>Delete<br/>Refresh|
|Groups|Item<br/>Count|Append<br/>Delete<br/>Refresh|
|Indexes|Item<br/>Count|Append<br/>Delete<br/>Refresh|
|Keys|Item<br/>Count|Append<br/>Delete<br/>Refresh|
|Procedures|Item<br/>Count|Append<br/>Delete<br/>Refresh|
|Tables|Item<br/>Count|Append<br/>Delete<br/>Refresh|
|Users|Item<br/>Count|Append<br/>Delete<br/>Refresh|
|Views|Item<br/>Count|Append<br/>Delete<br/>Refresh|


## Tablesコレクション


### プロパティとメソッド

|プロパティ|説明|
|---|---|
|Item|引数で指定したメンバー(テーブル)を示す|
|Count|メンバー(テーブル)の数|

|メソッド|説明|
|---|---|
|Append|メンバー(テーブル)を追加する|
|Delete|メンバー(テーブル)を削除する|
|Refresh|変更内容を反映する|


## Columnsコレクション


### プロパティとメソッド

|プロパティ|説明|
|---|---|
|Item|引数で指定したメンバー(列)を示す|
|Count|メンバー(列)の数|

|メソッド|説明|
|---|---|
|Append|メンバー(列)を追加する|
|Delete|メンバー(列)を削除する|
|Refresh|変更内容を反映する|


### Appendメソッドの引数

|引数|説明|備考|
|---|---|---|
|第1引数|列の名前||
|第2引数|データ型||
|第3引数|文字列、バイナリのサイズを指定||


## テーブルと列を追加する例

```vb
Dim cat  'Catalogオブジェクト
Dim cols 'Columnsコレクション
Dim tbl  'Tableオブジェクト

Set cat = CreateObject("ADOX.Catalog")
Set tbl = CreateObject("ADOX.Table")

cat.ActiveConnection = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=test.mdb;" 'データベースに接続
tbl.Name = "日記"

cat.Tables.Append tbl '日記テーブルを追加

Set cols = cat.Tables("日記").Columns '日記テーブルのColumnsコレクションを取得

'列を追加する
cols.Append "日付", 7         '日付/時刻型
cols.Append "内容", 203, 1000 '文字列, サイズ1000
```


## 参考

- [ADO プログラマのリファレンス トピック (MicroSoft Learn)](https://learn.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/ado-programmer-s-reference-topics)
