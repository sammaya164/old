---
title: "ADOXの説明"
permalink: /adox/
last_modified_at: 2022-11-15T14:00:00+09:00
toc: true
---


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

### Createの使用例

(その1) Connectionオブジェクトを引数にして呼び出す

```vb
Dim con              'Connectionオブジェクト
Dim cat              'Catalogオブジェクト

Set con = CreateObject("ADODB.Connection")
Set cat = CreateObject("ADOX.Catalog")

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

## Catalog以外のオブジェクト

|オブジェクト|プロパティ、メソッド|
|---|---|
|Column|Name<br/>Type<br/>Attributes<br/>DefinedSize, NumeriScale, Precision, ParentCatalog, RelatedColumn, |
|Group|Name, GetPermission, SetPermission, SortOrder, Properties|
|||
|||
|||
|||
|||
|||
