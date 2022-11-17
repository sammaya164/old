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
|第2引数|列のデータ型|
|第3引数|文字列のサイズ|省略可|


### テーブルと列を追加する例

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
cols.Append "日付", 7        '日付/時刻型
cols.Append "内容", 203      'メモ型
cols.Append "要約", 202, 40  'テキスト型, 最大40文字
```


### データ型の種類
    
VBAでは「定数」も使えるが、VBScriptでは「値」で指定する。

使用頻度の高いデータ型:

|定数|値|サイズ|説明|
|---|---|---|---|
|adSmallInt|2|2バイト|符号付き整数<br/>VBScriptのIntegerに相当|
|adInteger|3|4バイト|符号付き整数<br/>VBScriptのLongに相当|
|adSIngle|4|4バイト|単精度浮動小数点型|
|adDouble|5|8バイト|倍精度浮動小数点型|
|adCurrency|6|8バイト|通貨型、小数点以下4桁の固定小数点|
|adDate|7|8バイト|日付/時刻型|
|adBoolean|11||ブール値|
|adBigInt|20|8バイト|符号付き整数|
|adVarWChar|202|255バイト|テキスト型、Unicode文字列|
|adLongVarWChar|203|536,870,910バイト|メモ型、長いUnicode文字列|

ほかのデータ型:

|定数|値|サイズ|説明|
|---|---|---|---|
|adEmpty|0||値なし|
|adBSTR|8||Unicode文字列|
|adError|10|4バイト|エラーコード|
|adDecimal|14||正確な数値|
|adTinyInt|16|1バイト|符号付き整数|
|adUnsignedTinyInt|17|1バイト|符号なし整数|
|adUnsignedSmallInt|18|2バイト|符号なし整数|
|adUnsignedInt|19|4バイト|符号なし整数|
|adUnsignedBigInt|21|8バイト|符号なし整数|
|adFileTime|64|8バイト|1601年1月1日からの時間|
|adGUID|72||グローバル一意識別子(GUID)|
|adChar|129|12バイト|文字列、ブランクが補われる|
|adBinary|128||バイナリ値|
|adWChar|130||Unicode文字列|
|adNumeric|131||正確な数値|
|adUserDefined|132||ユーザ定義の変数|
|adDBDate|133|6バイト|日付|
|adDBTime|134|6バイト|時刻|
|adDBTimeStamp|135|16バイト|日付/時刻|
|adChapter|136|4バイト|チャプター値|
|adPropVariant|138||PROPVARIANT|
|adVarNumeric|139||数値(Parameterオブジェクトのみ)|
|adVarChar|200||文字列|
|adLongVarChar|201||長い文字列|
|adVarBinary|204||バイナリ値(Parameterオブジェクトのみ)|
|adLongVarBinary|205||ロングバイナリ値|

情報間違ってたらすいません。


## 参考

- [ADO プログラマのリファレンス トピック (MicroSoft Learn)](https://learn.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/ado-programmer-s-reference-topics)
- [The Connection Strings Reference](https://www.connectionstrings.com/)
