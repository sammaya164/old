---
title: "ADOのデータ型"
permalink: /fldtype/
last_modified_at: 2022-11-17T15:00:00+09:00
toc: true
---

## データ型の種類
    
使用頻度の高いデータ型:

|定数|値|サイズ|説明|
|---|---|---|---|
|adInteger|3|4バイト|符号付き整数<br/>VBScriptのLongに相当|
|adDouble|5|8バイト|倍精度浮動小数点型|
|adDate|7|8バイト|日付/時刻型|
|adBoolean|11|2バイト|ブール値|
|adVarWChar|202||Unicode文字列|
|adLongVarWChar|203||長いUnicode文字列|

ほかのデータ型:

|定数|値|サイズ|説明|
|---|---|---|---|
|adEmpty|0||空|
|adSmallInt|2|2バイト|符号付き整数<br/>VBScriptのIntegerに相当|
|adSIngle|4|4バイト|単精度浮動小数点型|
|adCurrency|6|8バイト|通貨型、小数点以下4桁の固定小数点|
|adBSTR|8||Unicode文字列|
|adError|10|4バイト|エラーコード|
|adDecimal|14|16バイト|正確な数値|
|adTinyInt|16|1バイト|符号付き整数|
|adUnsignedTinyInt|17|1バイト|符号なし整数|
|adUnsignedSmallInt|18|2バイト|符号なし整数|
|adUnsignedInt|19|4バイト|符号なし整数|
|adBigInt|20|8バイト|符号付き整数|
|adUnsignedBigInt|21|8バイト|符号なし整数|
|adFileTime|64|8バイト|1601年1月1日からの時間|
|adGUID|72|16バイト|グローバル一意識別子(GUID)|
|adBinary|128||バイナリ値|
|adChar|129||文字列|
|adWChar|130||Unicode文字列|
|adNumeric|131|19バイト|正確な数値|
|adUserDefined|132||ユーザ定義の変数|
|adDBDate|133|6バイト|日付|
|adDBTime|134|6バイト|時刻|
|adDBTimeStamp|135|16バイト|日付/時刻|
|adChapter|136||チャプター値|
|adPropVariant|138||PROPVARIANT|
|adVarNumeric|139||数値(Parameterオブジェクトのみ)|
|adVarChar|200||文字列|
|adLongVarChar|201||長い文字列|
|adVarBinary|204||バイナリ値(Parameterオブジェクトのみ)|
|adLongVarBinary|205||ロングバイナリ値|


## データ型のDefinedSizeを確認してみる

```vb
Dim rs     'レコードセットオブジェクト
Dim flds   'フィールドコレクション

Set rs   = CreateObject("ADODB.RecordSet")
Set flds = rs.Fields

flds.Append "SmallInt", 2
flds.Append "Integer" , 3
flds.Append "Single"  , 4
flds.Append "Double"  , 5
flds.Append "Currency", 6
flds.Append "Date"    , 7
flds.Append "BSTR", 8
flds.Append "Error", 10
flds.Append "Boolean" , 11
flds.Append "Decimal", 14
flds.Append "TinyInt", 16
flds.Append "UnsignedTinyInt", 17
flds.Append "UnsignedSmallInt", 18
flds.Append "UnsignedInt", 19
flds.Append "BigInt"  , 20
flds.Append "UnsignedBigInt", 21
flds.Append "FileTime", 64
flds.Append "GUID", 72
flds.Append "Numeric", 131
flds.Append "DBDate", 133
flds.Append "DBTime", 134
flds.Append "DBTimeStamp", 135

rs.Open
rs.AddNew
rs.Update

Dim fld
Dim buf

For Each fld In flds
    buf = buf & fld.Name & ", " & fld.DefinedSize & vbCr
Next

rs.Close

Msgbox buf
```


### 実行例

![fldtype](/vbscript/assets/images/fldType.png)
