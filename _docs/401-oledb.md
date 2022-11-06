---
title: "Access形式データベースへの接続1"
permalink: /oledb/
last_modified_at: 2022-11-03T17:00:00+09:00
toc: true
---

## OLEDBプロバイダーを確認する

VBScriptでデータベースへ接続するための事前準備として、利用可能なOLEDBプロバイダーを確認します。

PowerShellで以下のコマンドを実行します。

```shell
(New-Object data.oledb.oledbenumerator).getElements() | select SOURCES_NAME, SOURCES_DESCRIPTION
```

64bit版のOLEDBが表示されます。必要な行だけ示してます。 

```shell
SOURCES_NAME               SOURCES_DESCRIPTION
------------               -------------------
Microsoft.ACE.OLEDB.12.0   Microsoft Office 12.0 Access Database Engine OLE DB Provider
Microsoft.ACE.OLEDB.16.0   Microsoft Office 16.0 Access Database Engine OLE DB Provider
```

つぎに32bit版PowerShellを起動して同じコマンドを実行します。

**ヒント:** 32bit版PowerShellは `C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe` にあります。
{: .notice--info}

32bit版のOLEDBが表示されます. 必要な行だけ示してます。

```shell
SOURCES_NAME               SOURCES_DESCRIPTION
------------               -------------------
Microsoft.Jet.OLEDB.4.0    Microsoft Jet 4.0 OLE DB Provider
```

以上から自分の環境では64bit版の「Microsoft.ACE.OLEDB.12.0」、「Microsoft.ACE.OLEDB.16.0」と
32bit版の「Microsoft.Jet.OLEDB.4.0」が利用できることが分かります。

自分の環境は64bit版Officeをインストール済みなのでこのような構成になっているものと思われます。

VBScriptを動かしているWSHにも32bit版と64bit版があります。
しかし、64bit版WSHからはデータベース接続に使うADODB.Connectionオブジェクトを使用できないようです。(VBAからは可能)

したがって自分の環境でAccess形式データベースへ接続するには、32bit版WSHから`Microsoft.Jet.OLEDB.4.0`を利用することになります。

## MDBファイルを作成する

Accessがインストールされていなくても、VBScriptからMDBファイルを作成し、使用することができます。


```vb
Dim con 'Connectionオブジェクト
Dim cat 'Catalogオブジェクト
Dim tbl 'Tableオブジェクト

Set con = CreateObject("ADODB.Connection")
Set cat = CreateObject("ADOX.Catalog")
Set tbl = CreateObject("ADOX.Table")

con.ConnectionString = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=test.mdb;" '接続文字列

Msgbox con.Mode '

cat.Create con 'データベースを作成

Msgbox con.Mode '

tbl.Name = "data" 'テーブル名を設定

cat.Tables.Append tbl 'テーブルを追加

con.Close



```


## 参考

- [ADO プログラマのリファレンス トピック (MicroSoft Learn)](https://learn.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/ado-programmer-s-reference-topics)
- [The Connection Strings Reference](https://www.connectionstrings.com/)
- [Microsoft.ACE.OLEDBについてまとめてみた](https://qiita.com/yaju/items/7b0aa9e9f30005f60388) 
