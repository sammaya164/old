---
title: "Access形式データベースへの接続1"
permalink: /oledb/
last_modified_at: 2022-11-03T17:00:00+09:00
toc: true
---


## OLEDBプロバイダーを確認する

自分の環境はOS:Windows11で64bit版Officeをインストール済みです. 

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

**ヒント:** 32bit版PowerShellは C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe にあります。
{: .notice--info}

つぎに32bit版PowerShellを起動して同じコマンドを実行します。

32bit版のOLEDBが表示されます. 必要な行だけ示してます。

```shell
SOURCES_NAME               SOURCES_DESCRIPTION
------------               -------------------
Microsoft.Jet.OLEDB.4.0    Microsoft Jet 4.0 OLE DB Provider
```

以上から自分の環境では64bit版の`Microsoft.ACE.OLEDB.12.0`、`Microsoft.ACE.OLEDB.16.0`と
32bit版の`Microsoft.Jet.OLEDB.4.0`が利用できることが分かります。
{: .text-justify}

**ヒント:** 32bit版のOfficeをインストールしている環境(会社とか)では、32bit版の`Microsoft.Jet.OLEDB.4.0`、
`Microsoft.ACE.OLEDB.12.0`、`Microsoft.ACE.OLEDB.16.0`が利用できると思います。
{: .notice--info}

VBScriptを動かすWSHにも32bit版と64bit版があります。

一見、32bit版WSHからは`Microsoft.Jet.OLEDB.4.0`を、64bit版WSHからは`Microsoft.ACE.OLEDB.12.0`、
`Microsoft.ACE.OLEDB.16.0`を利用できそうに思います。

ところが、64bit版WSHからはのADODB.Connectionオブジェクトを使用できないらしく(VBAからは可能)、
自分の環境でAccess形式のデータベースに接続するには、32bit版WSHから`Microsoft.Jet.OLEDB.4.0`を利用することになります。  
(勉強が目的なので、それで問題ありません。)

## 参考

- [ADO プログラマのリファレンス トピック (MicroSoft Learn)](https://learn.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/ado-programmer-s-reference-topics)
- [The Connection Strings Reference](https://www.connectionstrings.com/)
- [Microsoft.ACE.OLEDBについてまとめてみた](https://qiita.com/yaju/items/7b0aa9e9f30005f60388) 
