---
title: "Access形式データベースへの接続1"
permalink: /mdb1/
last_modified_at: 2022-11-06T11:00:00+09:00
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

Accessがインストールされていなくても、VBScriptでMDBファイルを作成することができます。

```vb
'AccessのMDBファイルを作成する

Option Explicit

Dim con 'Connectionオブジェクト
Dim cat 'Catalogオブジェクト
Dim tbl 'Tableオブジェクト
Dim cols 'Columnsオブジェクト

Set con = CreateObject("ADODB.Connection")
Set cat = CreateObject("ADOX.Catalog")
Set tbl = CreateObject("ADOX.Table")

'test.mdbを削除する
'call ClearTestFile("test.mdb") '必要に応じアンコメントする

con.ConnectionString = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=test.mdb;" '接続文字列

cat.Create con 'データベースを作成

tbl.Name = "data" 'テーブル名を設定

cat.Tables.Append tbl 'テーブルを追加

'dataテーブルが存在しているのを確認する
Msgbox GetTableInfo(cat.Tables)

Set cols = cat.Tables("data").Columns

'列を追加する
cols.Append "短整数",   2 'VBScriptのInt型に対応
cols.Append "整数"  ,   3 'VBScriptのLong型に対応
cols.Append "テキスト", 202

'列を削除する
cols.Delete "短整数"

'列を確認する1
Msgbox GetColInfo(cols)

'列を確認する2
Msgbox GetColInfo2(cols)



'ファイルを削除する
Function ClearTestFile(fileName)
    Dim fso
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(fileName) Then
        fso.DeleteFile fileName

    End If

End Function



'Tablesコレクションに含まれるテーブルの名前とタイプを返す
Function GetTableInfo(tables)
    Dim table
    Dim buf

    For Each table In tables
        buf = buf & table.Name & ", " & table.Type & vbCr 

    Next

    GetTableInfo = buf

End Function



'Columnsコレクションに含まれる列の名前とタイプを返す(その1)
Function GetColInfo(cols)
    Dim col
    Dim buf

    For Each col In cols
        buf = buf & col.Name & ", タイプ: " & col.Type & vbCr 

    Next

    GetColInfo = buf

End Function



'Columnsコレクションに含まれる列の名前とタイプを返す(その2)
Function GetColInfo2(cols)
    Dim i
    Dim buf

    For i = 0 To cols.Count - 1
        buf = buf & cols(i).Name & ", タイプ: " & cols(i).Type & vbCr 

    Next

    GetColInfo2 = buf

End Function
```


64bitのWindowsでは、上記スクリプトファイルをWクリックで起動してもConnectionオブジェクトが動かないのでエラーになります。

32bit版のWSHから起動するために、下記スクリプトファイルを作成し、これに上記スクリプトファイルをドラッグ&ドロップすると正常に動作するはずです。

どちらもShift_JISの文字コードで保存してください。


```vb
'このスクリプトファイルにドラッグ&ドロップしたスクリプトファイルを32bit版のWSHで実行する

Option Explicit

Dim shell 'WshShellオブジェクト
Dim fso   'FileSystemObjectオブジェクト
Dim path


If WScript.Arguments.Count = 1 Then '1ファイルをドラッグ＆ドロップしたとき
    Set shell = CreateObject("WScript.Shell")
    Set fso   = CreateObject("Scripting.FileSystemObject")

    'ドラッグ＆ドロップしたスクリプトファイルのパス
    path = WScript.Arguments(0)

    '作業フォルダとしてpathの親フォルダを指定
    shell.CurrentDirectory = fso.GetParentFolderName(path)

    '32bit版WSHでスクリプトを実行する
    shell.Run "C:\windows\SysWOW64\wscript.exe """ & path & """"

End If
```

スクリプトファイルと同じフォルダ内にtest.mdbが作成されたと思います。


## 参考

- [ADO プログラマのリファレンス トピック (MicroSoft Learn)](https://learn.microsoft.com/ja-jp/office/client-developer/access/desktop-database-reference/ado-programmer-s-reference-topics)
- [The Connection Strings Reference](https://www.connectionstrings.com/)
- [Microsoft.ACE.OLEDBについてまとめてみた](https://qiita.com/yaju/items/7b0aa9e9f30005f60388) 
