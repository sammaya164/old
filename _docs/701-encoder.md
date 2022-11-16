---
title: "VBScriptの暗号化"
permalink: /encoder/
last_modified_at: 2022-11-16T14:00:00+09:00
toc: true
---

## 個人的動機

VBScriptを組織内の不特定多数で使用する場合にはコードを改変される可能性がある。

改変されると最悪の場合、PCのシステムファイルを破壊したり、機密情報を外部へ送信したりと有害な動作をする可能性がある。

コードの改変を簡単にできないようにする手段として、コードを暗号化するという選択肢がある。

## 暗号化のスクリプト

次のコードはhttps://gallery.technet.microsoft.com/scriptcenter/16439c02-3296-4ec8-9134-6eb6fb599880
からの転載である。

dim->Dim、set->Setなど元の記述から小文字/大文字を一部変更した。
{: .notice--info}

2020年にTechNetギャラリーが廃止された為、転載元は無くなっている。
{: .notice--info}

```vb
Option Explicit 
 
Dim oEncoder, oFilesToEncode, file, sDest 
Dim sFileOut, oFile, oEncFile, oFSO, i 
Dim oStream, sSourceFile 
 
Set oFilesToEncode = WScript.Arguments 
Set oEncoder = CreateObject("Scripting.Encoder") 
For i = 0 to oFilesToEncode.Count - 1 
    Set oFSO = CreateObject("Scripting.FileSystemObject") 
    file = oFilesToEncode(i) 
    Set oFile = oFSO.GetFile(file) 
    Set oStream = oFile.OpenAsTextStream(1) 
    sSourceFile=oStream.ReadAll 
    oStream.Close 
    sDest = oEncoder.EncodeScriptFile(".vbs",sSourceFile,0,"") 
    sFileOut = Left(file, Len(file) - 3) & "vbe" 
    Set oEncFile = oFSO.CreateTextFile(sFileOut) 
    oEncFile.Write sDest 
    oEncFile.Close 
Next 
```

## 使い方

1. このコードをencode.vbsなどの名前で保存する。
1. 暗号化したいvbsファイルをencode.vbsにドラッグ＆ドロップすると、暗号化されて拡張子vbeのファイルが作成される。
1. 複数のvbsファイルをまとめてドラッグ＆ドロップすることもできる。

暗号化したコードを復号化するツールも存在するらしいので、コードの改変を完全に防げるわけではない。
{: .notice--info}

## EncoderオブジェクトのEncodeScriptFileメソッド

||説明|値|
|---|---|---|
|第1引数|拡張子|".vbs"<br/>".js"<br/>"html"<br/>"htm"<br/>他|
|第2引数|暗号化前のテキスト||
|第3引数|フラグ?|0|
|第4引数|デフォルト言語|""(これで良い)<br/>"VBScript"<br/>"JScript"|
|戻り値|暗号化後のテキスト||


前掲のスクリプトを少し変更すれば、JScriptやHTMLも暗号化できそう。
