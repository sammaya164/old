---
title: "VBScriptの暗号化"
permalink: /encoder/
excerpt: "scrrun.dllのEncoderオブジェクトを利用してVBScriptを暗号化するスクリプトを作成できる"
last_modified_at: 2021-10-30T00:00:00-02:00
toc: false

---

VBScriptを個人で使用する場合は問題無いが、組織内の不特定多数で使用する場合にはコードを改変される可能性がある。
改変されると最悪の場合、PCのシステムファイルを破壊したり、機密情報を外部へ送信したりと有害な動作をする可能性がある。

コードの改変を簡単にできないようにする手段として、コードを暗号化するという選択肢がある。

以前は Windows Script Encoder というツールがマイクロソフトから提供されていたが、現在はWindow OSに標準で暗号化機能が用意されている。
scrrun.dllのEncoderオブジェクトを利用してVBScriptを暗号化するスクリプトを作成できる。

次のコードはhttps://gallery.technet.microsoft.com/scriptcenter/16439c02-3296-4ec8-9134-6eb6fb599880からの転載。
2020年にTechNetギャラリーが廃止された為、転載元は無くなっている。

{% highlight vb %}
Option Explicit 
 
dim oEncoder, oFilesToEncode, file, sDest 
dim sFileOut, oFile, oEncFile, oFSO, i 
dim oStream, sSourceFile 
 
set oFilesToEncode = WScript.Arguments 
set oEncoder = CreateObject("Scripting.Encoder") 
For i = 0 to oFilesToEncode.Count - 1 
    set oFSO = CreateObject("Scripting.FileSystemObject") 
    file = oFilesToEncode(i) 
    set oFile = oFSO.GetFile(file) 
    set oStream = oFile.OpenAsTextStream(1) 
    sSourceFile=oStream.ReadAll 
    oStream.Close 
    sDest = oEncoder.EncodeScriptFile(".vbs",sSourceFile,0,"") 
    sFileOut = Left(file, Len(file) - 3) & "vbe" 
    set oEncFile = oFSO.CreateTextFile(sFileOut) 
    oEncFile.Write sDest 
    oEncFile.Close 
Next 
{% endhighlight %}

1. このコードをencode.vbsなどの名前で保存する。
1. 暗号化したいvbsファイルをencode.vbsにドラッグ＆ドロップすると、暗号化されて拡張子vbeのファイルが作成される。
1. 複数のvbsファイルをまとめてドラッグ＆ドロップすることもできる。

暗号化したコードを復号化するツールも存在するらしいので、コードの改変を完全に防げるわけではない。