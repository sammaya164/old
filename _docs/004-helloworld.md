---
title: "Hello, World!と表示させる方法"
permalink: /helloworld/
excerpt: ""
last_modified_at: 2021-10-30T00:00:00-02:00
toc: false
---

後に出てくるコードをテキストエディタで打ち込み、hello.vbsの名前で保存する。

以下の3通りの方法で起動した場合の動作を示す。

1. ダブルクリックで起動した場合

1. コマンドプロンプトから`wscript.exe hello.vbs`で起動した場合

1. コマンドプロンプトから`cscript.exe hello.vbs`で起動した場合


## MsgBox関数を使う方法
通常はこれ。

```vb
MsgBox("Hello, World!")
```
1. ダイアログボックスが表示される。
1. ダイアログボックスが表示される。
1. ダイアログボックスが表示される。


## WScriptオブジェクトのEchoメソッドを使う方法
コマンドラインで使いたい場合はこれ。

```vb
WScript.Echo("Hello, World!") 
```
1. ダイアログボックスが表示される。
1. ダイアログボックスが表示される。
1. コマンドラインに表示される。


## 標準出力オブジェクトのWriteLineメソッドを使う方法
コマンドライン専用。自分はほとんど使わない。

```vb
WScript.StdOut.WriteLine("Hello, World!") 
```
1. エラーになる。
1. エラーになる。
1. コマンドラインに表示される。


## ついでにVBAの場合

```vb
Sub Hello()
    Msgbox("Hello, World!")
End Sub
```
標準モジュールなどに書き込む。

`開発`タブの`マクロ`をクリックし、Helloを実行するとダイアログボックスが表示される。


## さらにHTAの場合

```vb
<script type="text/vbscript">
    MsgBox("Hello, World!")
</script>
```
hello.htaの名前で保存し、ダブルクリックするとダイアログボックスが表示される。


