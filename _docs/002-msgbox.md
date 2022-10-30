---
title: "MsgBox"
permalink: /msgbox/
excerpt: "VBScriptの標準的な出力機能としてMsgBox関数がある"
last_modified_at: 2021-10-30T01:00:00-02:00
---

VBScriptの標準的な出力機能としてMsgBox関数がある

MsgBox関数の引数と戻り値:

|引数|説明|必須/省略可|
|---|---|---|
|第1引数|メッセージ|必須|
|第2引数|ボタンの種類|省略可|
|第3引数|タイトル|省略可|
|第4引数|ヘルプファイル|省略可|
|第5引数|コンテキスト|省略可|
|戻り値|押されたボタン|-|

MsgBox関数を使って`Hello`と表示させるには、以下のように数通りの書き方がある

{% highlight vb %}
MsgBox "Hello"
{% endhighlight %}

{% highlight vb %}
MsgBox("Hello")
{% endhighlight %}

{% highlight vb %}
Call MsgBox("Hello")
{% endhighlight %}

第2引数でボタンの種類を変更できる。

{% highlight vb %}
MsgBox "Hello", vbOKCancel
{% endhighlight %}

{% highlight vb %}
Call MsgBox("Hello", vbOKCancel)
{% endhighlight %}

複数の引数を()で囲む場合、Callを付けるか変数に代入する必要がある。
次のように書くとエラーになる。

{% highlight vb %}
MsgBox("Hello", vbOKCancel)
{% endhighlight %}

ボタンの種類:

|種類|定数|値|
|---|---|---|
|OK|vbOKOnly|0|
|OK、キャンセル|vbOKCancel|1|
|中止、再試行、無視|vbAbortRetryIgnore|2|
|はい、いいえ、キャンセル|vbYesNoCancel|3|
|はい、いいえ|vbYesNo|4|
|再試行、キャンセル|vbRetryCancel|5|
|警告アイコンを表示|vbCritical|16|
|問い合わせアイコンを表示|vbQuestion|32|
|注意アイコンを表示|vbInformation|48|
|情報アイコンを表示|vbExclamation|64|
|第1ボタンを標準|vbDefaultButton1|0|
|第2ボタンを標準|vbDefaultButton2|256|
|第3ボタンを標準|vbDefaultButton3|512|
|第4ボタンを標準|vbDefaultButton4|768|

ボタンの種類は値の足し算で組合せが可能。

{% highlight vb %}
MsgBox "問題が発生しました", vbRetryCancel + vbCritical
{% endhighlight %}

押されたボタンを取得するには、次のように返り値を調べる。

{% highlight vb %}
val = MsgBox("実行しますか？", vbYesNo)
If val = vbYes Then
    'はいの場合の動作
End If
{% endhighlight %}

変数に代入する場合、引数を()で囲む必要がある。
次のように書くとエラーになる。

{% highlight vb %}
val = MsgBox "Hello"
{% endhighlight %}

押されたボタン:

|ボタン|定数|値|
|---|---|---|
|OK|vbOK|1|
|キャンセル|vbCancel|2|
|中止|vbAbort|3|
|再試行|vbRetry|4|
|無視|vbIgnore|5|
|はい|vbYes|6|
|いいえ|vbNo|7|

Select Caseステートメントと組み合わせるとすっきり書ける場合が多い。

{% highlight vb %}
Select Case MsgBox("実行しますか？", vbYesNoCancel)
Case vbYes
    'はいの場合の動作
Case vbNo
    'いいえの場合の動作
Case vbCancel
    'キャンセルの場合の動作
End Select
{% endhighlight %}
