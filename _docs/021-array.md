---
title: "配列の作り方"
permalink: /array/
excerpt: "Array関数"
last_modified_at: 2021-11-01T00:00:00-02:00
toc: false
---

Array関数を使って配列を作成できる。

{% highlight vb %}
Dim fruit

fruit = Array("みかん", "りんご", "メロン")

MsgBox fruit(0) 'みかんと表示される
MsgBox fruit(1) 'りんごと表示される
MsgBox fruit(2) 'メロンと表示される
{% endhighlight %}

`Array()`とすれば要素をもたない配列も作成できる。

Array関数の引数と戻り値：

|引数|説明|
|---|---|
|引数|要素|
|戻り値|配列|


VBScriptの配列関連:

|項目|種類|説明|
|---|---|---|
|Array|関数|配列を作成する|
|Dim|ステートメント|変数を宣言する|
|Erase|ステートメント|配列を初期化する|
|Filter|関数|条件に合致する要素からなる配列を返す|
|For Each ... Next|ステートメント|配列の要素を順番に取得する|
|IsArray|関数|変数が配列か否かをブール値で返す|
|Join|関数|配列の各要素を結合して文字列を返す|
|LBound|関数|配列の添字番号の最小値を返す|
|ReDim|ステートメント|配列の次元やサイズを変更できる|
|Split|関数|文字列から配列を作成する|
|UBound|関数|配列の添字番号の最大値を返す|
