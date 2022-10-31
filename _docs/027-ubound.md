---
title: "LBound関数、UBound関数"
permalink: /ubound/
last_modified_at: 2021-11-02T00:00:00-02:00
toc: false
---

LBound関数を使って配列の添字番号の最小値を取得できる。  
UBound関数を使って配列の添字番号の最大値を取得できる。  

LBound関数の引数と戻り値:

|引数|説明|必須/省略可|
|---|---|---|
|第1引数|配列|必須|
|第2引数|次元を指定(デフォルトは1)|省略可|
|戻り値|添字番号の最小値|-|

UBound関数の引数と戻り値:

|引数|説明|必須/省略可|
|---|---|---|
|第1引数|配列|必須|
|第2引数|次元を指定(デフォルトは1)|省略可|
|戻り値|添字番号の最大値|-|

1次元配列の場合:

{% highlight vb %}
Dim fruit(2)

MsgBox LBound(fruit) '0と表示される
MsgBox UBound(fruit) '2と表示される
{% endhighlight %}

2次元配列の場合:

{% highlight vb %}
Dim fruit(3,7)

MsgBox LBound(fruit, 1) '0と表示される
MsgBox UBound(fruit, 1) '3と表示される
MsgBox LBound(fruit, 2) '0と表示される
MsgBox UBound(fruit, 2) '7と表示される
{% endhighlight %}
