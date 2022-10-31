---
title: "配列を初期化する"
permalink: /erase/
last_modified_at: 2021-11-02T00:00:00-02:00
toc: false
---

Eraseステートメントを使用して配列を初期化できる。
文字列は""へ、数値は0へ初期化される。

{% highlight vb %}
Dim fruit(2)

fruit(0) = "みかん"
fruit(1) = "りんご"
fruit(2) = "メロン"

MsgBox Join(fruit, ",") 'みかん,りんご,メロンと表示される
Erase fruit '配列を初期化
MsgBox Join(fruit, ",") ',,と表示される
{% endhighlight %}
