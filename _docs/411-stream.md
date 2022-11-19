---
title: "Streamオブジェクト"
permalink: /stream/
last_modified_at: 2022-11-19T18:00:00+09:00
toc: true
---

## Streamオブジェクトのプロパティとメソッド


|プロパティ|説明|
|---|---|
|Charset|文字セット|
|EOS|ストリーム位置の末尾であればTrue|
|LineSeparator|テキストの改行文字|
|Mode|Streamのアクセスモード|
|Position|Stream内の現在位置|
|Size|Streamのバイト数|
|State|Streamの状態|
|Type|Stream内のデータ型|



|メソッド|説明|
|---|---|
|Cancel|非同期Streamを停止|
|LoadFromFile|ファイルを開く|
|Open|Streamを開く|
|Close|Streamを閉じる|
|Write|バイトをStreamに入力する|
|WriteText|テキストをStreamに入力する|
|Read|Streamからバイトを読み取る|
|ReadText|Streamからテキストを読み取る|
|Flush|Streamの基になるオブジェクトに書き込む|
|CopyTo|Streamの内容を別のStreamにコピーする|
|SaveToFile|ファイルに保存する|
|SetEOS|ストリーム位置の末尾を設定する|
|SkipLine|行をスキップする|



## 参考

- [Streamオブジェクト(ADO)(MicroSoft)](https://learn.microsoft.com/ja-jp/sql/ado/reference/ado-api/stream-object-ado?source=recommendations&view=sql-server-ver16)
