---
title: "Streamオブジェクト"
permalink: /stream/
last_modified_at: 2022-11-20T10:38:00+09:00
toc: true
---

## Streamオブジェクトのプロパティとメソッド


|プロパティ|説明|備考|
|---|---|---|
|Charset|保存するときの文字セット||
|EOS|ストリーム位置の末尾であればTrue|True or False|
|LineSeparator|テキストの改行文字||
|Mode|Streamのアクセスモード||
|Position|Stream内の現在位置|先頭は0|
|Size|Streamのバイト数||
|State|Streamの状態||
|Type|Stream内のデータ型|1:バイナリ<br/>2:テキスト|



|メソッド|説明|備考|
|---|---|---|
|Cancel|非同期Streamを停止||
|Close|Streamを閉じる||
|CopyTo|Streamの内容を別のStreamにコピーする||
|Flush|Streamの基になるオブジェクトに書き込む||
|LoadFromFile|ファイルを開く||
|Open|Streamを開く||
|Read|Streamからバイトを読み取る|Type=1のとき|
|ReadText|Streamからテキストを読み取る|Type=2のとき|
|SaveToFile|ファイルに保存する||
|SetEOS|ストリーム位置の末尾を設定する||
|SkipLine|行をスキップする|Type=2のとき|
|Write|バイトをStreamに入力する|Type=1のとき|
|WriteText|テキストをStreamに入力する|Type=2のとき|



## 参考

- [Streamオブジェクト(ADO)(MicroSoft)](https://learn.microsoft.com/ja-jp/sql/ado/reference/ado-api/stream-object-ado?source=recommendations&view=sql-server-ver16)
