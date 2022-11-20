---
title: "バイナリファイルを読む"
permalink: /mnist/
last_modified_at: 2022-11-19T21:00:00+09:00
toc: true
---


## バイナリファイルの入手

ニューラルネットワークの訓練と評価に使用されている手書き数字画像の、MNISTデータベースを[ここ](http://yann.lecun.com/exdb/mnist/)からダウンロードします。

- train-images-idx3-ubyte.gz
- train-labels-idx1-ubyte.gz
- t10k-images-idx3-ubyte.gz
- t10k-labels-idx1-ubyte.gz

圧縮されているので7-Zipなどを使って展開すると、下記ファイルになります。

|ファイル名|説明|Byte|
|---|---|---|
|train-images.idx3-ubyte|60,000枚分の訓練用画像|47,040,016|
|train-labels.idx1-ubyte|訓練用画像の正解ラベル|60,008|
|t10k-images.idx3-ubyte|10,000枚分の評価用画像|7,840,016|
|t10k-labels.idx1-ubyte|評価用画像の正解ラベル|10,008|

これらはバイナリファイルです。1枚ずつの画像ファイルではありません。


### train-images.idx3-ubyte、t10k-images.idx3-ubyteのファイル構造

|データ|サイズ|説明|
|---|---|---|
|ヘッダー|16byte||
|1枚目|784byte|横28×縦28=784ピクセル、1ピクセル=1byte<br/>1ピクセルは0～255のグレースケールを表す|
|...|...|...|
|n枚目|784byte|同上|

- 16byte+784byte/枚×60,000枚=47,040,016byte
- 16byte+784byte/枚×10,000枚=7,840,016byte


### train-labels.idx1-ubyte、t10k-labels.idx1-ubyteのファイル構造

|データ|サイズ|説明|
|---|---|---|
|ヘッダー|8byte||
|1枚目|1byte|0～9のいずれかの数値を表している|
|...|...|...|
|n枚目|1byte|同上|

- 8byte+1byte/枚×60,000枚=60,008byte
- 8byte+1byte/枚×10,000枚=10,008byte


## 訓練用画像を確認する

バイナリファイルからデータを読み込んで、疑似的な画像をダイアログボックス上に表示するプログラムです。

```vb
Dim input1
Dim input2

'MNISTデータベースファイルをC:\testに保存している場合
input1 = "C:\test\train-images.idx3-ubyte"
input2 = "C:\test\train-labels.idx1-ubyte"

Dim images
Dim labels

'バイナリ形式でファイルを開く
Set images = CreateObject("ADODB.Stream")
Set labels = CreateObject("ADODB.Stream")
images.Type = 1 'BINARY
labels.Type = 1 'BINARY
images.Open
labels.Open
images.LoadFromFile(input1)
labels.LoadFromFile(input2)

Dim myVal
Dim label
Dim image(783)
Dim i
Dim buf

Randomize '乱数ジェネレータを初期化

'キャンセルボタンが押されるまで繰り返す
Do
    myVal = Int((Rnd * 60000) + 1) '1～60000の乱数
    images.Position = 16 + 784 * (myVal - 1)
    labels.Position = 8 + (myval - 1)

    '1画像の各ピクセルデータを1次元の配列に格納する
    For i = 0 To 783
        image(i) = AscB(images.Read(1))
    Next

    '正解の数値
    label = AscB(labels.Read(1))

    '画像をダイアログボックス上に疑似的に表示する
    buf = ""
    For i = 0 To 783
        If image(i) > 128 Then
            buf = buf & "■"
        Else
            buf = buf & "□"
        End If

        If (i + 1) Mod 28 = 0 Then
            buf = buf & vbCr
        End If
    Next

    If Msgbox(buf & vbCr & "正解: " & label, vbOKCancel, "No." & myVal) = vbCancel Then
        Exit Do
    End If
    
Loop

images.Close
labels.Close
```
