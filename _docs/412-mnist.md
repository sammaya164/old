---
title: "バイナリファイルを読む"
permalink: /mnist/
last_modified_at: 2022-11-20T10:49:00+09:00
toc: true
---


## バイナリファイルの入手

ニューラルネットワークの訓練と評価に使用されている、MNISTデータベースを[ここ](http://yann.lecun.com/exdb/mnist/)からダウンロードします。

手書き数字画像のデータです。
{: .notice--info}

- train-images-idx3-ubyte.gz
- train-labels-idx1-ubyte.gz
- t10k-images-idx3-ubyte.gz
- t10k-labels-idx1-ubyte.gz

圧縮されているので7-Zipなどを使って展開します。

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


## 訓練用画像を表示する

バイナリファイルからデータを読み込んで、疑似的な画像をダイアログボックス上に表示するプログラムです。

```vb
'手書き数字画像のバイナリファイルを読んで、疑似的な画像を表示する

Dim path1, path2
Dim sm1  , sm2
Dim num

'MNISTデータベースファイルをC:\samplesに保存している場合
path1 = "C:\samples\train-images.idx3-ubyte" '画像データ
path2 = "C:\samples\train-labels.idx1-ubyte" '値データ

'バイナリ形式でファイルを開く
Set sm1 = P_OpenStream(path1)
Set sm2 = P_OpenStream(path2)

Call Randomize() '乱数ジェネレータを初期化

'キャンセルボタンが押されるまで繰り返す
Do
    num = Int(Rnd * 60000) '0～59999のランダムな整数

    '疑似的な画像を表示する、タイトルに数値を表示する
    If Msgbox(P_GetImage(num, sm1), vbOKCancel, "正解: " & P_GetValue(num, sm2)) = vbCancel Then
        Exit Do
    End If
    
Loop

sm1.Close
sm2.Close



'バイナリ形式でファイルを開く
Function P_OpenStream(path)
    dim sm

    Set sm = CreateObject("ADODB.Stream")
    sm.Type = 1 'BINARY
    sm.Open
    sm.LoadFromFile(path)
    Set P_OpenStream = sm 'Streamオブジェクトを返す

End Function



'疑似的な画像データを取得する
Function P_GetImage(n, sm)
    Dim i, j '画像の横と縦の座標
    Dim bt
    Dim buf

    sm.Position = 16 + 784 * n '現在位置をn+1番目の画像データの先頭に移動する
    
    '1画像の各ピクセルデータから□と■で作成した疑似的な画像を作成する
    For j = 0 To 27
        For i = 0 To 27
            bt = AscB(sm.Read(1)) '1バイト読んで数字に変換
            If bt <128 Then
                buf = buf & "■"
            Else
                buf = buf & "□"
            End If
        Nextc
        buf = buf & vbCr
    Next

    P_GetImage = buf

End Function



'値データを取得する
Function P_GetValue(n, sm)
    sm.Position = 8 + n '現在位置をn+1番目の値データに移動する
    P_GetValue = AscB(sm.Read(1)) '1バイト読んで数字に変換
    
End Function
```

Read(1)で1バイト読むと、現在位置は次のバイトへ移動します。
{: .notice--primary}

読み取ったバイトをAscB関数で0～255の数値へ変換します。
{: .notice--primary}


### 実行結果

![実行結果](/vbscript/assets/images/mnist.jpg)
