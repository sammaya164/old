---
title: "カレンダー計算1"
permalink: /calender1/
last_modified_at: 2022-11-27T08:00:00+09:00
toc: true
---

## 日付を日数で表す

年月日で表される日付は、いわば年・月・日の3つの軸をもつ空間に並べられた
点の集まりと解釈できます。
しかし天文計算においては、日付は1本の線上に並べたほうが扱い易くなります。
すなわち、ある特定の日を第0日としたときの通算の日数で日付を表すという方法です。

### ユリウス日(Julian Day)

- ユリウス日(ユリウス通日、JD)は、紀元前4713年1月1日正午(ユリウス暦)を第0日としたときの通算の日数で日付を表したものです。
- 正負の整数であれば正午を表しますが、小数部で時刻を表します。

2023年1月1日の正午で、2,459,946になります。
{: .notice--info}

### 修正ユリウス日(Modified Julian Day)

- 修正ユリウス日(準ユリウス日、MJD)は、ユリウス日よりも小さい桁で日付を扱えるよう考え出され、ユリウス日から2400000.5日を引いたものとして定義されます。
- 1858年11月17日(グレゴリオ暦)を第0日としたときの通算の日数で日付を表します。
- 正負の整数であれば午前0時を表しますが、小数部で時刻を表します。

2023年1月1日の午前0時で59,945になります。
{: .notice--info}


## 西暦

西暦には、うるう年をどのように定めるかによって、ユリウス暦とグレゴリオ暦があります。
現在使われているのはグレゴリオ暦です。

### ユリウス暦

以下のルールがあります。

1. うるう年は4年に1回。

すなわち4年は正確に1,461日(=365×3+366)、1年あたりは365.25日(=1,461/4)。

実際の1太陽年はおよそ365.2422日なので「春分、夏至、秋分、冬至」といった季節の移り変わりに対し、
1年あたり-0.0078日(=365.2422-365.25)の誤差が生じます。1600年経てば約12.5日のズレになります。

ユリウス暦は紀元前45年にローマで導入されましたが、当初は4年に1回とすべきうるう年を、3年に1回で誤運用してしまいました。 
誤りに気づいた後、調整の為うるう年を設けない期間が設けられ、西暦8年のうるう年以降は4年に1回で運用されました。
以下のプログラムにおけるユリウス暦の計算では、このへんの歴史は考慮せず、単純に4年に1回うるう年を設けています。
{: .notice--info}

### グレゴリオ暦

以下のルールがあります。

1. うるう年は4年に1回
1. ただし100年に1回(西暦が100で割り切れる年)はうるう年を取消し。
1. ただし400年に1回(西暦が400で割り切れる年)はやっぱりうるう年。

すなわち400年は正確に146,097日(=365×303+366×97)、1年あたりは365.2425日(=146,097/400)。

期間の選び方によって、4年は1,460日の場合と1,461日の場合があり、100年は36,524日の場合と36,525日の場合があります。
{: .notice--info}

グレゴリオ暦は、ユリウス暦1582年10月4日の翌日、グレゴリオ暦1582年10月15日からローマで採用されました。
グレゴリオ暦の導入時期は国によって異なり、日本では旧暦明治5年12月2日の翌日、新暦明治6年1月1日から開始されました。

季節の移り変わりに対し、1年あたり約+0.0003日の誤差が生じます。
1万年経っても約3日のズレしかない計算になります。

なお地球の自転や公転のスピードは徐々に変動している為、それらの比である1太陽年の日数も変動します。
1万年後のズレは3日より小さいかも知れないし、大きいかも知れません。
{: .notice--info}


### 紀元前の表現

西暦1年の前年は紀元前1年ですが、天文学では年数計算を簡便にする為、紀元前1年を西暦0年、
紀元前2年を西暦-1年、・・、紀元前N年を西暦-(N-1)年とする紀年法が使われるそうです。 
以下のプログラムもこれに従います。たとえば紀元前4713年は西暦-4712年になります。


## ユリウス日、修正ユリウス日を計算する

年月日からユリウス日、修正ユリウス日を算出して表示するプログラムです。重要なのは`DateToMJD`関数の箇所です。

```vb
'ユリウス日、修正ユリウス日を計算する
Dim ret
Dim dt '年月日を表す配列
Dim buf(6)
Dim mjd

Do
    ret = InputBox("日付を入力して下さい", , Date())
    If ret = "" Then Exit Do
    dt = Split(ret, "/") 'Splitで取得した配列要素は文字列
    dt = ToNumeric(dt)   '配列要素を数値へ変換

    mjd = DateToMJD(dt(0), dt(1), dt(2), 1)
    buf(0) = "ユリウス暦の" & ret & "は"
    buf(1) = "JD=" & mjd + 2400000.5
    buf(2) = "MJD=" & mjd

    mjd = DateToMJD(dt(0), dt(1), dt(2), 2)
    buf(4) = "グレゴリオ暦の" & ret & "は"
    buf(5) = "JD=" & mjd + 2400000.5
    buf(6) = "MJD=" & mjd

    Msgbox Join(buf, vbCr),, "ユリウス日計算" '配列をCR改行でつなげて表示
Loop



'配列の要素を数値へ変換する
Function ToNumeric(ByVal arr)
    Dim i

    For i = 0 To UBound(arr)
        If IsNumeric(arr(i)) Then '数値として解釈できるなら
            arr(i) = CDbl(arr(i)) 'Double型へ変換
        End If
    Next
    ToNumeric = arr
End Function



'修正ユリウス日を算出
'intCalender:暦を指定 (1:ユリウス暦, 2:グレゴリオ暦)
Function DateToMJD(ByVal Y, ByVal M, ByVal D, intCalender)
    '1,2月を前年の13,14月にする
    If M < 3 Then
        Y = Y - 1
        M = M + 12
    End If
    '暦に応じて日付から修正ユリウス日を算出
    Select Case intCalender
    Case 1
        DateToMJD = - 678884 + 365 * Y + MonthDay(M - 1) + D + Int(Y / 4)
    Case 2
        DateToMJD = - 678882 + 365 * Y + MonthDay(M - 1) + D + Int(Y / 4) - Int(Y / 100) + Int(Y / 400)
    End Select
End Function



' 2月末日を第0日としてM月末日までの日数を返す(M=2～13)
Function MonthDay(M)
    MonthDay = Int(30.59 * (M - 1)) - 30
End Function
```

**ByVal:** VBScriptでは引数はデフォルトで参照渡しですが、ByValをつけることで値渡しになります。
{: .notice--primary}

1,2月を13,14月に変換するのは、計算上2月末日を1年の最終日、もしくは翌年の第0日として扱うためです。
{: .notice--info}

MJDの計算式に出てくる-678884と-678882は、それぞれユリウス暦とグレゴリオ暦における紀元前1年2月29日のMJDです。
{: .notice--info}

上記のMonthDay関数の引数と戻り値は次表のようになります。

|引数|2|3|4|5|6|7|8|9|10|11|12|13|
|---|---|---|---|---|---|---|
|戻り値|0|31|61|92|122|153|184|214|245|275|306|337|

次のように書いたほうが、何をやっているかはわかりやすいかも知れません。

```vb
Function MonthDay(M)
    Dim days(13)
    Dim i
    
    days(4)  = 30
    days(5)  = 31
    days(6)  = 30
    days(7)  = 31
    days(8)  = 31
    days(9)  = 30
    days(10) = 31
    days(11) = 30
    days(12) = 31
    days(13) = 31
    
    For i = 4 To M
        MonthDay = MonthDay + days(i)
    Next
End Function
```

## 暦の計算

ユリウス日や修正ユリウス日は、天文計算のほか、暦の計算にも使えます。

- ユリウス暦、グレゴリオ暦の間で日付を変換したい場合:  
  一旦修正ユリウス日に変換し、他方の日付へ変換します
- 曜日を知りたい場合:  
  (修正ユリウス日+3)を7で割った余りから曜日を算出します(0:日、1:月、2:火、3:水、4:木、5:金、6:土)
- 日干支を知りたい場合:  
  修正ユリウス日を10で割った余りから十干を算出(0:甲、1:乙、2:丙、3:丁、4:戊、5:己、6:庚、7:辛、8:壬、9:癸)  
  (修正ユリウス日+2)を12で割った余りから十二支を算出(0:子、1丑、2:寅、3:卯、4:辰、5:巳、6:午、7:未、8:申、9:酉、10:戌、11:亥)  
  2023年1月1日は5と7で己未(つちのとひつじ)


## 修正ユリウス日から日付へ変換する

日付への変換は、その逆より難しいです。

```vb
'MJDから日付へ変換して表示する

Dim ret
Dim msg(4)
msg(1) = "ユリウス暦:"
msg(3) = "グレゴリオ暦:"

Do
    ret = InputBox("MJDを入力してください。")
    If ret = "" Then Exit Do
    ret = CDbl(ret)
    msg(0) = "MJD=" & ret
    msg(2) = JulianDate(ret)
    msg(4) = GregorianDate(ret)
    Msgbox Join(msg, vbCr)  '配列をCR改行でつなげて表示
Loop



'MJDからグレゴリオ暦の日付に変換して返す
Function GregorianDate(mjd)
    Dim y, m, d

    'MJDをグレゴリオ暦の西暦0年(紀元前1年)2月29日を第0日とした日数へ変換
    d = mjd - (-678882)
    y = 0
    m = 3

    Dim arr1, arr2, arr3

    '年の計算
    arr1 = Array(1,365,1461,36524,146097)
    arr2 = Array(0,1,4,100,400)
    arr3 = Div(d, arr1)
    y = y + Mult(arr2, arr3)
    d = arr3(0)

    '月の計算
    '2月末日の場合
    If d < 1 Then
        m = 2
        If arr3(1) = 0 And (arr3(2) <> 0 Or arr3(3) = 0)  Then
            d = 29 + d
        Else
            d = 28 + d
        End If
    Else '2月末日以外の場合
        Do While MonthDay(m) < d
            m = m + 1
        Loop
        d = d - MonthDay(m - 1)
        If m > 12 Then
            y = y + 1
            m = m - 12
        End If
    End If
    GregorianDate =  Join(Array(y,m,Round(d, 3)), "/") '日は小数点以下3桁まで
End Function



'MJDからユリウス暦の日付に変換して返す
Function JulianDate(mjd)
    Dim y, m, d

    'MJDをユリウス暦の西暦0年(紀元前1年)2月29日を第0日とした日数へ変換
    d = mjd - (-678884)
    y = 0
    m = 3

    Dim arr1, arr2, arr3

    '年の計算
    arr1 = Array(1,365,1461)
    arr2 = Array(0,1,4)
    arr3 = Div(d, arr1)
    y = y + Mult(arr2, arr3)
    d = arr3(0)

    '月の計算
    '2月末日の場合
    If d < 1 Then
        m = 2
        If arr3(1) = 0 Then
            d = 29 + d
        Else
            d = 28 + d
        End If
    Else '2月末日以外の場合
        Do While MonthDay(m) < d
            m = m + 1
        Loop
        d = d - MonthDay(m - 1)
        If m > 12 Then
            y = y + 1
            m = m - 12
        End If
    End If
    JulianDate =  Join(Array(y,m,Round(d, 3)), "/") '日は小数点以下3桁まで
End Function



'例えばarr=[1,365,1461]の場合、a(0)+365*a(1)+1461*a(2)=valとなる配列aを返す
Function Div(val, arr)
    Dim iMax
    Dim buf
    Dim i

    iMax = UBound(arr)
    ReDim buf(iMax)
    buf(0) = val
    For i = iMax To 1 Step -1
        buf(i) = Int(buf(0) / arr(i))     '商の整数部
        buf(0) = buf(0) - arr(i) * buf(i) '余り(>0)
    Next
    Div = buf
End Function



'2つの配列の「内積」を返す
Function Mult(arr1, arr2)
    Dim i

    For i = 0 To UBound(arr1)
        Mult = Mult + arr1(i) * arr2(i)
    Next
End Function



'2月末日を第0日としてM月末日までの日数を返す(M=2～13)
Function MonthDay(M)
  MonthDay = Int(30.59 * (M - 1)) - 30
End Function

```


## 参考

- [ユリウス通日(Wikipedia)](https://ja.wikipedia.org/wiki/%E3%83%A6%E3%83%AA%E3%82%A6%E3%82%B9%E9%80%9A%E6%97%A5)
- マイコン宇宙講座―楽しい軌道計算プログラム(中野圭一著)  
  自分が購入したときは(2009年)送料込みで1,090円だったのですが、今Amazonを見ると滅茶苦茶高騰してました...。
