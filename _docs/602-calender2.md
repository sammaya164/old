---
title: "カレンダー計算2"
permalink: /calender2/
last_modified_at: 2022-11-30T22:51:00+09:00
toc: true
---


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


## 万年カレンダー

カレンダーを表示するプログラムです。

1582年10月15日以降はグレゴリオ暦、それより前はユリウス暦で表示します。

```vb
'MJDや日付を指定すると、カレンダーを表示する
Call Main()



Function Main()
    Dim ret
    Dim mjd
    Dim dat
    Dim msg(7)

    msg(2) = "ユリウス暦:"
    msg(4) = "グレゴリオ暦:"
    ret = Date()

    Do  
        ret = InputBox("MJDまたは日付を入力してください。", "万年カレンダー", ret)
        If ret = "" Then Exit Do
        If IsNumeric(ret) Then
            mjd    = CDbl(ret)
            dat    = GetDate(mjd) '年月日の配列
            msg(0) = "MJD=" & mjd
            msg(3) = ToString(JulianDate(mjd))
            msg(5) = ToString(GregorianDate(mjd))
            msg(7) = GetCalender(dat(0), dat(1)) '引数は年と月
            Msgbox Join(msg, vbCr),,"万年カレンダー"  '配列を改行でつなげて表示
        
        ElseIf UBound(Split(ret, "/")) = 2 Then
            dat    = Split(ret, "/") '年月日の配列
            dat    = ToNumeric(dat)  '配列要素を数値に変換
            mjd    = GetMJD(dat(0), dat(1), dat(2))
            msg(0) = ret & " (MJD=" & mjd & ")"
            msg(3) = ToString(JulianDate(mjd))
            msg(5) = ToString(GregorianDate(mjd))
            msg(7) = GetCalender(dat(0), dat(1)) '引数は年と月
            Msgbox Join(msg, vbCr),,"万年カレンダー"  '配列を改行でつなげて表示
        Else
            ret = Date()
        End If
    Loop
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



'日付からユリウス暦かグレゴリオ暦かを判断してMJDを返す
Function GetMJD(y, m, d)
    If 10000*y + 100*m + d < 15821015 Then
        GetMJD = DateToMJD(y, m, d, 1)
    Else
        GetMJD = DateToMJD(y, m, d, 2)
    End If
End Function



'MJDからユリウス暦の日付(配列)に変換して返す
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
    JulianDate =  Array(y,m,Round(d, 5)) '日は小数点以下5桁まで
End Function



'MJDからグレゴリオ暦の日付(配列)に変換して返す
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
    GregorianDate =  Array(y,m,Round(d, 5)) '日は小数点以下5桁まで
End Function



'MJDからユリウス暦かグレゴリオ暦かを判断して日付を返す
Function GetDate(mjd)
    If mjd >= -100840 Then
        GetDate = GregorianDate(mjd) '1582/10/15以降ならグレゴリオ暦の日付を返す
    Else
        GetDate = JulianDate(mjd) 'それより前ならユリウス暦の日付を返す
    End If
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



'日付の配列を文字列に変換する
Function ToString(dat)
    ToString = Join(dat, "/")
End Function



'カレンダーを返す
Function GetCalender(y, m)
    Dim mjd1, mjd2
    Dim cale(5, 6) 'カレンダーを表す配列
    Dim i '週
    Dim j '曜日: 0(日)～6(土)
    Dim buf

    mjd1 = GetMJD(y, m, 1)      '当月1日の修正ユリウス日
    mjd2 = GetMJD(y, m + 1, 1) '翌月1日の修正ユリウス日

    j = (mjd1 + 3) - Int((mjd1 + 3) / 7) * 7 '当月1日の曜日
    i = 0 '1週目
    Do While mjd1 < mjd2
        cale(i, j) = GetDate(mjd1)(2) '日付
        mjd1 = mjd1 + 1
        j = j + 1
        If j > 6 Then
            j = 0
            i = i + 1
        End If
    Loop

    For i = 0 To 5
        For j = 0 To 6
            buf = buf & String(4-Len(cale(i,j)), "_") & cale(i,j)
        Next
        buf = buf & vbCr '週ごとに改行
    Next

    buf = "Sun_Mon_Tue_Wed_Thu_Fri_Sat_" & vbCrLf & buf
    buf = y & "年" & m & "月" & vbCrLf & buf

    '結果を返す
    GetCalender = buf

End Function
```
