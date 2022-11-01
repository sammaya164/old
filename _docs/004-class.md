---
title: "クラスを自作する"
permalink: /class/
last_modified_at: 2022-11-01T22:10:00+09:00
toc: false
---

クラスの使用例。

```vb
'クラス宣言
Class C_Player
    'ここでは変数宣言にDimは使えない

    'パブリックなメンバを宣言
    Public Name '名前
    
    'プライベートなメンバを宣言
    Private m_life     'ライフ
    Private m_max      'ライフ最大値
    Private m_strength '攻撃力
    Private m_enemy    '敵オブジェクト

    'プロパティへ値の代入を可能にする
    Public Property Let Strength(val)
        m_strength = val
        m_life = Int(100000 / val)
        m_max = m_life
    End Property
    
    'プロパティの値を取得可能にする
    Public Property Get Strength()
        Strength = m_strength
    End Property

    'プロパティへオブジェクトの代入を可能にする
    Public Property Set Enemy(obj)
        Set m_enemy = obj
    End Property
    
    Private Sub Class_Initialize()
        'インスタンス作成時に行いたい処理を記載   
    End Sub
    
    Private Sub Class_Terminate()
        'インスタンス破棄時に行いたい処理を記載
    End Sub
    
    'パブリックなSubプロシージャ(メソッド)
    Public Sub Atack()
        'プロシージャ内ならDimを使える
        Dim buf
        buf = Int(m_strength * Rnd())
        Msgbox Name & "の攻撃！"
        m_enemy.Damage(buf)
    End Sub
    
    'パブリックなFunctionプロシージャ(メソッド)
    Public Function Damage(val)
        m_life = m_life - val
        If m_life < 0 Then m_life = 0
        Msgbox Name & "は" & val & "のダメージを受けた" _
        & vbCr & "(´・ω・) " & m_life & "/" & m_max
        If HasDead Then
            Msgbox Name & "は倒れた！" & vbCr & "( -ω-) ｽﾔｧ"
        Else
            Call Atack()
        End If
    End Function
    
    'プライベートなFunctionプロシージャ
    Private Function HasDead()
        If m_life > 0 Then
            HasDead = False
        Else
            HasDead = True
        End If
    End Function
    
End Class


'変数宣言
Dim P1
Dim P2

'インスタンスを作成
Set P1 = New C_Player
Set P2 = New C_Player

'プロパティに値を代入
P1.Name = "勇者"
P1.Strength = 100
P2.Name = "魔王"
P2.Strength = 1000

'プロパティにオブジェクトを代入
Set P1.Enemy = P2
Set P2.Enemy = P1

'乱数ジェネレータを初期化
Randomize()

'戦闘開始
P1.Atack()

'インスタンスを破棄
Set P1 = Nothing
Set P2 = Nothing
```

上記コードをメモ帳などに貼り付け、maou.vbsなどの名前で保存。ダブルクリックすると起動します。
文字コードはShift_JIS(ANSI)でないとエラーになります。

魔王と勇者、どちらかが倒れるまで戦闘が続きます。

![まおう](/vbscript/assets/images/class_maou.png)
