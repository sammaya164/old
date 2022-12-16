---
title: "クラスを自作する"
permalink: /class/
last_modified_at: 2022-12-16T23:30:00+09:00
toc: true
---

## クラスの作成例(勇者vs.魔王)

```vb
'クラス宣言
Class C_Player
    'ここでは変数宣言にDimは使えない

    'パブリックなメンバを宣言(プロパティ)
    '(クラスの外部からアクセス可能)
    Public Name '名前
    
    'プライベートなメンバを宣言
    '(クラスの外部からアクセス不可能)
    Private m_life     'ライフ
    Private m_max      'ライフ最大値
    Private m_strength '攻撃力
    Private m_enemy    '敵オブジェクト

    'Strengthプロパティへ値の代入を可能にする
    '(クラスの外部から代入可能)
    Public Property Let Strength(val)
        m_strength = val
        m_life = Int(100000 / val)
        m_max = m_life
    End Property
    
    'Strengthプロパティの値を取得可能にする
    '(クラスの外部から取得可能)
    Public Property Get Strength()
        Strength = m_strength
    End Property

    'Enemyプロパティへオブジェクトの代入を可能にする
    '(クラスの外部から代入可能)
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
    '(クラスの外部からアクセス可能)
    Public Sub Atack()
        'プロシージャ内ならDimを使える
        Dim buf
        buf = Int(m_strength * Rnd())
        Msgbox Name & "の攻撃！"
        m_enemy.Damage(buf)
    End Sub
    
    'パブリックなFunctionプロシージャ(メソッド)
    '(クラスの外部からアクセス可能)
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
    '(クラスの外部からアクセス不可能)
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
Call Randomize()

'戦闘開始
Call P1.Atack()

'インスタンスを破棄
Set P1 = Nothing
Set P2 = Nothing
```

上記コードをテキストエディタに貼り付け、maou.vbsなどの名前で保存してください。

Shift_JIS形式で保存してください。
{: .notice--primary}


## 説明

ダブルクリックすると起動します。  
魔王と勇者、どちらかが倒れるまで戦闘が続きます。

<button type="button" onclick="maou();">JavaScriptによるデモ</button>{: .btn .btn--success .btn--large}

![魔王](/vbscript/assets/images/maou3.jpg)

<script>
    // <!--

    class Player{
      constructor(name, strength){
        this.name = name;
        this.strength = strength;
        this.lifeMax = 100000/strength;
        this.life = this.lifeMax;
      }

      damage(val){
        this.life = this.life - val;
        if (this.life < 0) {
          this.life = 0;
        }
        alert(this.name + 'は' + val + 'のダメージを受けた\n(´・ω・) ' + this.life + '/' + this.lifeMax);
        if (this.life == 0) {
          alert(this.name + 'は倒れた！\n( -ω-) ｽﾔｧ');
        } else {
          this.atack();
        }
      }

      atack(){
        alert(this.name + 'の攻撃！')
        this.enemy.damage(Math.floor(Math.random() * this.strength));
      }

      get enemy(){
        return this._enemy;
      }

      set enemy(obj){
        this._enemy = obj;
      }
    }

    function maou() {
      
      let p1 = new Player ('勇者', 100);
      let p2 = new Player ('魔王', 1000);
      p1.enemy = p2;
      p2.enemy = p1;
      p1.atack();
    }

    // -->
</script>


## VBAで同じことをしたい場合

下記手順で動きます。
    
1. `Class C_Player` ～ `End Class` の中身だけをクラスモジュールに記載する。
2. クラスモジュールの名前を `C_Player` に変更する。
3. `'変数宣言` 以降を標準モジュールの `Sub Maou` ～ `End Sub` に記載する。
4. 開発タブの `マクロ` をクリックし `Maou` を実行する。

