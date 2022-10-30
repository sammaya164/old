---
title: "ファイル名の頭に日付を加える"
permalink: /rename/
excerpt: ""
last_modified_at: 2021-11-13T00:00:00-02:00
toc: false
---

ファイル(フォルダ)をドラッグ＆ドロップすると名前を変更する。

`test.txt` -> `211113-1_test.txt`

```vb
'1ファイル(フォルダ)をドラッグ＆ドロップすると名前を変更する
Select Case WScript.Arguments.Count
Case 1
    Call Rename(WScript.Arguments.Item(0))
Case Else
    Msgbox "1ファイル(フォルダ)をドラッグ＆ドロップして下さい。",, "使い方"
End Select


'ファイル(フォルダ)名変更
Function Rename(sPath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
  
    Dim f
    'ファイル(フォルダ)オブジェクトを取得する
    If fso.FileExists(sPath) Then
        Set f = fso.GetFile(sPath)
    ElseIf fso.FolderExists(sPath) Then
        Set f = fso.GetFolder(sPath)
    End If
    
    Dim head
    head = GetDateString() '6桁の日付

    'ファイル(フォルダ)名を変更
    Call pr_Rename(f, head, 1) '3番目の引数は名前の重複を避けるための枝番

    '終了処理
    Set fso = Nothing
    Set f   = Nothing
End Function


'6桁の日付を返す
Function GetDateString()
    Dim buf
    '使用者が数字を入力する
    buf = Inputbox("何日前の日付にしますか", "名前変更", 0)

    '数字でない場合やキャンセルボタンが押された場合は終了
    If buf = "" Then WScript.Quit()
    If Not IsNumeric(buf) Then WScript.Quit()

    '数字として扱う
    buf = CLng(buf)

    '日付を返す
    GetDateString = Right(Replace(Date() - buf, "/", ""), 6)
End Function


'ファイル(フォルダ)名変更_再帰呼び出し用
Function pr_Rename(f, head, branch)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim parent
    parent = fso.GetParentFolderName(f)
    If Right(parent, 1) <> "\" Then parent = parent & "\"

    Dim name
    name = fso.GetFileName(f)
    
    Dim path
    path = parent & "\" & head & "-" & branch & "_" & name
    
    '変更後と同名のファイル(フォルダ)が無ければ名前を変更する
    If (TypeName(f) = "File"   And Not fso.FileExists  (path)) Or _
       (TypeName(f) = "Folder" And Not fso.FolderExists(path)) Then
        f.Move(path) '変更
    Else
        '同名のファイル(フォルダ)があれば枝番を1増やして再起呼び出し
        Call pr_Rename(f, head, branch + 1)
    End If
    
    '終了処理
    Set fso = Nothing
End Function
```
