---
title: "ファイル名、パス名を取得する"
permalink: /path/
excerpt: "FileSystemObjectオブジェクト"
last_modified_at: 2021-11-14T00:00:00-02:00
toc: false
---

## FileSystemObjectのファイル名、パス名を返すメソッド

引数にパスやファイル名を指定して使用する。

|メソッド|説明|
|---|---|
|GetAbsolutePathName|フルパスを返す|
|GetBaseName|拡張子を除いたファイル名を返す|
|GetExtensionName|拡張子を返す|
|GetFileName|ファイル名を返す|
|GetDriveName|ドライブ名を返す|
|GetParentFolderName|親フォルダ名を返す|

使用例：
スクリプトファイルをC:\Scriptに置いている場合。

```vb
Dim fso, path
Set fso = CreateObject("Scripting.FileSystemObject")

'フルパスで指定した場合
path = "C:¥Script¥test.txt"
Msgbox fso.GetAbsolutePathName(path) 'C:¥Script¥test.txt
Msgbox fso.GetBaseName(path)         'test
Msgbox fso.GetExtensionName(path)    'txt
Msgbox fso.GetFileName(path)         'test.txt
Msgbox fso.GetDriveName(path)        'C:
Msgbox fso.GetParentFolderName(path) 'C:¥Script

'ファイル名で指定した場合
path = "test.txt"
Msgbox fso.GetAbsolutePathName(path) 'C:¥Script¥test.txt
Msgbox fso.GetBaseName(path)         'test
Msgbox fso.GetExtensionName(path)    'txt
Msgbox fso.GetFileName(path)         'test.txt
Msgbox fso.GetDriveName(path)        '
Msgbox fso.GetParentFolderName(path) '

'長さ0の文字列で指定した場合
path = ""
Msgbox fso.GetAbsolutePathName(path) 'C:¥Script
Msgbox fso.GetBaseName(path)         '
Msgbox fso.GetExtensionName(path)    '
Msgbox fso.GetFileName(path)         '
Msgbox fso.GetDriveName(path)        '
Msgbox fso.GetParentFolderName(path) '

'現在の作業フォルダで指定した場合
path = "."
Msgbox fso.GetAbsolutePathName(path) 'C:¥Script
Msgbox fso.GetBaseName(path)         '
Msgbox fso.GetExtensionName(path)    '
Msgbox fso.GetFileName(path)         '.
Msgbox fso.GetDriveName(path)        '
Msgbox fso.GetParentFolderName(path) '

'現在の作業フォルダの親フォルダで指定した場合
path = ".."
Msgbox fso.GetAbsolutePathName(path) 'C:¥
Msgbox fso.GetBaseName(path)         '.
Msgbox fso.GetExtensionName(path)    '
Msgbox fso.GetFileName(path)         '..
Msgbox fso.GetDriveName(path)        '
Msgbox fso.GetParentFolderName(path) '
```

引数のファイルが実在しなくてもエラーは発生せず値は返ってくる。

## FileSystemObjectのBuildPathメソッドを使う

2つの引数を適宜`\`を補足してつないだパスを返す。

```vb
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Msgbox fso.BuildPath("C:\Script", "test.txt")   'C:\Script\test.txt
Msgbox fso.BuildPath("C:\Script\", "test.txt")  'C:\Script\test.txt
Msgbox fso.BuildPath("C:\Script", "\test.txt")  'C:\Script\test.txt
Msgbox fso.BuildPath("C:\Script\", "\test.txt") 'C:\Script\test.txt
Msgbox fso.BuildPath("C:\Script\test.txt", "new.txt") 'C:\Script\test.txt\new.txt
Msgbox fso.BuildPath("C:", "test.txt")          'C:test.txt
Msgbox fso.BuildPath("C:\", "test.txt")         'C:\test.txt
```

このメソッドは自分は一度も使ったことが無かった。

## FileSystemObjectのGetSpecialFolderメソッドを使う

|引数|説明|
|---|---|
|0|Windowsフォルダ|
|1|システムフォルダ|
|2|Tempフォルダ|

使用例：

```vb
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Msgbox fso.GetSpecialFolder(0) 'C:\Windows
Msgbox fso.GetSpecialFolder(1) 'C:\Windows\System32
Msgbox fso.GetSpecialFolder(2) 'C:\Users\ログインユーザ\AppData\Local\Temp
```

GetSpecialFolderメソッドで取得できるのはオブジェクトだが、
Msgboxに渡すとデフォルトプロパティのパスが返るのでエラーにはならない。

## WScriptオブジェクトのファイル名、パス名を返すプロパティ

|プロパティ|説明|
|---|---|
|Name|Windows Script Host|
|Path|Windows Script Hostプログラムがある場所|
|ScriptName|スクリプトファイルのファイル名|
|ScriptFullName|スクリプトファイルのフルパス|

使用例：

```vb
Msgbox WScript.Name
Msgbox WScript.Path
Msgbox WScript.ScriptName
Msgbox WScript.ScriptFullName
```

## 特殊フォルダのパスを取得する

WshShellオブジェクトのSpecialFoldersプロパティを使って、
SpecialFoldersコレクションを取得し、特殊フォルダのパスを取得できる。

```vb
Dim col
Set col = CreateObject("WScript.Shell").SpecialFolders

Msgbox col(0)                 'C:\Users\Public\Desktop
Msgbox col("AllUsersDesktop") 'C:\Users\Public\Desktop
```

|数値|指定名|説明|
|---|---|---|
|0|AllUsersDesktop|全ユーザ共通のデスクトップ|
|1|AllUsersStartMenu|全ユーザ共通のスタートメニュー|
|2|AllUsersPrograms|全ユーザ共通のスタートメニューのプログラムフォルダ|
|3|AllUsersStartup|全ユーザ共通のスタートアップフォルダ|
|4|Desktop|デスクトップ|
|5|AppData|アプリケーションデータ|
|6|PrintHood|プリンタ|
|7|Templates|テンプレート|
|8|Fonts|フォント|
|9|NetHood|ネットワーク|
|10|Desktop|デスクトップ|
|11|StartMenu|スタートメニュー|
|12|SendTo|送るフォルダ|
|13|Recent|最近使ったファイル|
|14|Startup|スタートアップフォルダ|
|15|Favorites|お気に入り|
|16|MyDocuments|マイドキュメント|
|17|Programs|プログラムフォルダ|

## 現在の作業フォルダの取得と変更

WshShellオブジェクトのCurrentDirectoryプロパティを使って、
現在の作業フォルダの取得、変更ができる。

```vb
Dim shell
Set shell = CreateObject("WScript.Shell")
'取得
Msgbox shell.CurrentDirectory      'C:\Script
'変更
shell.CurrentDirectory = "C:\Test"
Msgbox shell.CurrentDirectory      'C:\Test
```

VBAだと`ChDir`と`ChDrive`を使って現在の作業フォルダを変更できる。
以前VBScriptでの変更方法が分からないことがあった。この方法で変更できる。
