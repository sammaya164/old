---
title: "File、Folder、Driveオブジェクト"
permalink: /file/
last_modified_at: 2022-11-18T16:00:00+09:00
toc: true
---

## Fileオブジェクト

|プロパティ|説明|
|---|---|
|Attributes|ファイルの属性を返す|
|DataCreated|作成日|
|DataLastAccesed|最後にアクセスされた日時|
|DataLastModified|最後に更新された日時|
|Drive|ドライブの名前|
|Name|名前|
|ParentFolder|親フォルダー|
|Path|パス|
|Size|サイズ|
|Type|ファイルの種類|

|メソッド|説明|
|---|---|
|Copy()|コピーする|
|Move()|移動する|
|Delete()|削除する|
|OpenAsTextStream()|テキストファイルとして開く|
 
 
## Folderオブジェクト

|プロパティ|説明|
|---|---|
|Attributes|フォルダーの属性を返す|
|DataCreated|作成日|
|DataLastAccesed|最後にアクセスされた日時|
|DataLastModified|最後に更新された日時|
|Drive|ドライブの名前|
|Name|名前|
|ParentFolder|親フォルダー|
|Path|パス|
|Size|サイズ|
|Type|フォルダーの種類|
|Files|ファイルのコレクション|
|SubFolders|サブフォルダーのコレクション|
|IsRootFolder|ルートフォルダーならTrue|

|メソッド|説明|
|---|---|
|Copy|コピーする|
|Move|移動する|
|Delete|削除する|
 

## Driveオブジェクト

|プロパティ、メソッド|説明|
|---|---|
|DriveLetter|ドライブ名|
|DriveType|ドライブの種類|
|IsReady|準備できていればTrue|
|Path|ドライブのパス|
|RootFolder|ドライブのルートフォルダ|
|SerialNumber|ドライブのシリアル番号|
|VolumeName|ボリューム名|
 

## Files、Folders、Drivesコレクション

|プロパティ|説明|
|---|---|
|Count|コレクションに含まれるメンバーオブジェクトの数|
|Item|引数で指定したメンバーオブジェクトを取得|


### Foldersコレクションのメソッド

|メソッド|説明|
|---|---|
|Add|引数で指定した名前のサブフォルダを追加する|
