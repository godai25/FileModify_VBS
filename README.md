# FileModify_VBS
VBS program to process the contents of a file

## 概要 Overview
引数に渡したファイルに対して、ファイル内部の文字列加工を行うVisual Basic Script


## 機能 Function
- ログファイルは[プログラム名].logで出力します。
- ファイルの有無等のエラーハンドリングを行う
- エラー時は、ログファイルを日時（[プログラム名]_yyyyMMdd_hhmm.log）付きに変更して上書きを避けます。
- 終了時にメールを送信します。また、エラー時でもメールを送信します。
