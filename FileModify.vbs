Option Explicit
On Error Resume Next

' ################################################################################
' #
' #  Programe Name： FileModify.vbs
' #  OverView     ： Performs character string processing inside the file
' #                  Email me every time the process is finished
' #  Create Date  ： 2021/11/12
' #  Create Person： M.Hayashi
' #  Version      ： 1.0
' #  Remark       ： Arg1 = Read file (relative path)
' #               ： Arg2 = [Not used]
' #
' ################################################################################


' ------------------------------------------------------
'                    ０．準      備
' ------------------------------------------------------
' ファイルオブジェクト宣言
Dim FSO
Set FSO = WScript.CreateObject("Scripting.FileSystemObject")

' 正規表現オブジェクト宣言
Dim REG
Set REG = WScript.CreateObject("VBScript.RegExp")

' ログファイル名称取得
Dim LOG_FILE_NAME
LOG_FILE_NAME = Replace(WScript.ScriptFullName,".vbs",".log")

' メール用の定数
Dim strSMTP, strFromAddress, strToAddress, strNormalSubject, strErrSubject
strSMTP = ""		' IP Address of Mail Server
strFromAddress = "*****@***.***"
strToAddress   = "*****@***.***"	' 送り先。複数ある時は";"区切りで
strNormalSubject   = "[Project Name] 加工作業 正常終了"		' 正常終了時の件名
strErrSubject      = "[Project Name] 加工作業 エラー"		' エラー時の件名



' ------------------------------------------------------
'                    １．初期処理
' ------------------------------------------------------


' 前回のログファイルを削除
If FSO.FileExists(LOG_FILE_NAME) Then
	FSO.DeleteFile LOG_FILE_NAME
	If Err.Number <> 0 Then
		Call subErrHandling("ログファイルの削除が失敗しました。")
		WScript.Quit 9		' エラー値:9とする
	End If
End If

Call subWriteLog("プログラム開始。")



' ------------------------------------------------------
'                    ２．引数を受ける
' ------------------------------------------------------
Dim strInputFileName
strInputFileName = WScript.Arguments.Item(0)

If strInputFileName = "" Then
	Call subErrHandling("引数1の値がブランクの為終了しました。")
	WScript.Quit 9		' エラー値:9とする

ElseIf FSO.FileExists( strInputFileName ) = False Then
	Call subSendMail(strNormalSubject, "正常終了しました。")

	Call subWriteLog("変換ファイル[" & strInputFileName & "]が存在無し。処理なし。")
	Call subWriteLog("プログラム正常終了。")
	WScript.Quit 0

End If


' ------------------------------------------------------
'           ３．ファイルを仕様に基づき変容させる
' ------------------------------------------------------

Call subWriteLog("ファイル読み込み、書き出し開始。")

Dim fsoInput
Dim fsoOutput

' 読み込みファイルOBJECT生成
Set fsoInput = FSO.OpenTextFile(strInputFileName, 1, False, 0)


' 書き出しファイル名を生成
Dim strOutputFileName
REG.Pattern = "\.txt$"
strOutputFileName = REG.Replace(strInputFileName, "_2.txt")

' 前回の書き出しファイル名を削除
If FSO.FileExists(strOutputFileName) Then
	FSO.DeleteFile strOutputFileName
	If Err.Number <> 0 Then
		Call subErrHandling("前回の出力ファイル[" & strOutputFileName & "]の削除が失敗しました。")
		WScript.Quit 9		' エラー値:9とする
	End If
End If



' 書き出しファイルOBJECT生成
Set fsoOutput = FSO.OpenTextFile(strOutputFileName, 2, True)

' 読み込みファイルから1行ずつ読み込み、書き出しファイルに書き出すのを最終行まで繰り返す
Dim rCnt, wCnt
Do Until fsoInput.AtEndOfStream
	Dim strLine
	strLine = fsoInput.ReadLine
	rCnt = rCnt+1


	' ■TODO : オリジナルの加工方法をfncChkLineStr()に記載すること
	If fncChkLineStr(strLine) = True Then
		fsoOutput.WriteLine strLine
		wCnt = wCnt+1
	End If



Loop

' バッファを Flush してファイルを閉じる
fsoInput.Close
fsoOutput.Close

Call subWriteLog("[" & strInputFileName & "]の全" & rCnt & "行から" & wCnt & "行書き出ししました。" )
Call subWriteLog("ファイル読み込み、書き出し終了。")


' ------------------------------------------------------
'              ４．古いファイル削除、リネーム
' ------------------------------------------------------

' 読み込んだファイルを削除
FSO.DeleteFile strInputFileName
If Err.Number <> 0 Then
	Call subErrHandling("読み込みファイル[" & strInputFileName & "]の削除が失敗しました。")
	WScript.Quit 9		' エラー値:9とする
End If
Call subWriteLog("読み込みファイル[" & strInputFileName & "]を削除。" )


' 書き出したファイルをリネーム
REG.Pattern = "_2\.txt$"
Dim fsoFile, strFileName
Set fsoFile = FSO.GetFile( strOutputFileName )
strFileName = FSO.GetFileName( strOutputFileName )
fsoFile.Name = REG.Replace(strFileName, ".txt")
If Err.Number <> 0 Then
	Call subErrHandling("書き出したファイル[" & strOutputFileName & "]のリネームに失敗しました。")
	WScript.Quit 9		' エラー値:9とする
End If

Call subWriteLog("書き出しファイル[" & strOutputFileName & "]を[" & REG.Replace(strFileName, ".txt") & "]にリネーム。" )


' ------------------------------------------------------
'                  ５．メールを送信する
' ------------------------------------------------------
Call subSendMail(strNormalSubject, "正常終了しました。")
Call subWriteLog("メールを送信しました。")


' ------------------------------------------------------
'                    ６．終      了
' ------------------------------------------------------
Set FSO       = Nothing 
Set REG       = Nothing
Set fsoInput  = Nothing
Set fsoOutput = Nothing
Set fsoFile   = Nothing

Call subWriteLog("プログラム正常終了。")

WScript.Quit 0



' ======================================================================================
' ======================================================================================




' ------------------------------------------------------
'                    ログ書き出し処理
' ------------------------------------------------------
Sub subWriteLog(strMessage)
	' 概要：ログファイルに書き込みをする

	Dim fsoLog
	Set fsoLog = FSO.OpenTextFile(LOG_FILE_NAME, 8, True)
	fsoLog.WriteLine strMessage

	fsoLog.Close
	Set fsoLog = Nothing

End Sub

' ------------------------------------------------------
'                    テキスト内容変更
' ------------------------------------------------------
Function fncChkLineStr(strLine)
	' 概要：テキストの内容を変容させる

	' ■TODO:オリジナルの加工方法をここに記載すること

	fncChkLineStr = True
End Function

' ------------------------------------------------------
'                      メール送信
' ------------------------------------------------------
Sub subSendMail(strSubject, strTextBody)
	' 概要：メール送信
	Dim objMsg
    Set objMsg = WScript.CreateObject("CDO.Message")  

    objMsg.From     = strFromAddress
    objMsg.To       = strToAddress
    objMsg.Subject  = strSubject
    objMsg.TextBody = strTextBody

    objMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2  
    objMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTP
    objMsg.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25  
    objMsg.Configuration.Fields.Update  
    objMsg.Send  

    Set objMsg = Nothing

End Sub



' ------------------------------------------------------
'                   エラーハンドリング
' ------------------------------------------------------
Sub subErrHandling(strMsg)
	' 概要：ログファイルをリネーム、エラーの内容のメール送る

	' ログを記載後、リネームする
	Call subWriteLog(strMsg)
	Call subRenameLogFile()

	' メールを送る
	Call subSendMail(strErrSubject, strMsg)

End Sub



' ------------------------------------------------------
'                  ログファイルをリネーム
' ------------------------------------------------------
Sub subRenameLogFile()
	' 概要：ログファイルを日時付きの名称にRename

	If FSO.FileExists( LOG_FILE_NAME ) Then
		REG.Pattern = "\.vbs$"
		Dim fsoLogFile
		Set fsoLogFile = FSO.GetFile( LOG_FILE_NAME )
		fsoLogFile.Name = REG.Replace(WScript.ScriptName, "_" & fncGetStringDateTime() & ".log")
		Set fsoLogFile = Nothing
	End If

End Sub

' ------------------------------------------------------
'              今日の日時の文字列を返す
' ------------------------------------------------------
Function fncGetStringDateTime()
	' 概要：実行日時を返す（エラーログ用）
	Dim strDate
	Dim strTime
	
	strDate = Replace(FormatDateTime(Now, 1),"/","")
	strTime = Replace(FormatDateTime(Now, 4),":","")

	' 「yyyyMMdd_hhmm」の形式で返す
	fncGetStringDateTime = strDate & "_" & strTime

End Function
