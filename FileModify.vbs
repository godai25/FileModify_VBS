Option Explicit
On Error Resume Next

' ################################################################################
' #
' #  Programe Name�F FileModify.vbs
' #  OverView     �F Performs character string processing inside the file
' #                  Email me every time the process is finished
' #  Create Date  �F 2021/11/12
' #  Create Person�F M.Hayashi
' #  Version      �F 1.0
' #  Remark       �F Arg1 = Read file (relative path)
' #               �F Arg2 = [Not used]
' #
' ################################################################################


' ------------------------------------------------------
'                    �O�D��      ��
' ------------------------------------------------------
' �t�@�C���I�u�W�F�N�g�錾
Dim FSO
Set FSO = WScript.CreateObject("Scripting.FileSystemObject")

' ���K�\���I�u�W�F�N�g�錾
Dim REG
Set REG = WScript.CreateObject("VBScript.RegExp")

' ���O�t�@�C�����̎擾
Dim LOG_FILE_NAME
LOG_FILE_NAME = Replace(WScript.ScriptFullName,".vbs",".log")

' ���[���p�̒萔
Dim strSMTP, strFromAddress, strToAddress, strNormalSubject, strErrSubject
strSMTP = ""		' IP Address of Mail Server
strFromAddress = "*****@***.***"
strToAddress   = "*****@***.***"	' �����B�������鎞��";"��؂��
strNormalSubject   = "[Project Name] ���H��� ����I��"		' ����I�����̌���
strErrSubject      = "[Project Name] ���H��� �G���["		' �G���[���̌���



' ------------------------------------------------------
'                    �P�D��������
' ------------------------------------------------------


' �O��̃��O�t�@�C�����폜
If FSO.FileExists(LOG_FILE_NAME) Then
	FSO.DeleteFile LOG_FILE_NAME
	If Err.Number <> 0 Then
		Call subErrHandling("���O�t�@�C���̍폜�����s���܂����B")
		WScript.Quit 9		' �G���[�l:9�Ƃ���
	End If
End If

Call subWriteLog("�v���O�����J�n�B")



' ------------------------------------------------------
'                    �Q�D�������󂯂�
' ------------------------------------------------------
Dim strInputFileName
strInputFileName = WScript.Arguments.Item(0)

If strInputFileName = "" Then
	Call subErrHandling("����1�̒l���u�����N�̈׏I�����܂����B")
	WScript.Quit 9		' �G���[�l:9�Ƃ���

ElseIf FSO.FileExists( strInputFileName ) = False Then
	Call subSendMail(strNormalSubject, "����I�����܂����B")

	Call subWriteLog("�ϊ��t�@�C��[" & strInputFileName & "]�����ݖ����B�����Ȃ��B")
	Call subWriteLog("�v���O��������I���B")
	WScript.Quit 0

End If


' ------------------------------------------------------
'           �R�D�t�@�C�����d�l�Ɋ�Â��ϗe������
' ------------------------------------------------------

Call subWriteLog("�t�@�C���ǂݍ��݁A�����o���J�n�B")

Dim fsoInput
Dim fsoOutput

' �ǂݍ��݃t�@�C��OBJECT����
Set fsoInput = FSO.OpenTextFile(strInputFileName, 1, False, 0)


' �����o���t�@�C�����𐶐�
Dim strOutputFileName
REG.Pattern = "\.txt$"
strOutputFileName = REG.Replace(strInputFileName, "_2.txt")

' �O��̏����o���t�@�C�������폜
If FSO.FileExists(strOutputFileName) Then
	FSO.DeleteFile strOutputFileName
	If Err.Number <> 0 Then
		Call subErrHandling("�O��̏o�̓t�@�C��[" & strOutputFileName & "]�̍폜�����s���܂����B")
		WScript.Quit 9		' �G���[�l:9�Ƃ���
	End If
End If



' �����o���t�@�C��OBJECT����
Set fsoOutput = FSO.OpenTextFile(strOutputFileName, 2, True)

' �ǂݍ��݃t�@�C������1�s���ǂݍ��݁A�����o���t�@�C���ɏ����o���̂��ŏI�s�܂ŌJ��Ԃ�
Dim rCnt, wCnt
Do Until fsoInput.AtEndOfStream
	Dim strLine
	strLine = fsoInput.ReadLine
	rCnt = rCnt+1


	' ��TODO : �I���W�i���̉��H���@��fncChkLineStr()�ɋL�ڂ��邱��
	If fncChkLineStr(strLine) = True Then
		fsoOutput.WriteLine strLine
		wCnt = wCnt+1
	End If



Loop

' �o�b�t�@�� Flush ���ăt�@�C�������
fsoInput.Close
fsoOutput.Close

Call subWriteLog("[" & strInputFileName & "]�̑S" & rCnt & "�s����" & wCnt & "�s�����o�����܂����B" )
Call subWriteLog("�t�@�C���ǂݍ��݁A�����o���I���B")


' ------------------------------------------------------
'              �S�D�Â��t�@�C���폜�A���l�[��
' ------------------------------------------------------

' �ǂݍ��񂾃t�@�C�����폜
FSO.DeleteFile strInputFileName
If Err.Number <> 0 Then
	Call subErrHandling("�ǂݍ��݃t�@�C��[" & strInputFileName & "]�̍폜�����s���܂����B")
	WScript.Quit 9		' �G���[�l:9�Ƃ���
End If
Call subWriteLog("�ǂݍ��݃t�@�C��[" & strInputFileName & "]���폜�B" )


' �����o�����t�@�C�������l�[��
REG.Pattern = "_2\.txt$"
Dim fsoFile, strFileName
Set fsoFile = FSO.GetFile( strOutputFileName )
strFileName = FSO.GetFileName( strOutputFileName )
fsoFile.Name = REG.Replace(strFileName, ".txt")
If Err.Number <> 0 Then
	Call subErrHandling("�����o�����t�@�C��[" & strOutputFileName & "]�̃��l�[���Ɏ��s���܂����B")
	WScript.Quit 9		' �G���[�l:9�Ƃ���
End If

Call subWriteLog("�����o���t�@�C��[" & strOutputFileName & "]��[" & REG.Replace(strFileName, ".txt") & "]�Ƀ��l�[���B" )


' ------------------------------------------------------
'                  �T�D���[���𑗐M����
' ------------------------------------------------------
Call subSendMail(strNormalSubject, "����I�����܂����B")
Call subWriteLog("���[���𑗐M���܂����B")


' ------------------------------------------------------
'                    �U�D�I      ��
' ------------------------------------------------------
Set FSO       = Nothing 
Set REG       = Nothing
Set fsoInput  = Nothing
Set fsoOutput = Nothing
Set fsoFile   = Nothing

Call subWriteLog("�v���O��������I���B")

WScript.Quit 0



' ======================================================================================
' ======================================================================================




' ------------------------------------------------------
'                    ���O�����o������
' ------------------------------------------------------
Sub subWriteLog(strMessage)
	' �T�v�F���O�t�@�C���ɏ������݂�����

	Dim fsoLog
	Set fsoLog = FSO.OpenTextFile(LOG_FILE_NAME, 8, True)
	fsoLog.WriteLine strMessage

	fsoLog.Close
	Set fsoLog = Nothing

End Sub

' ------------------------------------------------------
'                    �e�L�X�g���e�ύX
' ------------------------------------------------------
Function fncChkLineStr(strLine)
	' �T�v�F�e�L�X�g�̓��e��ϗe������

	' ��TODO:�I���W�i���̉��H���@�������ɋL�ڂ��邱��

	fncChkLineStr = True
End Function

' ------------------------------------------------------
'                      ���[�����M
' ------------------------------------------------------
Sub subSendMail(strSubject, strTextBody)
	' �T�v�F���[�����M
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
'                   �G���[�n���h�����O
' ------------------------------------------------------
Sub subErrHandling(strMsg)
	' �T�v�F���O�t�@�C�������l�[���A�G���[�̓��e�̃��[������

	' ���O���L�ڌ�A���l�[������
	Call subWriteLog(strMsg)
	Call subRenameLogFile()

	' ���[���𑗂�
	Call subSendMail(strErrSubject, strMsg)

End Sub



' ------------------------------------------------------
'                  ���O�t�@�C�������l�[��
' ------------------------------------------------------
Sub subRenameLogFile()
	' �T�v�F���O�t�@�C��������t���̖��̂�Rename

	If FSO.FileExists( LOG_FILE_NAME ) Then
		REG.Pattern = "\.vbs$"
		Dim fsoLogFile
		Set fsoLogFile = FSO.GetFile( LOG_FILE_NAME )
		fsoLogFile.Name = REG.Replace(WScript.ScriptName, "_" & fncGetStringDateTime() & ".log")
		Set fsoLogFile = Nothing
	End If

End Sub

' ------------------------------------------------------
'              �����̓����̕������Ԃ�
' ------------------------------------------------------
Function fncGetStringDateTime()
	' �T�v�F���s������Ԃ��i�G���[���O�p�j
	Dim strDate
	Dim strTime
	
	strDate = Replace(FormatDateTime(Now, 1),"/","")
	strTime = Replace(FormatDateTime(Now, 4),":","")

	' �uyyyyMMdd_hhmm�v�̌`���ŕԂ�
	fncGetStringDateTime = strDate & "_" & strTime

End Function
