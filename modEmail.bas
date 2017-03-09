Attribute VB_Name = "modEmail"
Option Explicit

Sub SendTestMail()
    frmMain.MailUC.FromEmail = "donotreply@odysseyclassic.com"
    frmMain.MailUC.FromName = "Odyssey Classic"
    frmMain.MailUC.ToEmail = "remote@codemallet.com"
    frmMain.MailUC.ToName = "test"
    frmMain.MailUC.Server = "127.0.0.1"
    frmMain.MailUC.Subject = "Recovery Test"
    'MUC.UserName = TDat(6).Text
    'MUC.Password = TDat(7).Text
    frmMain.MailUC.Port = 25
    frmMain.MailUC.PrepareEmail "Recovery Text", ""
    Dim sErr As String, LRet As Long
    LRet = frmMain.MailUC.ConnectAndSend(sErr)
    If (LRet <> 100) Then MsgBox sErr
End Sub

Private Sub MUC_OnSendComplete(Success As Boolean, sErr As String)
    MsgBox "Mail successful: " & CStr(Success) & "--" & sErr
End Sub

Function EncryptString(St As String) As String
Dim TempStr As String, TempStr2 As String
Dim A As Integer, TmpNum As Integer

TempStr = ""
TempStr2 = ""

For A = 1 To Len(St)
    TempStr = Mid$(St, A, 1)
    TmpNum = Asc(TempStr)
    TempStr2 = TempStr2 + Chr$(TmpNum + 3 - 10)
Next A

EncryptString = Trim$(TempStr2)
End Function
Function DecipherString(St As String) As String
Dim TempStr As String, TempStr2 As String
Dim A As Integer, TmpNum As Integer

TempStr = ""
TempStr2 = ""

For A = 1 To Len(St)
    TempStr = Mid$(St, A, 1)
    TmpNum = Asc(TempStr)
    TempStr2 = TempStr2 + Chr$(TmpNum - 3 + 10)
Next A

DecipherString = Trim$(TempStr2)
End Function

