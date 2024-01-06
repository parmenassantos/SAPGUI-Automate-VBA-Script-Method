'
'   version: 2023-12-28
'   Created by Parmenas Santos
'   parmenassantos@gmail.com
'
'   GitHub repository - check for updates
'   https://github.com/parmenassantos/SAPGUI-Auto-VBA-Method.git
'
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Copy and paste data below in your SUB end enjoy.
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Sub Example()
'Dim Session As Object
'Set Session = uSession("LOGON_SAP", "TRANSACTION")
'If Session Is Nothing Then
'    MsgBox "Session is empty.", vbInformation, "SCRIPT: Error Into Session Script"
'    Exit Sub
'End If
'    Write your code below. Enjoy!!
' ...
'End Sub
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Const SAPLOGON = "C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe"
Private Const NOTFOUND = "Not Found"

' LOGON_SAP
Public Const uECC = "1. ECC - Produção (DFP)"
Public Const uEWM = "2. EWM - Produção (EWP)"
Public Const uEvent = "3. Event - Produção (EMP)"
Public Const uGRC = "4. GRC - Produção (NFP)"
Public Const uBW = "5. BW - Produção (BWP)"
Public Const uS4P = "6. S4P - Produção LATAM"

'Transaction EWM
Public Const MONITOR = "/n/scwm/mon"
Public Const CONS_ESTORNO = "/n/DAFITI/EST_CONS"

'Transaction ECC
Public Const PRINTNF = "j1bnfe"
Public Const ZMM017 = "ZMM017"
Public Const MIRO = "MIRO"

'Don't Touch Me
Private Const SYS1001 = "SAPMSYST"
Private Const SYS1002 = "SAPLSMTR_NAVIGATION"
Private Const SYS2001 = "SESSION_MANAGER"
Private Const MaxSession = 6

'DIMS
Private EnConnnection        As Integer
Private Num                  As Integer
Private Windows              As Integer
Private WshShell             As Object
Private SapGuiAuto           As Object
Private SapApplication       As Object
Private InitSession          As Object
Private Init                 As Object
Private EachSession          As Object
Private SYS2000              As String
Private WDS                  As Variant
Private credentialsSaved     As Boolean

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function will set or will open end  will set SAPGUI as Object.
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function SapGui() As Object
On Error GoTo Error
If SapGui Is Nothing Then
    Set SapGui = GetObject("SAPGUI")
    Exit Function
Error:
    Shell SAPLOGON, vbnormalfocus
    Set WshShell = CreateObject("WScript.Shell")
    Do Until WshShell.AppActivate("SAP Logon")
        Application.Wait Now + TimeValue("00:00:01")
    Loop
    Set SapGui = GetObject("SAPGUI")
    Set WshShell = Nothing
End If
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function will open or set appropriate Connection to use.
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function SapConnection(ByVal LOGONSAP As String) As Object
Set SapGuiAuto = SapGui
Set SapApplication = SapGuiAuto.GetScriptingEngine
Num = SapApplication.Connections.Length
EnConnnection = 0
Do While EnConnnection < Num
    Set SapConnection = SapApplication.Children(CInt(EnConnnection))
    If SapConnection.Description = LOGONSAP Then
        Exit Function
    End If
    EnConnnection = EnConnnection + 1
Loop
Set SapConnection = Nothing
Set SapConnection = SapApplication.OpenConnection(LOGONSAP, True)
Set SapConnection = SapApplication.Children(CInt(Mid(SapConnection.Name, 5, 1)))
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function will open or set appropriate Session.
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function uSession(ByVal LOGONSAP As String, Transaction As String) As Object
Set InitSession = SapConnection(LOGONSAP)
Windows = InitSession.Children.Length
Set Init = InitSession.Children(Windows - 1)
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function will connect end will open search appropriate Connection to Session.
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
If Windows = 1 And Init.info.program = SYS1001 Then
    Set uSession = InitSession.Children(Windows - 1)
    uSession.findById("wnd[0]").maximize
    credentialsSaved = CheckCredentialsSaved(LOGONSAP)
    If Not credentialsSaved Then
        SaveCredentials LOGONSAP, uSession
    End If
    LoadCredentials LOGONSAP, uSession
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Function will set opened Session in "SESSION_MANAGER".
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
ElseIf Windows < 7 And Init.info.program <> SYS1001 Then
    For Each EachSession In InitSession.Children
        SYS2000 = EachSession.info.Transaction
        If SYS2000 = SYS2001 Then
            WDS = Mid(EachSession.Name, 5, 1)
            Set uSession = InitSession.Children(CInt(WDS) + 0)
            Exit For
        End If
    Next EachSession
    If Windows = MaxSession And WDS = Empty Then
        MsgBox "Há " & Windows & " sessões sendo utilizadas, impossível continuar." & vbNewLine & _
        "Fecha 1 ou mais janelas para poder continuar.", vbInformation, LOGONSAP
        Exit Function
    End If
End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   If it have Session opened but nothing in transaction: "SESSION_MANAGER", then set new Session as Object.
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
If uSession Is Nothing Then
    Set uSession = InitSession.Children(Windows - 1)
    uSession.createsession
    On Error Resume Next
    Do Until Z = SYS1002
        Set uSession = InitSession.Children(Windows + 0)
        Z = uSession.info.program
    Loop
    On Error GoTo 0
End If
uSession.findById("wnd[0]").maximize
uSession.findById("wnd[0]/tbar[0]/okcd").Text = Transaction
uSession.findById("wnd[0]").sendVKey 0
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Credentials schema below.
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function Session(ByVal LOGONSAP As String, Transaction As String) As Object
uSession.findById("wnd[0]").maximize
credentialsSaved = CheckCredentialsSaved(LOGONSAP)
If Not credentialsSaved Then
    SaveCredentials LOGONSAP, uSession
End If
LoadCredentials LOGONSAP, uSession
uSession.findById("wnd[0]/tbar[0]/okcd").Text = Transaction
uSession.findById("wnd[0]").sendVKey 0
End Function

Private Function CheckCredentialsSaved(LOGONSAP As String) As Boolean
Dim filePath As String
filePath = Environ("APPDATA") & "\SAPCredentials_" & LOGONSAP & ".txt"
CheckCredentialsSaved = Dir(filePath) <> ""
End Function

Private Sub SaveCredentials(LOGONSAP As String, uSession As Object)
Dim filePath As String
Dim username As String
Dim password As String
filePath = Environ("APPDATA") & "\SAPCredentials_" & LOGONSAP & ".txt"
username = InputBox("Enter SAP username:")
password = InputBox("Enter SAP password:")
Open filePath For Output As #1
Print #1, username
Print #1, password
Close #1
End Sub

Private Sub LoadCredentials(LOGONSAP As String, uSession As Object)
Dim filePath As String
Dim username As String
Dim password As String
filePath = Environ("APPDATA") & "\SAPCredentials_" & LOGONSAP & ".txt"
If Dir(filePath) <> "" Then
    Open filePath For Input As #1
    Line Input #1, username
    Line Input #1, password
    Close #1

    On Error Resume Next
    uSession.findById("wnd[0]/usr/txtRSYST-BNAME").Text = username
    uSession.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = password
    uSession.findById("wnd[0]").sendVKey 0
    On Error GoTo 0
End If
End Sub


