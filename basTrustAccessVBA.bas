Attribute VB_Name = "basTrustAccessVBA"
Option Explicit


' http://www.mrexcel.com/forum/excel-questions/550141-there-way-programatically-check-trust-access-vbulletin-project-setting-2.html
Sub ToggleTrust()
    On Error Resume Next
    
'    Debug.Print DBG_MSG; "Toggle VBA object access settings (current: "; IsVBATrusted; ")"
    LogEvent "Toggle VBA object access settings (current: " & IsVBATrusted & ")", LTP.INFO_L, "ToggleTrust"
    AppActivate Application.Caption
    Call SendKeys("%TMS{TAB}{TAB} {ENTER}")
    DoEvents
'    Debug.Print DBG_MSG; "Toggle VBA object access settings (current: "; IsVBATrusted; ")"
    LogEvent "Toggle VBA object access settings (current: " & IsVBATrusted & ")", LTP.INFO_L, "ToggleTrust"
    
End Sub

Function IsVBATrusted() As Boolean

    Const HKEY_CURRENT_USER = &H80000001
    'Const HKEY_USERS = &H80000003

    Dim sComputer As String
    Dim sOfficeVersion As String
    Dim sKeyPath As String
    Dim sValueName As String
    Dim lValue As Long
    Dim oReg As Object
    
    sOfficeVersion = Application.Version
    sComputer = "."
    
    Set oReg = GetObject( _
       "winmgmts:{impersonationLevel=impersonate}!\\" & _
        sComputer & "\root\default:StdRegProv")
        
    sKeyPath = "Software\Microsoft\Office\" & _
    sOfficeVersion & "\Excel\Security"
    
    sValueName = "AccessVBOM"
    
    oReg.GetDWORDValue HKEY_CURRENT_USER, _
    sKeyPath, sValueName, lValue
    
    IsVBATrusted = lValue
End Function
