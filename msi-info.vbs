Const msiOpenDatabaseModeReadOnly     = 0

' Store the arguments in a variable:
Set objArgs = Wscript.Arguments

'Count the arguments

If objArgs.Count <> 1 Then
    MsgBox "Need to pass a file"
    WScript.Quit
end if

On Error Resume Next ' defer error handling

Dim installer
Set installer = CreateObject("WindowsInstaller.Installer")

Dim database
Set database = installer.OpenDatabase(objArgs(0), msiOpenDatabaseModeReadOnly)

' test for error
If Err.Number <> 0 Then
    Dim message, errorRec
    message = Err.Source & " " & Hex(Err.Number) & ": " & Err.Description
    If Not installer Is Nothing Then
        ' try to obtain extended error info
        Set errorRec = installer.LastErrorRecord
        If Not errorRec Is Nothing Then message = message & vbNewLine & errorRec.FormatText
    End If

    MsgBox message
End If

Dim View, Record
Set View = database.OpenView("SELECT Property, Value FROM Property") 
View.Execute
Do
 Set Record = View.Fetch
 If Record Is Nothing Then Exit Do


    If (StrComp(UCase(Record.StringData(1)), Record.StringData(1)) = 0) Then
        Wscript.Echo  Record.StringData(1) + "=" +  Record.StringData(2)
    End If
Loop
Set View = Nothing
