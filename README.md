
Sub login()
    Dim objie As InternetExplorer
    Dim uid As String
    Dim pwd As String
    Dim rng As Range
    
    Set rng = Sheets(1).Range("B2")
    
    If Trim(rng.Value) = "" Or Trim(rng.Offset(1).Value) = "" Then
        MsgBox "USER ID AND PWD are mandatory."
        Exit Sub
    End If
    
    On Error Resume Next
    objie.Quit
    Set objie = Nothing
    On Error GoTo 0
    
    Set objie = New InternetExplorer 'Initialize internet object
    objie.Visible = True
    
    uid = rng.Value
    pwd = rng.Offset(1).Value
    
    objie.navigate "http://www.gmail.com"
    Do While objie.Busy = True
        DoEvents
    Loop
    
    Do While objie.Busy: Loop
    Do While objie.readyState <> READYSTATE_COMPLETE: Loop
    Do While objie.Busy: Loop
    
    objie.document.forms.gaia_loginform.Email.Value = uid
    objie.document.forms.gaia_loginform.Passwd.Value = pwd
    objie.document.forms.gaia_loginform.signIn.Click
    
    Do While objie.Busy = True
        DoEvents
    Loop
    
    Do While objie.Busy: Loop
    Do While objie.readyState <> READYSTATE_COMPLETE: Loop
    Do While objie.Busy: Loop
    
    Set objie = Nothing
    
    MsgBox "DONE..."
End Sub

Sub delay(ByVal interval As String)
Dim waituntil
waituntil = Now + TimeValue("00:00:" & interval)
Do
    DoEvents
Loop Until Now >= waituntil
End Sub
