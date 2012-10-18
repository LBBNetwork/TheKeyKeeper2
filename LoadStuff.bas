Attribute VB_Name = "LoadStuff"
Dim args(50) As String
Sub DoLoad()
'Dim args(50) As String

nf = FreeFile

Open "E:\bpk\Beta Key Keeper\cfg\keys.db" For Input As nf
'Do Until EOF(nf)
'Line Input #nf, tmp$
'Loop

'Select Case LCase$(args(0))
'Case "b"
' ShowKeys
'End Select

Do Until EOF(nf)
 Line Input #nf, tmp$
  tmp$ = LTrim$(tmp$)
  If Left$(tmp$, 1) <> "#" Then
    carg = 0
    Do Until Len(tmp$) = 0
      xp = InStr(1, tmp$, ":")
      If xp < 1 Then xp = Len(tmp$) + 1
      args(carg) = Left$(tmp$, xp - 1)
      tmp$ = Mid$(tmp$, xp + 1)
      carg = carg + 1
    Loop
      Select Case LCase$(args(0))
        Case "b"
          ShowKeys
          'frmMain.Keylist.AddItem args(2)
      End Select
  End If
Loop
End Sub

Sub ShowKeys()
Select Case frmMain.Buildbox.ListIndex
Case 0
 Select Case args(1)
  Case "56"
   frmMain.Keylist.AddItem args(2)
 End Select
End Select
End Sub
