VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Key Keeper"
   ClientHeight    =   3675
   ClientLeft      =   2835
   ClientTop       =   2745
   ClientWidth     =   5145
   Icon            =   "frmMain1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5145
   Begin VB.ComboBox Buildbox 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   12
      ToolTipText     =   "Displays a list of builds for the above Windows version."
      Top             =   720
      Width           =   4215
   End
   Begin VB.ComboBox Productbox 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   11
      ToolTipText     =   "Displays a list of Windows versions."
      Top             =   120
      Width           =   4215
   End
   Begin VB.ListBox Keylist 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   600
      TabIndex        =   7
      ToolTipText     =   "Displays all possible keys for the selected build."
      Top             =   1440
      Width           =   4455
   End
   Begin VB.TextBox biostext 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "Displays the BIOS date required to install that build of Windows so you will have a few days before the timebomb kicks in."
      Top             =   2520
      Width           =   3975
   End
   Begin VB.TextBox Passtext 
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Some build of Windows require a password to use. The password will be returned here."
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmMain1.frx":0442
      Top             =   3240
      Width           =   480
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   5640
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label PreRelLabel 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label preinstall 
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Does Windows need to be preinstalled? "
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "BIOS Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Pass:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Key:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5520
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label2 
      Caption         =   "Build:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Product:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Menu tkkFile 
      Caption         =   "&Tools"
      Begin VB.Menu makeBootimage 
         Caption         =   "&Make bootable disk image"
         Shortcut        =   {F2}
      End
      Begin VB.Menu ViewClipBoard 
         Caption         =   "&View clipboard contents"
      End
      Begin VB.Menu bar2 
         Caption         =   "-"
      End
      Begin VB.Menu updatechecker 
         Caption         =   "Check for &updates"
         Shortcut        =   {F5}
      End
      Begin VB.Menu bar 
         Caption         =   "-"
      End
      Begin VB.Menu exitcmd 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu usrkey 
      Caption         =   "&User Keys"
   End
   Begin VB.Menu aboutmenu 
      Caption         =   "&Help"
      Begin VB.Menu Helpcmd 
         Caption         =   "Help &Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu aboutKeyKeeper 
         Caption         =   "&About The Key Keeper"
      End
   End
   Begin VB.Menu PopUpBlocker 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu CopyPasta 
         Caption         =   "&Copy to clipboard"
      End
      Begin VB.Menu ViewCBContents 
         Caption         =   "&View clipboard contents"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Code (c) 2010 The Little Beige Box
' http://www.beige-box.com


Private Sub Productbox_click()
Select Case Productbox.ListIndex
Case 0
 Open "e:\bpk\Beta Key Keeper\cfg\chicago_builds.bld" For Input As #2
 Do Until EOF(2)
   Input #2, Buildstr
   Buildbox.AddItem Buildstr
 Loop
 Close #2
End Select
End Sub

Private Sub Form_Load()

PreRelLabel.Caption = "The Key Keeper | Build " & App.Revision & " | v" & App.Major & "." & App.Minor

On Error GoTo errhandle

Open "E:\bpk\Beta Key Keeper\cfg\codenames.cfg" For Input As #1
Do Until EOF(1)
  Input #1, Codenamesstr
  Productbox.AddItem Codenamesstr
Loop
Close #1

Productbox.ListIndex = 0
'Buildbox.ListIndex = 0

errhandle:
Select Case Err
Case Is <> 0
 errorbox = MsgBox("Could not load configuration files.", vbCritical, errorbox)
 End
End Select

End Sub


Private Sub clipboarddo()
 Clipboardviewer.Show
End Sub

Private Sub Helpcmd_Click()
 Helpcmd1 = MsgBox("Helpcmd", vbInformation, Helpcmd1)
End Sub



Private Sub updatechecker_Click()
' updatechk = MsgBox("Updatechecker", vbInformation, updatechk)
frmUpdate.Show
End Sub

Private Sub viewcbcontents_click()
 clipboarddo
End Sub
Private Sub CopyPasta_Click()
   Clipboard.Clear
   Clipboard.SetText List3.List(List3.ListIndex)
   
End Sub
Private Sub List3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button 'Button = vbRightButton Then
Case vbRightButton
 Select Case List3
 Case ""
 
 Case Else
  PopupMenu PopUpBlocker, vbPopupMenuRightButton
  'End If
 End Select
End Select
End Sub

Private Sub makeBootimage_Click()
On Error GoTo ErrorHandling
' Errbox = MsgBox("Place holder.", vbCritical, "Error")
RaWrite = Shell("rawrite -f boot.img -d a", vbNormalFocus)

ErrorHandling:
Select Case Err
Case Is <> 0
 ErrDialog = MsgBox("Could not launch RAWRITE." & vbNewLine & "Check to make sure that RAWRITE.EXE is in the Key Keeper's program directory." & vbNewLine & "If it is not there, re-install The Key Keeper." & vbNewLine & "If it is there, then another error occurred.", vbCritical, "Could not load RAWRITE")
End Select
End Sub
Private Sub exitcmd_Click()
 End
End Sub
Private Sub aboutKeyKeeper_Click()
 frmAbout.Show
End Sub

Private Sub ViewClipBoard_Click()
 clipboarddo
End Sub



