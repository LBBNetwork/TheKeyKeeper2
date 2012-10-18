VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About The Key Keeper"
   ClientHeight    =   4395
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3033.508
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   3840
      Width           =   1260
   End
   Begin VB.Label Label2 
      Caption         =   "http://microsoftcollectionbook.com"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   3360
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Some info provided by:"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   1080
      Picture         =   "frmAbout.frx":0884
      Top             =   2040
      Width           =   4185
   End
   Begin VB.Label lblDescription 
      Caption         =   "Program (c) 2010 The Little Beige Box, www.beige-box.com"
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   1080
      TabIndex        =   2
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "The Key Keeper"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   780
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About The Key Keeper"
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & " (Build " & App.Revision & ")"
    lblTitle.Caption = "The Key Keeper"

End Sub


