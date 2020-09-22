VERSION 5.00
Begin VB.Form frmAboutBox 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2490
   ClientLeft      =   3015
   ClientTop       =   2535
   ClientWidth     =   4485
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAboutBox.frx":0000
   ScaleHeight     =   2490
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   Begin EasyX.EasyButton EasyButton1 
      Height          =   360
      Left            =   3240
      TabIndex        =   0
      Top             =   2025
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   635
      Caption         =   "&Close"
      Align           =   0
      Picture         =   "frmAboutBox.frx":C6CC
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      AccessKey       =   "C"
   End
End
Attribute VB_Name = "frmAboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub CenterForm(Frm)
        
    Frm.Top = (Screen.Height / 2) - (Frm.Height / 2)
    Frm.Left = (Screen.Width / 2) - (Frm.Width / 2)

End Sub

Private Sub EasyButton1_Click()
    
    Unload frmAboutBox
    Set frmAboutBox = Nothing

End Sub

Private Sub Form_Load()

    CenterForm Me

End Sub


