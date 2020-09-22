VERSION 5.00
Object = "*\AEasyButton.vbp"
Begin VB.Form frmHowTo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "How to draw a picture"
   ClientHeight    =   6405
   ClientLeft      =   3000
   ClientTop       =   1320
   ClientWidth     =   5760
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmHowTo.frx":0000
   ScaleHeight     =   427
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   384
   ShowInTaskbar   =   0   'False
   Begin EasyX.EasyButton EasyButton1 
      Height          =   525
      Left            =   4275
      TabIndex        =   9
      Top             =   5805
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   926
      Caption         =   "&Close"
      Align           =   0
      Picture         =   "frmHowTo.frx":2910
      BackColor       =   -2147483633
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
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   -90
      TabIndex        =   7
      Top             =   2700
      Width           =   6225
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -45
      TabIndex        =   6
      Top             =   5625
      Width           =   6090
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clickable area  (Black)"
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   6
      Left            =   3960
      TabIndex        =   11
      Top             =   5355
      Width           =   1605
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      Index           =   7
      X1              =   219
      X2              =   261
      Y1              =   327
      Y2              =   363
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      Index           =   6
      X1              =   288
      X2              =   312
      Y1              =   282
      Y2              =   282
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      Index           =   5
      X1              =   253
      X2              =   297
      Y1              =   321
      Y2              =   321
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No clickable      area      (any color)"
      ForeColor       =   &H00008000&
      Height          =   600
      Index           =   5
      Left            =   4545
      TabIndex        =   10
      Top             =   4635
      Width           =   885
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   0
      TabIndex        =   8
      Top             =   2385
      Width           =   5730
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Height          =   2445
      Left            =   1980
      Top             =   2880
      Width           =   1995
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      Index           =   4
      X1              =   84
      X2              =   129
      Y1              =   354
      Y2              =   354
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Same Picture"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   4
      Left            =   225
      TabIndex        =   5
      Top             =   5175
      Width           =   945
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mask (White)"
      ForeColor       =   &H00008000&
      Height          =   390
      Index           =   3
      Left            =   4725
      TabIndex        =   4
      Top             =   4095
      Width           =   675
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      Index           =   3
      X1              =   252
      X2              =   288
      Y1              =   309
      Y2              =   282
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MouseDown"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   2
      Left            =   495
      TabIndex        =   3
      Top             =   4230
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   2
      X1              =   96
      X2              =   156
      Y1              =   291
      Y2              =   291
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MouseOver"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   1
      Left            =   4320
      TabIndex        =   2
      Top             =   3645
      Width           =   825
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   1
      X1              =   237
      X2              =   285
      Y1              =   252
      Y2              =   252
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MouseOut"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   0
      Left            =   585
      TabIndex        =   1
      Top             =   3105
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   0
      X1              =   93
      X2              =   153
      Y1              =   216
      Y2              =   216
   End
   Begin VB.Image Image1 
      Height          =   2100
      Left            =   2160
      Picture         =   "frmHowTo.frx":5F74
      Top             =   3015
      Width           =   1665
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2220
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   5055
   End
End
Attribute VB_Name = "frmHowTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EasyButton1_Click()

    Unload Me

End Sub

Private Sub Form_Load()
    
    Label1 = "You can make buttons in any format." & Chr(10)
    Label1 = Label1 & "(The secret is the picture)" & Chr(10) & Chr(10)
    Label1 = Label1 & "How to make a picture:" & Chr(10)
    Label1 = Label1 & "     You must draw 4 images of same size in a only image." & Chr(10)
    Label1 = Label1 & "     The first image is 'mouse out'." & Chr(10)
    Label1 = Label1 & "     Second image is 'mouse over'." & Chr(10)
    Label1 = Label1 & "     Third is 'mouse down'" & Chr(10)
    Label1 = Label1 & "     Fourth is a mask. The mask define the clickable area." & Chr(10)
    Label3 = "See image below"

End Sub
Private Sub Form_Resize()
    
    For X = 1 To ScaleWidth Step 100
        For Y = 1 To ScaleHeight Step 100
            PaintPicture Picture, X, Y
        Next
    Next

End Sub


