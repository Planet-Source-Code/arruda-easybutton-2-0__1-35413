VERSION 5.00
Object = "*\AEasyButton.vbp"
Begin VB.Form Test 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3795
   ClientLeft      =   1785
   ClientTop       =   1650
   ClientWidth     =   5730
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Test.frx":0000
   ScaleHeight     =   3795
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   Begin EasyX.EasyButton EasyButton11 
      Height          =   375
      Left            =   4140
      TabIndex        =   10
      Top             =   3195
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      Caption         =   "&Close"
      Align           =   0
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
   Begin EasyX.EasyButton EasyButton7 
      Height          =   525
      Left            =   225
      TabIndex        =   5
      Top             =   3105
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   926
      Caption         =   "About..."
      Align           =   0
      Picture         =   "Test.frx":C042
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
   End
   Begin EasyX.EasyButton EasyButton4 
      Height          =   615
      Left            =   3105
      TabIndex        =   3
      Top             =   315
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   1085
      Caption         =   "               Exotic"
      SoundFileName   =   "D:\VB60\PSC\EASYBU~1.0\C_BANG.WAV"
      Align           =   0
      Picture         =   "Test.frx":F6A6
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   65280
   End
   Begin EasyX.EasyButton EasyButton1 
      Height          =   360
      Left            =   4230
      TabIndex        =   0
      Top             =   2025
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   635
      Caption         =   "&Metal"
      SoundFileName   =   "D:\VB60\PSC\EASYBU~1.0\C_BANG.WAV"
      Align           =   0
      Picture         =   "Test.frx":15F86
      BackColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AccessKey       =   "M"
   End
   Begin EasyX.EasyButton EasyButton5 
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   2340
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      Caption         =   "Gel Button"
      SoundFileName   =   "D:\VB60\PSC\EASYBU~1.0\C_BANG.WAV"
      Align           =   0
      Picture         =   "Test.frx":1805A
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
      ForeColor       =   4210752
   End
   Begin EasyX.EasyButton EasyButton8 
      Height          =   330
      Left            =   2295
      TabIndex        =   6
      Top             =   3195
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      Caption         =   "Glass"
      Align           =   0
      Picture         =   "Test.frx":1AD4E
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
      ForeColor       =   4210752
   End
   Begin EasyX.EasyButton EasyButton10 
      Height          =   615
      Left            =   2520
      TabIndex        =   8
      Top             =   1125
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Caption         =   "02"
      Align           =   0
      Picture         =   "Test.frx":1CD1E
      BackColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   13817821
   End
   Begin EasyX.EasyButton EasyButton9 
      Height          =   495
      Left            =   180
      TabIndex        =   7
      Top             =   180
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   873
      Caption         =   "Wood Button"
      Align           =   0
      Picture         =   "Test.frx":1EDA2
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
      ForeColor       =   7312816
   End
   Begin EasyX.EasyButton EasyButton6 
      Height          =   630
      Left            =   495
      TabIndex        =   4
      Top             =   2160
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1111
      Caption         =   "01"
      Align           =   0
      Picture         =   "Test.frx":229A6
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
      ForeColor       =   16771549
   End
   Begin EasyX.EasyButton EasyButton3 
      Height          =   375
      Left            =   4140
      TabIndex        =   2
      Top             =   2610
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      Caption         =   "How To..."
      SoundFileName   =   "D:\VB60\PSC\EASYBU~1.0\C_BANG.WAV"
      Align           =   0
      Picture         =   "Test.frx":24ADA
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
   End
   Begin EasyX.EasyButton EasyButton2 
      Height          =   525
      Left            =   180
      TabIndex        =   1
      Top             =   1035
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   926
      Caption         =   "Wood Button"
      Align           =   0
      Picture         =   "Test.frx":272C2
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
      ForeColor       =   7312816
   End
   Begin VB.Image Image4 
      Height          =   1920
      Left            =   0
      Picture         =   "Test.frx":2B456
      Top             =   1935
      Width           =   1920
   End
   Begin VB.Image Image3 
      Height          =   1920
      Left            =   3870
      Picture         =   "Test.frx":37498
      Top             =   0
      Width           =   1920
   End
   Begin VB.Image Image2 
      Height          =   1920
      Left            =   1935
      Picture         =   "Test.frx":434DA
      Top             =   1935
      Width           =   1920
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   1935
      Picture         =   "Test.frx":4F51C
      Top             =   0
      Width           =   1920
   End
End
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EasyButton11_Click()

    End

End Sub

Private Sub EasyButton3_Click()

    frmHowTo.Show 1
    Unload frmHowTo
    Set frmHowTo = Nothing

End Sub


Private Sub EasyButton7_Click()

    EasyButton7.AboutBox

End Sub
