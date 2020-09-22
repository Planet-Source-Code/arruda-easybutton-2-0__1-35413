VERSION 5.00
Begin VB.UserControl EasyButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2340
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HitBehavior     =   0  'None
   ScaleHeight     =   47
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   156
   ToolboxBitmap   =   "EasyButton.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1530
      Top             =   1620
   End
End
Attribute VB_Name = "EasyButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As Long, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private Const SND_ASYNC = &H1
Private Const SND_MEMORY = &H4
Private Const SND_NOWAIT = &H2000
Private Const DT_BOTTOM = &H8
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_NOCLIP = &H100
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_DEFAULT = DT_CENTER Or DT_VCENTER

Public Enum Sounds
    [None]
    [Default]
    [FromFile]
End Enum

Public Enum Alignment
    [CenterCenter]
    [CenterTop]
    [CenterBottom]
    [LeftCenter]
    [LeftTop]
    [LeftBottom]
    [RightCenter]
    [RightTop]
    [RightBottom]
End Enum

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Dim Rt As RECT
Dim Pt As POINTAPI
Dim DC As Long
Dim Obj As Long
Dim MouseOver As Boolean
Dim MouseButton As Integer
Dim ButtonState As Integer
Dim PtIn As Boolean
Dim PicHeight As Integer
Dim PicWidth As Integer
Dim mCaption As String
Dim mAlign As Alignment
Dim HasPicture As Boolean
Dim B() As Byte
Dim mUseSound As Sounds
Dim mSoundFileName As String

Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552

    frmAboutBox.Show 1

End Sub
Public Property Get Caption() As String

    Caption = mCaption

End Property
Public Property Get SoundFileName() As String

    SoundFileName = mSoundFileName

End Property
Public Property Get UseSound() As Sounds

    UseSound = mUseSound

End Property
Public Property Get Enabled() As Boolean

    Enabled = UserControl.Enabled

End Property
Public Property Let Caption(ByVal newCaption As String)

    mCaption = newCaption
    PropertyChanged "Caption"
    Cls
    If ButtonState = 0 Then
        DrawMouseOut
    ElseIf ButtonState = 1 Then
        DrawUp
    Else
        DrawDown
    End If
    
End Property
Public Property Let SoundFileName(ByVal newSound As String)

    On Error Resume Next
    If Dir(newSound) = "" Then
        MsgBox "File not found", vbCritical, App.Title
        mUseSound = None
        mSoundFileName = ""
        Exit Property
    End If
    If Trim(newSound) = "" Then
        mUseSound = None
        mSoundFileName = ""
    Else
        mSoundFileName = newSound
        Open mSoundFileName For Binary As #1
        ReDim B(LOF(1))
        Get #1, , B
        Close #1
    End If
    PropertyChanged "SoundFileName"
    
End Property
Public Property Let UseSound(ByVal newData As Sounds)

    mUseSound = newData
    PropertyChanged "UseSound"
    
End Property
Public Property Let AccessKey(ByVal newKey As String)

    UserControl.AccessKeys() = newKey
    PropertyChanged "AccessKey"
    
End Property
Public Property Let Enabled(ByVal newEnabled As Boolean)

    UserControl.Enabled() = newEnabled
    Cls
    If ButtonState = 0 Then
        DrawMouseOut
    ElseIf ButtonState = 1 Then
        DrawUp
    Else
        DrawDown
    End If
    PropertyChanged "Enabled"
    
End Property
Public Property Let Align(ByVal newAlign As Alignment)

    mAlign = newAlign
    Cls
    If ButtonState = 0 Then
        DrawMouseOut
    ElseIf ButtonState = 1 Then
        DrawUp
    Else
        DrawDown
    End If
    PropertyChanged "Align"

End Property
Private Sub DrawMouseOut()
    
    On Error Resume Next
    If Not HasPicture Then
        Cls
        UserControl.Line (0, 0)-Step(ScaleWidth - 1, ScaleHeight - 1), &HCFCFCF, B
    Else
        BitBlt hdc, 0, 0, PicWidth, (PicHeight / 4), DC, 0, 0, vbSrcCopy
    End If
    If Trim(mCaption) <> "" Then
        Rt.Left = 0
        Rt.Top = 0
        Rt.Bottom = ScaleHeight
        Rt.Right = ScaleWidth
        OldColor = UserControl.ForeColor
        If Not UserControl.Enabled Then UserControl.ForeColor = vbGrayText
        DrawText hdc, mCaption, Len(mCaption), Rt, GetAlign(mAlign) Or DT_NOCLIP Or DT_SINGLELINE
        UserControl.ForeColor = OldColor
    End If
    Refresh
    ButtonState = 0

End Sub
Private Sub DrawUp()
    
    If ButtonState = 1 Then Exit Sub
    If Not HasPicture Then
        Cls
        UserControl.Line (0, 0)-Step(ScaleWidth - 1, 0), vb3DHighlight
        UserControl.Line (0, 0)-Step(0, ScaleHeight - 1), vb3DHighlight
        UserControl.Line (0, ScaleHeight - 1)-Step(ScaleWidth, 0), vb3DDKShadow
        UserControl.Line (ScaleWidth - 1, 0)-Step(0, ScaleHeight - 1), vb3DDKShadow
    Else
        BitBlt hdc, 0, 0, PicWidth, (PicHeight / 4), DC, 0, (PicHeight / 4), vbSrcCopy
    End If
    
    If Trim(mCaption) <> "" Then
        Rt.Left = 0
        Rt.Top = 0
        Rt.Bottom = ScaleHeight
        Rt.Right = ScaleWidth
        OldColor = UserControl.ForeColor
        If Not UserControl.Enabled Then UserControl.ForeColor = vbGrayText
        DrawText hdc, mCaption, Len(mCaption), Rt, GetAlign(mAlign) Or DT_NOCLIP Or DT_SINGLELINE
        UserControl.ForeColor = OldColor
    End If
    Refresh
    ButtonState = 1
    
End Sub
Private Sub DrawDown()
    
    If ButtonState = 2 Then Exit Sub
    If Not HasPicture Then
        Cls
        UserControl.Line (0, 0)-Step(ScaleWidth - 1, 0), vb3DDKShadow
        UserControl.Line (0, 0)-Step(0, ScaleHeight - 1), vb3DDKShadow
        UserControl.Line (0, ScaleHeight - 1)-Step(ScaleWidth, 0), vb3DHighlight
        UserControl.Line (ScaleWidth - 1, 0)-Step(0, ScaleHeight - 1), vb3DHighlight
    Else
        BitBlt hdc, 0, 0, PicWidth, (PicHeight / 4), DC, 0, (PicHeight / 4) * 2, vbSrcCopy
    End If
    
    If Trim(mCaption) <> "" Then
        Rt.Left = 1
        Rt.Top = 1
        Rt.Bottom = ScaleHeight + 1
        Rt.Right = ScaleWidth + 1
        OldColor = UserControl.ForeColor
        If Not UserControl.Enabled Then UserControl.ForeColor = vbGrayText
        DrawText hdc, mCaption, Len(mCaption), Rt, GetAlign(mAlign) Or DT_NOCLIP Or DT_SINGLELINE
        UserControl.ForeColor = OldColor
    End If
    Refresh
    ButtonState = 2
    
End Sub
Private Function GetAlign(ByVal Alng As Alignment) As Long
    
    Select Case Alng
        Case 0: GetAlign = DT_CENTER Or DT_VCENTER
        Case 1: GetAlign = DT_CENTER Or DT_TOP
        Case 2: GetAlign = DT_CENTER Or DT_BOTTOM
        Case 3: GetAlign = DT_LEFT Or DT_VCENTER
        Case 4: GetAlign = DT_LEFT Or DT_TOP
        Case 5: GetAlign = DT_LEFT Or DT_BOTTOM
        Case 6: GetAlign = DT_RIGHT Or DT_VCENTER
        Case 7: GetAlign = DT_RIGHT Or DT_TOP
        Case 8: GetAlign = DT_RIGHT Or DT_BOTTOM
    End Select

End Function
Private Sub Timer1_Timer()
    
    GetCursorPos Pt
    ScreenToClient UserControl.hwnd, Pt
    MouseOver = Not ((Pt.X < 0) Or (Pt.X > ScaleWidth) Or (Pt.Y < 0) Or (Pt.Y > ScaleHeight))
    If Not PtIn Then
        If HasPicture Then MouseOver = False
    End If
    If MouseOver Then
        If MouseButton = 1 Then
            DrawDown
        Else
            DrawUp
        End If
    Else
        If MouseButton = 1 Then
            DrawUp
        Else
            Timer1.Enabled = False
            DrawMouseOut
        End If
    End If

End Sub
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)

    If mUseSound > 0 Then PlaySound VarPtr(B(0)), App.hInstance, SND_ASYNC Or SND_MEMORY Or SND_NOWAIT
    RaiseEvent Click

End Sub
Sub GetBytes()

    ReDim B(0 To 169) As Byte
    B(0) = &H52: B(1) = &H49: B(2) = &H46: B(3) = &H46: B(4) = &H1A: B(5) = &H7: B(6) = &H0: B(7) = &H0: B(8) = &H57: B(9) = &H41: B(10) = &H56: B(11) = &H45: B(12) = &H66: B(13) = &H6D: B(14) = &H74: B(15) = &H20:
    B(16) = &H12: B(17) = &H0: B(18) = &H0: B(19) = &H0: B(20) = &H1: B(21) = &H0: B(22) = &H2: B(23) = &H0: B(24) = &H40: B(25) = &H1F: B(26) = &H0: B(27) = &H0: B(28) = &H80: B(29) = &H3E: B(30) = &H0: B(31) = &H0:
    B(32) = &H2: B(33) = &H0: B(34) = &H8: B(35) = &H0: B(36) = &H0: B(37) = &H0: B(38) = &H66: B(39) = &H61: B(40) = &H63: B(41) = &H74: B(42) = &H4: B(43) = &H0: B(44) = &H0: B(45) = &H0: B(46) = &HA: B(47) = &H0:
    B(48) = &H0: B(49) = &H0: B(50) = &H64: B(51) = &H61: B(52) = &H74: B(53) = &H61: B(54) = &H14: B(55) = &H0: B(56) = &H0: B(57) = &H0: B(58) = &H83: B(59) = &H83: B(60) = &HA9: B(61) = &HA9: B(62) = &H5A: B(63) = &H5A:
    B(64) = &H0: B(65) = &H0: B(66) = &H0: B(67) = &H0: B(68) = &H51: B(69) = &H51: B(70) = &HCE: B(71) = &HCE: B(72) = &H7D: B(73) = &H7D: B(74) = &H6E: B(75) = &H6E: B(76) = &H5A: B(77) = &H5A: B(78) = &H44: B(79) = &H49:
    B(80) = &H53: B(81) = &H50: B(82) = &HCC: B(83) = &H6: B(84) = &H0: B(85) = &H0: B(86) = &H8: B(87) = &H0: B(88) = &H0: B(89) = &H0: B(90) = &H28: B(91) = &H0: B(92) = &H0: B(93) = &H0: B(94) = &H1F: B(95) = &H0:
    B(96) = &H0: B(97) = &H0: B(98) = &H15: B(99) = &H0: B(100) = &H0: B(101) = &H0: B(102) = &H1: B(103) = &H0: B(104) = &H8: B(105) = &H0: B(106) = &H0: B(107) = &H0: B(108) = &H0: B(109) = &H0: B(110) = &H0: B(111) = &H0:
    B(112) = &H0: B(113) = &H0: B(114) = &H0: B(115) = &H0: B(116) = &H0: B(117) = &H0: B(118) = &H0: B(119) = &H0: B(120) = &H0: B(121) = &H0: B(122) = &H0: B(123) = &H0: B(124) = &H0: B(125) = &H0: B(126) = &H0: B(127) = &H0:
    B(128) = &H0: B(129) = &H0: B(130) = &H0: B(131) = &H0: B(132) = &H0: B(133) = &H0: B(134) = &H0: B(135) = &H0: B(136) = &HBF: B(137) = &H0: B(138) = &H0: B(139) = &HBF: B(140) = &H0: B(141) = &H0: B(142) = &H0: B(143) = &HBF:
    B(144) = &HBF: B(145) = &H0: B(146) = &HBF: B(147) = &H0: B(148) = &H0: B(149) = &H0: B(150) = &HBF: B(151) = &H0: B(152) = &HBF: B(153) = &H0: B(154) = &HBF: B(155) = &HBF: B(156) = &H0: B(157) = &H0: B(158) = &HC0: B(159) = &HC0:
    B(160) = &HC0: B(161) = &H0: B(162) = &HC0: B(163) = &HDC: B(164) = &HC0: B(165) = &H0: B(166) = &HF0: B(167) = &HCA: B(168) = &HA6: B(169) = &H0

End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    PtIn = (GetPixel(DC, X, Y + ((PicHeight / 4) * 3)) = 0)
    If Not HasPicture Then PtIn = True
    If PtIn Then
        If Not Timer1.Enabled Then Timer1.Enabled = True
    End If
    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    mAlign = PropBag.ReadProperty("Align", DT_DEFAULT)
    mCaption = PropBag.ReadProperty("Caption", "Command")
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HE0E0E0)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.AccessKeys = PropBag.ReadProperty("AccessKey", "")
    mUseSound = PropBag.ReadProperty("UseSound", 1)
    mSoundFileName = PropBag.ReadProperty("SoundFileName", "")

    If mUseSound = 1 Then
        GetBytes
    ElseIf mUseSound = 2 Then
        If Trim(mSoundFileName) <> "" Then
            Open mSoundFileName For Binary As #1
            ReDim B(LOF(1))
            Get #1, , B
            Close #1
        Else
            mUseSound = None
        End If
    End If

    If UserControl.Ambient.UserMode Then
        DeleteObject Obj
        DeleteDC DC
        DC = CreateCompatibleDC(0)
        Obj = SelectObject(DC, Picture)
    End If
    
    Rt.Left = 0
    Rt.Top = 0
    Rt.Right = ScaleWidth
    Rt.Bottom = ScaleHeight
    
    If Not HasPicture Then DrawMouseOut
    If Not UserControl.Enabled Then Exit Sub
    
    If Trim(mCaption) <> "" Then
        OldColor = UserControl.ForeColor
        If Not UserControl.Enabled Then UserControl.ForeColor = vbGrayText
        DrawText hdc, mCaption, Len(mCaption), Rt, GetAlign(mAlign) Or DT_NOCLIP Or DT_SINGLELINE
        UserControl.ForeColor = OldColor
        X = InStr(1, mCaption, "&")
        If X > 0 Then UserControl.AccessKeys = Mid(mCaption, X + 1, 1)
    End If

End Sub
Private Sub UserControl_Click()
    
    If Not PtIn And HasPicture Then Exit Sub
    If mUseSound > 0 Then PlaySound VarPtr(B(0)), App.hInstance, SND_ASYNC Or SND_MEMORY Or SND_NOWAIT
    RaiseEvent Click
    
End Sub
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    
    BackColor = UserControl.BackColor

End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    Cls
    If ButtonState = 0 Then
        DrawMouseOut
    ElseIf ButtonState = 1 Then
        DrawUp
    Else
        DrawDown
    End If
    
End Property
Private Sub UserControl_DblClick()
    
    RaiseEvent DblClick

End Sub
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    
    Set Font = UserControl.Font

End Property
Public Property Set Font(ByVal New_Font As Font)
    
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Cls
    If ButtonState = 0 Then
        DrawMouseOut
    ElseIf ButtonState = 1 Then
        DrawUp
    Else
        DrawDown
    End If
    
End Property
Public Property Get ForeColor() As OLE_COLOR
    
    ForeColor = UserControl.ForeColor
    
End Property
Public Property Get Align() As Alignment
    
    Align = mAlign
    
End Property
Public Property Get AccessKey() As String
    
    AccessKey = UserControl.AccessKeys
    
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    Cls
    If ButtonState = 0 Then
        DrawMouseOut
    ElseIf ButtonState = 1 Then
        DrawUp
    Else
        DrawDown
    End If
    
End Property
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    
    hwnd = UserControl.hwnd

End Property
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    RaiseEvent KeyDown(KeyCode, Shift)

End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    
    RaiseEvent KeyPress(KeyAscii)

End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    
    RaiseEvent KeyUp(KeyCode, Shift)
    
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    MouseButton = Button
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
End Sub
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    
    Set MouseIcon = UserControl.MouseIcon

End Property
Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"

End Property
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    
    MousePointer = UserControl.MousePointer

End Property
Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"

End Property
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If PtIn And MouseButton = 1 Then
        DrawUp
    End If
    MouseButton = 0
    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    
    Set Picture = UserControl.Picture

End Property
Public Property Set Picture(ByVal New_Picture As Picture)
    
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
    UserControl.BackStyle = 1
    If UserControl.Picture <> 0 Then
        PicHeight = ScaleY(New_Picture.Height, 8, 3)
        PicWidth = ScaleX(New_Picture.Width, 8, 3)
        Height = PicHeight / 4
        Width = ScaleX(PicWidth, 3, 1)
        PaintPicture New_Picture, 0, 0, PicWidth, PicHeight / 4, 0, (PicHeight / 4) * 3, PicWidth, PicHeight / 4, vbSrcCopy
        UserControl.BackStyle = 0
        UserControl.MaskPicture = Image
        UserControl.MaskColor = &HFFFFFF
        PaintPicture New_Picture, 0, 0, PicWidth, PicHeight / 4, 0, 0, PicWidth, PicHeight / 4, vbSrcCopy
        If Trim(mCaption) <> "" Then
            Rt.Left = 0
            Rt.Top = 0
            Rt.Bottom = ScaleHeight
            Rt.Right = ScaleWidth
            OldColor = UserControl.ForeColor
            If Not UserControl.Enabled Then UserControl.ForeColor = vbGrayText
            DrawText hdc, mCaption, Len(mCaption), Rt, GetAlign(mAlign) Or DT_NOCLIP Or DT_SINGLELINE
            UserControl.ForeColor = OldColor
        End If
        HasPicture = True
    Else
        HasPicture = False
    End If
    
End Property
Private Sub UserControl_InitProperties()
    
    Set UserControl.Font = Ambient.Font

End Sub
Private Sub UserControl_Resize()
    
    If UserControl.Picture <> 0 Then
        Height = ScaleY(PicHeight, 3, 1) / 4
        Width = ScaleX(PicWidth, 3, 1)
    Else
        Cls
        DrawMouseOut
    End If

End Sub
Private Sub UserControl_Terminate()
    
    DoEvents
    DeleteObject Obj
    DeleteDC DC
    
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", mCaption, "Command")
    Call PropBag.WriteProperty("UseSound", mUseSound, 1)
    Call PropBag.WriteProperty("SoundFileName", mSoundFileName, "")
    Call PropBag.WriteProperty("Align", mAlign, DT_DEFAULT)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HE0E0E0)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("AccessKey", UserControl.AccessKeys, "")
    
End Sub
