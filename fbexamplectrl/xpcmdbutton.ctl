VERSION 5.00
Begin VB.UserControl xpcmdbutton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1815
   DefaultCancel   =   -1  'True
   FillStyle       =   0  'Solid
   ScaleHeight     =   89
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   121
   ToolboxBitmap   =   "xpcmdbutton.ctx":0000
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1080
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   600
      Top             =   360
   End
   Begin VB.PictureBox pc 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   1440
      Picture         =   "xpcmdbutton.ctx":0312
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox pc 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   1080
      Picture         =   "xpcmdbutton.ctx":07EC
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox pc 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   720
      Picture         =   "xpcmdbutton.ctx":0CC6
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox pc 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   360
      Picture         =   "xpcmdbutton.ctx":11A0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox pc 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   0
      Picture         =   "xpcmdbutton.ctx":167A
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   360
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "xpcmdbutton"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   945
   End
End
Attribute VB_Name = "xpcmdbutton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'/----------------------------------------------------\
'/Description: Creation of Windows XP windows and     \
'/             controls in Visual Basic               \
'/Created by: Teh Ming Han (teh_minghan@hotmail.com)  \
'/                                                    \
'/If you use this code in your program please give me \
'/credit and e-mail me (teh_minghan@hotmail.com) and  \
'/tell me about your program.                         \
'/------Hope you find it useful!---------2001---------\
'/                                                    \
'/Modified by Peter Schellenbach in 2010 for use in   \
'/FireBreath ActiveX control wrapper example:         \
'/ - removed dependency on PictureClip control        \
'/ - fixed access key handling                        \
'/ - added support for DisplayAsDefault ambient       \
'/ - added focus rectangle & tracking                 \
'/                                                    \
'/Note: this modified control is not intended for     \
'/production use. Its only purpose is to illustrate   \
'/how to wrap an ActiveX control in a FireBreath      \
'/plug-in.                                            \
'/                                                    \
'/----------------------------------------------------\

Option Explicit

Private Type POINT_API
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINT_API) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lprc As RECT) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Enum ThemeConstants
    ThemeDefault = 0
    ThemeHomestead = 1
    ThemeMetalic = 2
End Enum

Private Const IDBMP = 100 'base resource ID

Private m_MouseIn As Boolean
Private m_Pressed As Boolean
Private m_IsDefault As Boolean
Private m_Focused As Boolean
Private m_AccessKey As Integer
Private m_Theme As ThemeConstants
Private m_Font As Font

Private Const m_Def_Theme = ThemeDefault

Event Click()
Attribute Click.VB_UserMemId = -600
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_UserMemId = -604
Event MouseEnter()
Event MouseLeave()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -607

Private Sub lbl_Change()
    UserControl_Resize
End Sub

Private Sub lbl_Click()
    UserControl_Click
End Sub

Private Sub lbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lbl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Timer1_Timer()
    Dim pnt As POINT_API
    GetCursorPos pnt
    ScreenToClient UserControl.hWnd, pnt
    If pnt.X < UserControl.ScaleLeft Or _
       pnt.Y < UserControl.ScaleTop Or _
       pnt.X > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or _
       pnt.Y > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
        Timer1.Enabled = False
        m_MouseIn = False
        RaiseEvent MouseLeave
        statevalue_pic
    End If
End Sub

Private Sub Timer2_Timer()
    Dim MyFocus As Boolean
    MyFocus = (GetFocus() = UserControl.hWnd)
    If MyFocus <> m_Focused Then
        m_Focused = MyFocus
        statevalue_pic
    End If
End Sub

Private Sub Timer3_Timer()
   Timer3.Enabled = False
   Timer3.Interval = 100 'restore default value
   m_Pressed = False
   statevalue_pic
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        RaiseEvent Click
        Timer3.Enabled = True
        m_Pressed = True
        statevalue_pic
    End If
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
   On Error Resume Next
   If StrComp(PropertyName, "DisplayAsDefault", vbTextCompare) = 0 Then
      m_IsDefault = UserControl.Ambient.DisplayAsDefault
      statevalue_pic
   End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_GotFocus()
   m_Focused = True
   Timer2.Enabled = True
   statevalue_pic
End Sub

Private Sub UserControl_Initialize()
    statevalue_pic
End Sub

Private Sub UserControl_InitProperties()
    On Error Resume Next
    m_IsDefault = UserControl.Ambient.DisplayAsDefault
    Enabled = True
    Caption = UserControl.Ambient.DisplayName
    Set Font = UserControl.Ambient.Font
    Theme = m_Def_Theme
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then
        m_Pressed = True
        statevalue_pic
    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    If KeyAscii = vbKeySpace Or KeyAscii = vbKeyReturn Then
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim DoAccessKey As Boolean
    RaiseEvent KeyUp(KeyCode, Shift)
    If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then
        m_Pressed = False
        statevalue_pic
    ElseIf KeyCode <> 0 And m_AccessKey <> 0 And m_Pressed = False Then
        If KeyCode = m_AccessKey And Shift = vbAltMask Then
            'released the access-key while holding alt-key
            DoAccessKey = True
        ElseIf KeyCode = vbKeyMenu And GetKeyState(m_AccessKey) < 0 Then
            'released alt-key while holding access-key
            DoAccessKey = True
        Else
            DoAccessKey = False
        End If
        If DoAccessKey Then
            RaiseEvent Click
            Timer3.Enabled = True
            m_Pressed = True
            statevalue_pic
        End If
    End If
End Sub

Private Sub UserControl_LostFocus()
   m_Focused = False
   Timer2.Enabled = False
   statevalue_pic
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = vbLeftButton Then
        m_Pressed = True
        statevalue_pic
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= 0 And Y >= 0 And _
        X <= UserControl.ScaleWidth And Y <= UserControl.ScaleHeight Then
        RaiseEvent MouseMove(Button, Shift, X, Y)
        If m_MouseIn = False Then
            m_MouseIn = True
            m_Pressed = (Button = vbLeftButton)
            Timer1.Enabled = True
            RaiseEvent MouseEnter
            statevalue_pic
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Button = vbLeftButton Then
        m_Pressed = False
        statevalue_pic
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    m_IsDefault = UserControl.Ambient.DisplayAsDefault
    Enabled = PropBag.ReadProperty("Enabled", True)
    Caption = PropBag.ReadProperty("Caption")
    Set Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
    Theme = PropBag.ReadProperty("Theme", m_Def_Theme)
End Sub

Private Sub UserControl_Resize()
    statevalue_pic
    lbl.Top = (UserControl.ScaleHeight - lbl.Height) / 2
    lbl.Left = (UserControl.ScaleWidth - lbl.Width) / 2
End Sub

Private Sub UserControl_Show()
    statevalue_pic
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Theme", m_Theme, m_Def_Theme)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Caption", lbl.Caption)
    Call PropBag.WriteProperty("Font", m_Font, UserControl.Ambient.Font)
End Sub

Public Property Get Caption() As String
    Caption = lbl.Caption
End Property

Public Property Let Caption(ByVal vNewCaption As String)
    Dim i As Long
    Dim j As Long
    Dim c As String
    lbl.Caption = vNewCaption
    m_AccessKey = 0
    j = 1
    Do
        i = InStr(j, vNewCaption, "&", vbBinaryCompare)
        If i = 0 Then Exit Do
        c = Mid$(vNewCaption, i + 1, 1)
        If c = "&" Then
            j = i + 2
        ElseIf c = "" Then
            Exit Do
        Else
            If Asc(UCase$(c)) >= 65 And Asc(UCase$(c)) <= 90 Then
                m_AccessKey = Asc(UCase$(c))
            End If
            Exit Do
        End If
    Loop
    PropertyChanged "Caption"
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    statevalue_pic
    If Enabled = True Then lbl.ForeColor = vbBlack Else lbl.ForeColor = RGB(161, 161, 146)
End Property

Public Sub FireClick(Optional ByVal Milliseconds As Long = -1)
   If Milliseconds < 0 Then Milliseconds = 100 'default is 100ms
   If Milliseconds > 30000 Then Milliseconds = 30000 'prevent overflow
   RaiseEvent Click
   If Milliseconds Then
      Timer3.Interval = Milliseconds
      Timer3.Enabled = True
      m_Pressed = True
      statevalue_pic
   End If
End Sub

Public Property Get Font() As Font
    Set Font = m_Font
End Property

Public Property Set Font(ByVal vNewFont As Font)
    Set m_Font = vNewFont
    Set UserControl.Font = vNewFont
    Set lbl.Font = m_Font
    Call UserControl_Resize
    PropertyChanged "Font"
End Property

Public Property Get Theme() As ThemeConstants
   Theme = m_Theme
End Property

Public Property Let Theme(ByVal vNewValue As ThemeConstants)
   Dim i As Integer
   Dim resid As Long
   Dim baseid As Long
   m_Theme = vNewValue
   PropertyChanged "Theme"
   baseid = IDBMP * (m_Theme + 1) '100 = blue, 200 = homestead, 300 = metalic
   For i = 0 To 4
      resid = baseid + i
      Set pc(i).Picture = LoadResPicture(resid, vbResBitmap)
   Next i
   statevalue_pic
End Property

Private Sub statevalue_pic()
    Dim s As Integer
    If UserControl.Enabled Then
        If m_Pressed Then
            s = 1 'pressed
        ElseIf m_MouseIn Then
            s = 3 'hot
        ElseIf m_IsDefault Then
            s = 4 'default
        Else
            s = 0 'normal
        End If
    Else
        s = 2 'disabled
    End If
    make_xpbutton s
End Sub

Private Sub make_xpbutton(z As Integer)
    UserControl.ScaleMode = 3 'Draw in pixels
    Dim brx, bry, bw, bh As Integer
    'Short cuts
    brx = UserControl.ScaleWidth - 3 'right x
    bry = UserControl.ScaleHeight - 3 'right y
    bw = UserControl.ScaleWidth - 6 'border width - corners width
    bh = UserControl.ScaleHeight - 6 'border height - corners height
    'Draws button
    'Goes clockwise first for corners(first four)
    'followed by borders(next four) and center(last step).
    UserControl.PaintPicture pc(z).Picture, 0, 0, 3, 3, 0, 0, 3, 3
    UserControl.PaintPicture pc(z).Picture, brx, 0, 3, 3, 15, 0, 3, 3
    UserControl.PaintPicture pc(z).Picture, brx, bry, 3, 3, 15, 18, 3, 3
    UserControl.PaintPicture pc(z).Picture, 0, bry, 3, 3, 0, 18, 3, 3
    UserControl.PaintPicture pc(z).Picture, 3, 0, bw, 3, 3, 0, 12, 3
    UserControl.PaintPicture pc(z).Picture, brx, 3, 3, bh, 15, 3, 3, 15
    UserControl.PaintPicture pc(z).Picture, 0, 3, 3, bh, 0, 3, 3, 15
    UserControl.PaintPicture pc(z).Picture, 3, bry, bw, 3, 3, 18, 12, 3
    UserControl.PaintPicture pc(z).Picture, 3, 3, bw, bh, 3, 3, 12, 15
    If m_Focused Then
       Dim rc As RECT
       rc.Left = 3
       rc.Top = 3
       rc.Right = brx
       rc.Bottom = bry
       DrawFocusRect UserControl.hDC, rc
    End If
End Sub

