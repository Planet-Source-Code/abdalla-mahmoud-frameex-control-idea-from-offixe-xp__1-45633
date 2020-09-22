VERSION 5.00
Begin VB.UserControl FrameEx 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   PropertyPages   =   "FrameEx.ctx":0000
   ScaleHeight     =   1980
   ScaleWidth      =   4800
   ToolboxBitmap   =   "FrameEx.ctx":0040
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   960
      Top             =   720
   End
End
Attribute VB_Name = "FrameEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'// Author  : Abdalla Mahmoud
'// Age     : 15
'// Country : Egypt
'// City    : Mansoura
'// E-Mails : la3toot@hotmail.com
'             la3toot@yahoo.com
Option Explicit
Implements ISubclass
'API Declarion
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
'API Types
Private Type POINTAPI
        x As Long
        y As Long
End Type
'API Constans
Private Const WM_SETFOCUS  As Long = &H7
Private Const WM_MOUSEMOVE As Long = &H200
'CONTRL VARIABLES
Private m_Caption        As String
Private m_BodyColor      As OLE_COLOR
Private m_ForeColor      As OLE_COLOR
Private m_HighlightColor As OLE_COLOR
Private m_Hwnd           As Long
Private m_Enabled        As Boolean
Private m_CollControls   As New Collection
Private m_TitleHeight    As Single
Private m_Moved          As Boolean
Private m_Picture        As IPictureDisp
Private m_AccessKeyPos   As Long
'CONTROL EVENTS
Public Event Click()
Public Event DblClick()
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Public Event MouseLeave(ByVal ControlActive As Boolean)
Attribute MouseLeave.VB_MemberFlags = "200"
Public Event ControlFocus(ByRef FocusControl As Object)

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer.EMsgResponse)

End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer.EMsgResponse
    ISubclass_MsgResponse = emrPostProcess
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Ambient.UserMode = False Then Exit Function
    If iMsg = WM_SETFOCUS Then
        RaiseEvent ControlFocus(m_CollControls("h" & hWnd))
        tmrMain_Timer
    End If
    If m_Moved Then Exit Function
    DrawFrameEx
    m_Moved = True
End Function

Private Sub tmrMain_Timer()
    Dim HW As Long
    Dim HW2 As Long
    Dim PO As POINTAPI
    Dim PA As Long
    Call GetCursorPos(PO)
    HW = WindowFromPoint(PO.x, PO.y)
    If HW <> m_Hwnd Then
        If GetParent(HW) <> m_Hwnd Then
            If GetParent(GetFocus) <> m_Hwnd Then
                DrawFrame
                tmrMain.Enabled = False
                RaiseEvent MouseLeave(False)
                m_Moved = False
            Else
                RaiseEvent MouseLeave(True)
            End If
        Else
            RaiseEvent MouseLeave(False)
        End If
    End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    m_Hwnd = UserControl.hWnd
    DrawFrame
End Sub

Private Sub UserControl_InitProperties()
    BodyColor = vbButtonShadow
    BackColor = vbButtonFace
    ForeColor = vbWhite
    Set Font = Ambient.Font
    Caption = Ambient.DisplayName
    HighlightColor = vbBlue
    MousePointer = 0
    DrawFrame
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
    If m_Moved Then Exit Sub
    DrawFrameEx
    m_Moved = True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
    On Error Resume Next
    If y <= m_TitleHeight Then m_CollControls(m_CollControls.Count).SetFocus
    Err.Clear
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BodyColor = PropBag.ReadProperty("BodyColor", vbButtonShadow)
    BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    ForeColor = PropBag.ReadProperty("ForeColor", vbWhite)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    HighlightColor = PropBag.ReadProperty("HighlightColor", vbBlue)
    Enabled = PropBag.ReadProperty("Enabled", True)
    MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

Private Sub UserControl_Resize()
    DrawFrame
End Sub

Public Property Get BodyColor() As OLE_COLOR
Attribute BodyColor.VB_Description = "Return\\Sets the body color of the frame ."
Attribute BodyColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BodyColor = m_BodyColor
End Property

Public Property Let BodyColor(ByVal vNewValue As OLE_COLOR)
    m_BodyColor = vNewValue
    DrawFrame
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Return\\Sets the backcolor of the frame"
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
    UserControl.BackColor = vNewValue
    DrawFrame
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Return\\Sets the forecolor of the caption ."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
    m_ForeColor = vNewValue
    DrawFrame
    PropertyChanged "ForeColor"
End Property

Public Property Get Font() As IFontDisp
Attribute Font.VB_Description = "Return\\Sets the display font ."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Let Font(ByVal vNewValue As IFontDisp)
    Set UserControl.Font = vNewValue
    m_TitleHeight = TextHeight("0")
    DrawFrame
    PropertyChanged "Font"
End Property

Public Property Set Font(ByVal vNewValue As IFontDisp)
    Set UserControl.Font = vNewValue
    m_TitleHeight = TextHeight("0")
    DrawFrame
    PropertyChanged "Font"
End Property

Private Sub UserControl_Show()
    If Ambient.UserMode = False Then Exit Sub
    On Error Resume Next
    Dim CTL As Control
    Dim tHwnd As Long
    Dim I As Long
    For Each CTL In UserControl.ContainedControls
        If Not CTL Is tmrMain Then
            tHwnd = 0
            tHwnd = CTL.hWnd
            If tHwnd <> 0 Then
                AttachMessage Me, tHwnd, WM_SETFOCUS
                AttachMessage Me, tHwnd, WM_MOUSEMOVE
                m_CollControls.Add CTL, "h" & tHwnd
            End If
        End If
    Next
    Err.Clear
    Set CTL = Nothing
    If GetParent(GetFocus) = m_Hwnd Then
        tmrMain.Enabled = True
    End If
End Sub

Private Sub UserControl_Terminate()
    Set m_CollControls = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BodyColor", BodyColor)
    Call PropBag.WriteProperty("BackColor", BackColor)
    Call PropBag.WriteProperty("ForeColor", ForeColor)
    Call PropBag.WriteProperty("Font", Font)
    Call PropBag.WriteProperty("Caption", Caption)
    Call PropBag.WriteProperty("HighlightColor", HighlightColor)
    Call PropBag.WriteProperty("Enabled", Enabled)
    Call PropBag.WriteProperty("MousePointer", MousePointer)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon)
    Call PropBag.WriteProperty("Picture", Picture)
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Return\\Sets the caption of the frame ."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal vNewValue As String)
    m_Caption = vNewValue
    DrawFrame
    PropertyChanged "Caption"
End Property

Private Sub DrawFrame()
    Cls
    If Not m_Picture Is Nothing Then _
    UserControl.PaintPicture m_Picture, 0, m_TitleHeight + 100, ScaleWidth, ScaleHeight - m_TitleHeight - 100
    Line (0, 0)-(ScaleWidth, 100 + m_TitleHeight), m_BodyColor, BF
    UserControl.CurrentX = (ScaleWidth - TextWidth(Caption)) / 2
    UserControl.CurrentY = 50
    UserControl.ForeColor = m_ForeColor
    Print m_Caption
    Line (0, 0)-(0, ScaleHeight), m_BodyColor
    Line (0, ScaleHeight - 15)-(ScaleWidth, ScaleHeight - 15), m_BodyColor
    Line (ScaleWidth - 15, 0)-(ScaleWidth - 15, ScaleHeight), m_BodyColor
End Sub

Private Sub DrawFrameEx()
    Cls
    If Not m_Picture Is Nothing Then _
    UserControl.PaintPicture m_Picture, 0, m_TitleHeight + 100, ScaleWidth, ScaleHeight - m_TitleHeight - 100
    Line (0, 0)-(ScaleWidth, 100 + m_TitleHeight), m_HighlightColor, BF
    UserControl.CurrentX = (ScaleWidth - TextWidth(Caption)) / 2
    UserControl.CurrentY = 50
    UserControl.ForeColor = m_ForeColor
    Print m_Caption
    Line (0, 0)-(0, ScaleHeight), m_HighlightColor
    Line (0, ScaleHeight - 15)-(ScaleWidth, ScaleHeight - 15), m_HighlightColor
    Line (ScaleWidth - 15, 0)-(ScaleWidth - 15, ScaleHeight), m_HighlightColor
    tmrMain.Enabled = True
End Sub

Public Property Get HighlightColor() As OLE_COLOR
Attribute HighlightColor.VB_Description = "Return\\Sets the color displayed when mouse on the frame or one of the children controls ."
    HighlightColor = m_HighlightColor
End Property

Public Property Let HighlightColor(ByVal vNewValue As OLE_COLOR)
    m_HighlightColor = vNewValue
    PropertyChanged "HighlightColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Return\\Sets the enabled of the frame ."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    UserControl.Enabled = vNewValue
    If vNewValue = False Then DrawFrame
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal vNewValue As MousePointerConstants)
    On Error Resume Next
    UserControl.MousePointer = vNewValue
    Err.Clear
End Property

Public Property Get MouseIcon() As IPictureDisp
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Let MouseIcon(ByVal vNewValue As IPictureDisp)
    On Error Resume Next
    Set UserControl.MouseIcon = vNewValue
    Err.Clear
End Property

Public Property Set MouseIcon(ByVal vNewValue As IPictureDisp)
    On Error Resume Next
    Set UserControl.MouseIcon = vNewValue
    Err.Clear
End Property

Public Property Get Picture() As IPictureDisp
    Set Picture = m_Picture
End Property

Public Property Let Picture(ByVal vNewValue As IPictureDisp)
    Set m_Picture = vNewValue
    If m_Moved Then DrawFrameEx Else DrawFrame
    PropertyChanged "Picture"
End Property

Public Property Set Picture(ByVal vNewValue As IPictureDisp)
    Set m_Picture = vNewValue
    If m_Moved Then DrawFrameEx Else DrawFrame
    PropertyChanged "Picture"
End Property
