VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolTipFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hwnd As Long
    uid As Long
    RECT As RECT
    hinst As Long
    lpszText As String
    lParam As Long
End Type

Private Const CW_USEDEFAULT = &H80000000
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const WM_USER = &H400
Private Const TTF_CENTERTIP = &H2
Private Const TTF_SUBCLASS = &H10
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_GETTEXTA = (WM_USER + 11)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTS_ALWAYSTIP = &H1
Private Const TTS_BALLOON = &H40
Private Const TOOLTIPS_CLASSA = "tooltips_class32"

Private mBackColor As OLE_COLOR
Private mForeColor As OLE_COLOR

Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    mBackColor = New_BackColor
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    mForeColor = New_ForeColor
End Property

Private Sub Class_Initialize()
SetColorDefaults
End Sub

Private Sub SetColorDefaults()
mBackColor = &H8000000F
mForeColor = &H80FFFF
End Sub
Public Sub AssignToolTip(ByRef hTarget As Object, ByRef sMessage As String)
    Dim TipWindow As Long
    Dim ti As TOOLINFO
    Dim uid As Long
    Dim ToolTipText As String
    Dim RECT As RECT
    uid = 0
    
    TipWindow = CreateWindowEx(0&, TOOLTIPS_CLASSA, "", _
     TTS_ALWAYSTIP Or TTS_BALLOON, 0, 0, _
    0, 0, _
    hTarget.hwnd, 0&, App.hInstance, 0&)

    SetWindowPos TipWindow, HWND_TOPMOST, 0, 0, 0, 0, _
    SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
    GetClientRect hTarget.hwnd, RECT

    'Fill the TOOLINFO structure with info about
    'our tooltip control's tool
    With ti
        .cbSize = Len(ti)
        .uFlags = TTF_CENTERTIP + TTF_SUBCLASS
        .hwnd = hTarget.hwnd
        .hinst = App.hInstance
        .uid = uid
        .lpszText = sMessage
        .RECT = RECT
        .lpszText = sMessage
    End With

    SendMessage TipWindow, TTM_ADDTOOLA, 0, ti
    SendMessage TipWindow, TTM_SETMAXTIPWIDTH, 0, 80
    SendMessage TipWindow, TTM_SETTIPBKCOLOR, mBackColor, 0
    SendMessage TipWindow, TTM_SETTIPTEXTCOLOR, mForeColor, 0
    SetColorDefaults
End Sub


