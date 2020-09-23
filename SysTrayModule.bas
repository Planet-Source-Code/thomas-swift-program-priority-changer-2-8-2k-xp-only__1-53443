Attribute VB_Name = "SysTrayModule"
'      Need to add to form using this:
'      PS: Also remember to "RemoveFromTray" when your form unloads
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'
' Dim Message As Long
'   On Error Resume Next
'    Message = x / Screen.TwipsPerPixelX
'    Select Case Message
'        'Your Choice:
'        Case WM_RBUTTONUP
'            PopupMenu [Menu]
'        Case WM_RBUTTONDOWN
'            PopupMenu [Menu]
'    End Select
'End Sub
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIM_MODIFY = &H1

Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4

Private Const WM_MOUSEMOVE = &H200

'Left-click constants.
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up

'Right-click constants.
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up

Dim TrayIcon As NOTIFYICONDATA

Public Sub AddToTray(frm As Form, ToolTip As String, Icon)
'On Error Resume Next
TrayIcon.cbSize = Len(TrayIcon)
TrayIcon.hWnd = frm.hWnd
TrayIcon.szTip = ToolTip & vbNullChar
TrayIcon.hIcon = Icon
TrayIcon.uID = vbNull
TrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
TrayIcon.uCallbackMessage = WM_MOUSEMOVE
Shell_NotifyIcon NIM_ADD, TrayIcon
End Sub
Public Sub RemoveFromTray()
Shell_NotifyIcon NIM_DELETE, TrayIcon
End Sub
Public Function GetY()
    Dim Point As POINTAPI, RetVal As Long
    RetVal = GetCursorPos(Point)
    GetY = Point.Y
End Function
