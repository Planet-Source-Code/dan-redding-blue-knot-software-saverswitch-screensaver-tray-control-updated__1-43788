Attribute VB_Name = "modSystray"
Option Explicit

'Structure for holding all the data for a tray icon
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'constants to specify action for Shell_NotifyIconA
Const NIM_ADD = 0
Const NIM_MODIFY = 1
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

'Message to send back to the window '
'(will make VB route callback to Form_MouseMove)
Const WM_MOUSEMOVE = &H200
   
'Work w/ Tray icon
Private Declare Function Shell_NotifyIconA Lib "shell32" _
    (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer
'Brings a window to the front/gives focus
Public Declare Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As Long) As Long

'create tray icon with callback to MouseMove of object specified by lhWnd
Public Sub SetTrayIcon(strMessage As String, lhWnd As Long, lIcon As Long)
Dim nID As NOTIFYICONDATA
    nID = setNOTIFYICONDATA(lhWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, lIcon, strMessage)
    Shell_NotifyIconA NIM_ADD, nID
End Sub

'Remove the tray icon associated with specified hWnd
Public Sub RemoveTrayIcon(lhWnd As Long)
On Error Resume Next
Dim i As Integer, nID As NOTIFYICONDATA, S As String
    nID = setNOTIFYICONDATA(lhWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, 0&, vbNullString)
    Shell_NotifyIconA NIM_DELETE, nID
End Sub

'Use to change message or icon without removing from tray
Public Sub ModifyTrayIcon(strMessage As String, lhWnd As Long, lIcon As Long)
Dim nID As NOTIFYICONDATA
    nID = setNOTIFYICONDATA(lhWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, lIcon, strMessage)
    Shell_NotifyIconA NIM_MODIFY, nID
End Sub

'Builds the basic tray icon structure
Private Function setNOTIFYICONDATA(hwnd As Long, ID As Long, Flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
Dim nidTemp As NOTIFYICONDATA
    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hwnd = hwnd
    nidTemp.uID = ID
    nidTemp.uFlags = Flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)
    setNOTIFYICONDATA = nidTemp
End Function


