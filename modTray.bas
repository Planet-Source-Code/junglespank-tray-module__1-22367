Attribute VB_Name = "modTray"
Option Explicit
'-------------------------------------------------------------------------------------
'WIN32 API Constants
'-------------------------------------------------------------------------------------
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GWL_WNDPROC = (-4)
Public Const IDANI_OPEN = &H1
Public Const IDANI_CLOSE = &H2
Public Const IDANI_CAPTION = &H3
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_MOUSEMOVE = &H200
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_USER = &H400

'----------------------------------------------------------------------------
'Application Defined Constants
'----------------------------------------------------------------------------
Public Const OPT_COPY = 0
Public Const OPT_MOVE = 1
Public Const OPT_DELETE = 2
Public Const OPT_RENAME = 3
Public Const WM_CALLBACK_MSG = WM_USER Or &HF

'----------------------------------------------------------------------------
'WIN32 Type Declares
'----------------------------------------------------------------------------
Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'-------------------------------------------------------------------------------------
'Application defined enumerations
'-------------------------------------------------------------------------------------
Public Enum ZoomTypes
    ZOOM_FROM_TRAY
    ZOOM_TO_TRAY
End Enum

'-------------------------------------------------------------------------------------
'WIN32 DLL Declares
'-------------------------------------------------------------------------------------
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
                 ByVal lpWindowName As String) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
                 ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function DrawAnimatedRects Lib "user32" (ByVal hwnd As Long, ByVal idAni As Long, lprcFrom As RECT, _
                 lprcTo As RECT) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                 ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                 ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
                 ByVal wParam As Long, lParam As Any) As Long
Declare Function Shell_NotifyIconA Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
                                                                                                         
'-------------------------------------------------------------------------------------
'Public Variables
'-------------------------------------------------------------------------------------
Public lngPrevWndProc As Long       'Original WNDPROC address.
Public lngWndID As Long             'Our unique icon identifier.
Public lngHwnd As Long              'The hwnd of program.
Public nidTray As NOTIFYICONDATA    'Global Icon Info Structure.
Public strToolTip As String         'The string we use in the tip text fro the tray.
Public intOptionChoice As Integer   'Option button tracker.

Public rctFrom As RECT, rctTo As RECT
Public lngTrayHand As Long
Public lngStartMenuHand As Long, lngChildHand As Long
Public strClass As String * 255
Public lngClassNameLen As Long
Public lngRetVal As Long

Public Function WndProcMain(ByVal hwnd As Long, ByVal message As Long, ByVal wParam As Long, _
                                        ByVal lParam As Long) As Long
'-------------------------------------------------------------------------------------
'Window Proc for System Tray.  Used for sub-classing.
'-------------------------------------------------------------------------------------
If lngWndID = wParam Then
    Select Case lParam              'Button Responses
        Case WM_LBUTTONDBLCLK
            'Left Double-Click Procedure Here
        Case WM_LBUTTONDOWN
            'Left Mouse Down Procedure Here
        Case WM_LBUTTONUP
            'Left Mouse Up Procedure Here
            ClearTray
            RestoreWindow
        Case WM_RBUTTONDBLCLK
            'Right Double-Click Procedure Here
        Case WM_RBUTTONDOWN
            'Right Mouse Down Procedure Here
        Case WM_RBUTTONUP
            'Right Mouse Up Procudure Here
            'Show our popup menu in the tray.
            frmMain.PopupMenu frmMain.mnuPopup
    End Select
End If

'Call the original WNDPROC (with the passed in values in tact) to handle any messages we ignored.
WndProcMain = CallWindowProc(lngPrevWndProc, hwnd, message, wParam, lParam)

End Function


Public Function ZoomForm(zoomToWhere As ZoomTypes, hwnd As Long) As Boolean
'-------------------------------------------------------------------------------------
'This function 'zooms' a window.
'-------------------------------------------------------------------------------------

'Select the type of zoom to do.
Select Case zoomToWhere

    'Zoom the window into the tray.
    Case ZOOM_FROM_TRAY
        'Get the handle to the start menu.
        lngStartMenuHand = FindWindow("Shell_TrayWnd", vbNullString)

        'Get the handle to the first child window of the start menu.
        lngChildHand = GetWindow(lngStartMenuHand, GW_CHILD)

        'Loop through all siblings until we find the 'System Tray' (A.K.A. --> TrayNotifyWnd)
        Do
                lngClassNameLen = GetClassName(lngChildHand, strClass, Len(strClass))
    
            'If it is the tray then store the handle.
            If InStr(1, strClass, "TrayNotifyWnd") Then
                lngTrayHand = lngChildHand
                Exit Do
            End If
    
            'If we didn't find it, go to the next sibling.
            lngChildHand = GetWindow(lngChildHand, GW_HWNDNEXT)

        Loop

        'Get the RECT of  our form.
        lngRetVal = GetWindowRect(hwnd, rctFrom)
        
        'Get the RECT of the Tray.
        lngRetVal = GetWindowRect(lngTrayHand, rctTo)
        
        'Zoom from the tray to where our form is.
        lngRetVal = DrawAnimatedRects(frmMain.hwnd, IDANI_CLOSE Or IDANI_CAPTION, rctTo, rctFrom)

    Case ZOOM_TO_TRAY
        'Get the handle to the start menu.
        lngStartMenuHand = FindWindow("Shell_TrayWnd", vbNullString)
        'Get the handle to the first child window of the start menu.
        lngChildHand = GetWindow(lngStartMenuHand, GW_CHILD)
        'Loop through all siblings until we find the 'System Tray' (A.K.A. --> TrayNotifyWnd)
        Do
            lngClassNameLen = GetClassName(lngChildHand, strClass, Len(strClass))
    
            'If it is the tray then store the handle.
            If InStr(1, strClass, "TrayNotifyWnd") Then
                lngTrayHand = lngChildHand
                Exit Do
            End If
    
            'If we didn't find it, go to the next sibling.
            lngChildHand = GetWindow(lngChildHand, GW_HWNDNEXT)

        Loop
        'Get the RECT of  our form.
        lngRetVal = GetWindowRect(hwnd, rctFrom)
        
        'Get the RECT of the Tray.
        lngRetVal = GetWindowRect(lngTrayHand, rctTo)
        
        'Zoom from where our form is to the tray .
        lngRetVal = DrawAnimatedRects(frmMain.hwnd, IDANI_OPEN Or IDANI_CAPTION, rctFrom, rctTo)
            
End Select

End Function

Public Sub SendToTray()
    lngHwnd = frmMain.hwnd                  'Get the hwnd from of the form.
    lngWndID = App.hInstance                'ID used for callback.
    strToolTip = frmMain.Caption & Chr(0)   'Initialize the icon tip.
    lngPrevWndProc = SetWindowLong(lngHwnd, GWL_WNDPROC, AddressOf WndProcMain)        'Subclass the window.
    
    'Set Tray Icon Information
    SetTrayInfo
    'Zoom to Tray
    ZoomForm ZOOM_TO_TRAY, lngHwnd
    'Call windows to add the icon to the tray.
    lngRetVal = Shell_NotifyIconA(NIM_ADD, nidTray)

    'Hide our form.
    frmMain.Hide
End Sub

Public Sub ClearTray()
    'If the user clicked the tray icon then release the hook on the window.
    'In other words un-subclass it by returning it's original WNDPROC address
    'that we captured in the original call to SetWindowLong.
    SetWindowLong lngHwnd, GWL_WNDPROC, lngPrevWndProc
    'Delete our icon from the tray.
    Shell_NotifyIconA NIM_DELETE, nidTray
End Sub

Public Sub RestoreWindow()
    'Zoom to show form
    ZoomForm ZOOM_FROM_TRAY, frmMain.hwnd
    'frmMain.WindowState = vbNormal
    frmMain.Show
    frmMain.SetFocus
End Sub

Public Sub SetTrayInfo()
    'Set all Tray Icon info.
    nidTray.hwnd = lngHwnd
    nidTray.cbSize = Len(nidTray)
    nidTray.hIcon = frmMain.Icon
    nidTray.szTip = strToolTip
    nidTray.uCallbackMessage = WM_CALLBACK_MSG
    nidTray.uID = lngWndID
    nidTray.uFlags = NIF_MESSAGE Or NIF_TIP Or NIF_ICON
End Sub
