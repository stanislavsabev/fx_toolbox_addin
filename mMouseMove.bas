Attribute VB_Name = "mMouseMove"
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MOUSEHOOKSTRUCT
    pt As POINTAPI
    hwnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type

Private Declare PtrSafe Function FindWindow Lib "User32" _
    Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr
        

Private Declare PtrSafe Function GetWindowLong Lib "user32.dll" _
    Alias "GetWindowLongA" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpWindowName As LongPtr) As LongPtr
        

Private Declare PtrSafe Function SetWindowsHookEx Lib "User32" _
    Alias "SetWindowsHookExA" ( _
        ByVal idHook As LongPtr, _
        ByVal lpfn As LongPtr, _
        ByVal hmod As LongPtr, _
        ByVal dwThreadId As LongPtr) As LongPtr


Private Declare PtrSafe Function CallNextHookEx Lib "User32" ( _
        ByVal hHook As LongPtr, _
        ByVal nCode As LongPtr, _
        ByVal wParam As LongPtr, _
        lParam As Any) As Long
        
Private Declare PtrSafe Function UnhookWindowHookEx Lib "User32" ( _
        ByVal hHook As LongPtr) As Long
    
Private Declare PtrSafe Function PostMessage Lib "user32.dll" _
    Alias "PostMessageA" ( _
        ByVal hwnd As LongPtr, _
        ByVal wMsg As LongPtr, _
        ByVal wParam As LongPtr, _
        ByVal lParam As LongPtr) As LongPtr

'
'Private Declare PtrSafe Function WindowFromPoint Lib "user32" ( _
'        ByVal xPoint As LongPtr, _
'        ByVal yPoint As LongPtr) As LongPtr
'Private Declare PtrSafe Function WindowFromPoint Lib "user32" _
'    (ByVal Point As LongLong) As LongPtr

#If Win32 Then
    Declare PtrSafe Function WindowFromPointXY Lib "User32" _
    (ByVal xPoint As Long, _
    ByVal yPoint As Long) As LongPtr
#Else
    Declare PtrSafe Function WindowFromPointYX Lib "User" _
    (ByVal yPoint As Integer, ByVal _
    xPoint As Integer) As Integer
#End If
        
Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
              
Private Declare PtrSafe Function GetCursorPos Lib "user32.dll" ( _
        ByRef lpPoint As POINTAPI) As Long
        
        
'Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long
Private Declare PtrSafe Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As LongPtr) As Long
Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

Private Const WH_MOUSE_LL As Long = 14
Private Const WM_MOUSEWHEEL = &H20A
Private Const HC_ACTION = 0
Private Const GWL_HINSTANCE = (-6)

Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const VK_UP = &H26
Private Const VK_DOWN = &H28
Private Const WM_LBUTTONDOWN = &H201

Private mLngMouseHook As LongPtr
Private mListBoxHwnd As LongPtr
Private mbHook As Boolean


Sub HookListBoxScroll()
    Dim lngAppInst As LongPtr
    Dim hwndUnderCursor As LongPtr
    Dim tPT As POINTAPI
    Dim ptLL As LongLong
    
    GetCursorPos tPT 'WindowFromPoint(tPT.x, tPT.Y)
    ptLL = PointToLongLong(tPT)
    hwndUnderCursor = VBWindowFromPoint(tPT.x, tPT.y)
    If mListBoxHwnd <> hwndUnderCursor Then
        UnhookListBoxScroll
        mListBoxHwnd = hwndUnderCursor
        lngAppInst = GetWindowLong(mListBoxHwnd, GWL_HINSTANCE)
        PostMessage mListBoxHwnd, WM_LBUTTONDOWN, 0&, 0&
        
        If Not mbHook Then
            mLngMouseHook = SetWindowsHookEx( _
                WH_MOUSE_LL, AddressOf MouseProc, lngAppInst, 0)
            mbHook = (mLngMouseHook <> 0)
        End If
    End If
End Sub

Sub UnhookListBoxScroll()
    If mbHook Then
        UnhookWindowHookEx mLngMouseHook
        mLngMouseHook = 0
        mListBoxHwnd = 0
        mbHook = False
    End If
End Sub

Private Function MouseProc( _
    ByVal nCode As LongPtr, ByVal wParam As LongPtr, _
    ByRef lParam As MOUSEHOOKSTRUCT) As LongPtr
    Dim ptLL As LongLong
    
    On Error GoTo errH
    If (nCode = HC_ACTION) Then
          'WindowFromPoint(lParam.pt.x, lParam.pt.Y)
        ptLL = PointToLongLong(lParam.pt)
        If VBWindowFromPoint(lParam.pt.x, lParam.pt.y) = mListBoxHwnd Then
            If wParam = WM_MOUSEWHEEL Then
                MouseProc = True
                If lParam.hwnd > 0 Then
                    PostMessage mListBoxHwnd, WM_KEYDOWN, VK_UP, 0
                Else
                    PostMessage mListBoxHwnd, WM_KEYDOWN, VK_DOWN, 0
                End If
                PostMessage mListBoxHwnd, WM_KEYUP, VK_UP, 0
            End If
        Else
            UnhookListBoxScroll
        End If
    End If
        
    MouseProc = CallNextHookEx( _
        mLngMouseHook, nCode, wParam, ByVal lParam)
Exit Function
errH:
    UnhookListBoxScroll
End Function

Function VBWindowFromPoint(ByVal x As Long, ByVal y As Long) As LongPtr
#If Win32 Then
VBWindowFromPoint = WindowFromPointXY(x, y)
#Else
VBWindowFromPoint = WindowFromPointYX(y, x)
#End If
End Function


Function PointToLongLong(Point As POINTAPI) As LongLong
    Dim LL As LongLong
    Dim cbLongLong As LongPtr
      
    cbLongLong = LenB(LL)
      
    ' make sure the contents will fit
    If LenB(Point) = cbLongLong Then
        CopyMemory LL, Point, cbLongLong
    End If
      
    PointToLongLong = LL
End Function
