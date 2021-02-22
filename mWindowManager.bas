Attribute VB_Name = "mWindowManager"
Option Explicit
Private Const ModuleName = "mWindowManager"


Public Sub ZoomAll()
    FxCallStack.Push ModuleName, "ZoomAll"
    On Error GoTo CleanFail
    Dim Perc        As Long

    frmZoomAll.Show
CleanExit:
    Application.ScreenUpdating = True
    FxCallStack.Pop
Exit Sub
CleanFail:
    FxCallStack.ShowExcInfoAndClear Err.Number, Err.Description
    Resume CleanExit
End Sub

Sub ZoomWindow(Perc As Long)
    FxCallStack.Push ModuleName, "ZoomWindow"
    On Error GoTo UnhandledFail
    Dim WsActive    As Worksheet
    Dim Wks         As Worksheet
    Dim Wnd         As Window
    Dim WndActive   As Window

    Set WndActive = ActiveWindow
    Application.ScreenUpdating = False

    For Each Wnd In ActiveWorkbook.Windows
        Wnd.Activate
        Set WsActive = Wnd.ActiveSheet
        For Each Wks In ActiveWorkbook.Worksheets
            If Wks.Visible = xlSheetVisible Then
                Wks.Activate
                Wnd.Zoom = Perc
            End If
        Next Wks
        WsActive.Activate
    Next Wnd

    WndActive.Activate
CleanExit:
    FxCallStack.Pop
Exit Sub
UnhandledFail:
    Err.Raise Err.Number, , Err.Description
End Sub

Private Function ValidateZoomVal(Val As Long) As Boolean
    ValidateZoomVal = (Val >= 10 And Val <= 400)
End Function

Public Sub SplitHorizontal()
    FxCallStack.Push ModuleName, "SplitHorizontal"
    On Error GoTo CleanFail
    Call SplitWindow("Horizontal")
CleanExit:
    FxCallStack.Pop
Exit Sub
CleanFail:
    FxCallStack.ShowExcInfoAndClear Err.Number, Err.Description
    Resume CleanExit
End Sub

Public Sub SplitVertical()
    FxCallStack.Push ModuleName, "SplitVertical"
    On Error GoTo CleanFail
    Call SplitWindow("Vertical")
CleanExit:
    FxCallStack.Pop
Exit Sub
CleanFail:
    FxCallStack.ShowExcInfoAndClear Err.Number, Err.Description
    Resume CleanExit
End Sub

Private Sub SplitWindow(Orientation As String)
    FxCallStack.Push ModuleName, "SplitWindow"
    On Error GoTo UnhandledFail
    Const SIDE_EDGE_SZ = 3
    Dim AdjustVert      As Long
    Dim AdjustTopEdge   As Long
    Dim AdjustSideEdge  As Long
    Dim ZoomVal         As Long
    Dim TopVal          As Long
    Dim LeftVal         As Long
    Dim WidthVal        As Long
    Dim HeightVal       As Long
    Dim WndCurrent      As Window
    Dim WndNew          As Window
    Dim WasMaximized    As Boolean

    ' NOTE: keeping track on the different windows
    ' for future improvements
    If ActiveWorkbook.Windows.Count = 0 Then
        GoTo CleanExit
    ElseIf ActiveWorkbook.Windows.Count >= 4 Then
        MsgBox "Max number of Windows is 4", vbOkOnly + vbExclamation
        GoTo CleanExit
    End If

    Set WndCurrent = ActiveWindow
    WasMaximized = (WndCurrent.WindowState = xlMaximized)
    If WasMaximized Then
        AdjustVert = 22
        AdjustTopEdge = 3
        AdjustSideEdge = SIDE_EDGE_SZ
    Else
        AdjustTopEdge = 1
        AdjustSideEdge = 0
    End If
    
    With WndCurrent
        ZoomVal = .Zoom
        TopVal = .Top + AdjustVert
        LeftVal = .Left
        WidthVal = .Width
        HeightVal = .Height - AdjustVert
        .WindowState = xlNormal
        Set WndNew = .NewWindow
        .Top = TopVal
        .Left = LeftVal + AdjustSideEdge * 2
    End With

    ' no point of doing it for more than 3 windows
    If Orientation = "Horizontal" Then
        WndCurrent.Height = HeightVal / 2 - AdjustTopEdge
        WndCurrent.Width = WidthVal - SIDE_EDGE_SZ * 2

        WndNew.Top = TopVal + HeightVal / 2 - AdjustTopEdge / 2
        WndNew.Height = HeightVal / 2 - AdjustTopEdge
        WndNew.Left = WndCurrent.Left
        WndNew.Width = WndCurrent.Width
    ElseIf Orientation = "Vertical" Then
        WndCurrent.Height = HeightVal - AdjustTopEdge
        WndCurrent.Width = WidthVal / 2 - AdjustSideEdge * 2

        WndNew.Top = WndCurrent.Top
        WndNew.Height = WndCurrent.Height
        WndNew.Left = LeftVal + WidthVal / 2
        WndNew.Width = WidthVal / 2 - AdjustSideEdge
    End If

    Call ZoomWindow(ZoomVal)
CleanExit:
    FxCallStack.Pop
Exit Sub
UnhandledFail:
    Err.Raise Err.Number, , Err.Description
End Sub

Public Function Validate(Val As Long) As Boolean
    Validate = (Val >= 10 And Val <= 400)
End Function

Public Sub StartInCenter(Frm As Object)
    ' start in the center
    With Application
        Frm.Top = (.Top + .Height - Frm.Height) * 0.5
        Frm.Left = (.Left + .Width - Frm.Width) * 0.5
    End With
End Sub

