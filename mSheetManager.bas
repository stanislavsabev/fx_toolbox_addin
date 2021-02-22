Attribute VB_Name = "mSheetManager"
Option Explicit
Private Const ModuleName = "mSheetManager"


Public Sub Protect()
    FxCallStack.Push ModuleName, "Protect"
    On Error GoTo CleanFail
    SheetProtectUnprotect "Protect"
CleanExit:
    FxCallStack.Pop
Exit Sub
CleanFail:
    FxCallStack.ShowExcInfoAndClear Err.Number, Err.Description
    Resume CleanExit
End Sub

Public Sub Unprotect()
    FxCallStack.Push ModuleName, "Unprotect"
    On Error GoTo CleanFail
    SheetProtectUnprotect "Unprotect"
CleanExit:
    FxCallStack.Pop
Exit Sub
CleanFail:
    FxCallStack.ShowExcInfoAndClear Err.Number, Err.Description
    Resume CleanExit
End Sub

Private Sub SheetProtectUnprotect(ActionName As String)
    FxCallStack.Push ModuleName, "SheetProtectUnprotect"
    On Error GoTo UnhandledFail
    Dim Wks         As Worksheet
    Dim Pw          As String
    Dim Failed      As New Collection
    Dim Cancelled   As Boolean

    Pw = GetPassword("Password: ", Cancelled)
    If Cancelled Then GoTo CleanExit

    For Each Wks In ActiveWorkbook.Sheets
        If ApplyAction(Wks, Pw, ActionName) = False Then
            Failed.Add Wks.Name
        End If
    Next Wks

    If Failed.Count Then
        MsgBox _
            "Failed for " & CStr(Failed.Count) & " sheets:" _
                & vbNewLine & fx.CollectionToString(Failed), _
            vbExclamation, "Warning"
    Else
        MsgBox "Done!", vbInformation
    End If
CleanExit:
    FxCallStack.Pop
Exit Sub
UnhandledFail:
    Err.Raise Err.Number, , Err.Description
End Sub

Private Function ApplyAction(Wks As Worksheet, Pw As String, _
        ActionName As String) As Boolean
    FxCallStack.Push ModuleName, "ApplyAction"
    On Error GoTo CleanFail
    Dim Result As Boolean
    Dim Skip   As Boolean

    If ActionName = "Protect" Then
        Skip = True
    ElseIf ActionName = "Unprotect" Then
        Skip = False
    End If

    If Wks.ProtectContents = Skip Then
        Result = True
        GoTo CleanExit
    End If

    Call VBA.CallByName(Wks, ActionName, VbCallType.VbMethod, Pw)
    Result = True
CleanExit:
    ApplyAction = Result
    FxCallStack.Pop
Exit Function
CleanFail:
    Result = False
    Resume CleanExit
End Function

Private Function GetPassword(ByVal Caption As String, _
        ByRef OutCancelled As Boolean) As String
    FxCallStack.Push ModuleName, "GetPassword"
    On Error GoTo UnhandledFail
    Dim Result As Variant
    OutCancelled = False
    
    Result = Application.InputBox(Caption, "Password Box", Type:=2)
    OutCancelled = TypeName(Result) = "Boolean" And Result = False
CleanExit:
    GetPassword = Result
    FxCallStack.Pop
Exit Function
UnhandledFail:
    OutCancelled = True
    Err.Raise Err.Number, , Err.Description
End Function

Public Sub UnhideAll()
    FxCallStack.Push ModuleName, "UnhideAll"
    On Error GoTo CleanFail
    Dim Wks As Object

    With Application
        .ScreenUpdating = False
        .EnableEvents = False

        For Each Wks In ActiveWorkbook.Sheets
            If Wks.Visible <> xlSheetVisible Then
                Wks.Visible = xlSheetVisible
            End If
        Next Wks

        .ScreenUpdating = True
        .EnableEvents = True
    End With
CleanExit:
    FxCallStack.Pop
Exit Sub
CleanFail:
    FxCallStack.ShowExcInfoAndClear Err.Number, Err.Description
    Resume CleanExit
End Sub

Public Sub HideNotSelected()
    FxCallStack.Push ModuleName, "HideNotSelected"
    On Error GoTo CleanFail
    Dim SelWks  As Object
    Dim Wks     As Object
    Dim Vis     As XlSheetVisibility

    With Application
        .ScreenUpdating = False
        .EnableEvents = False

        For Each Wks In ActiveWorkbook.Sheets
            Vis = xlSheetHidden
            For Each SelWks In ActiveWorkbook.SelectedSheets
                If Wks.Name = SelWks.Name Then
                    Vis = xlSheetVisible
                    Exit For
                End If
            Next SelWks
            Wks.Visible = Vis
        Next Wks

        .ScreenUpdating = True
        .EnableEvents = True
    End With

CleanExit:
    FxCallStack.Pop
Exit Sub
CleanFail:
    FxCallStack.ShowExcInfoAndClear Err.Number, Err.Description
    Resume CleanExit
End Sub

Public Sub VeryHideSelected()
    FxCallStack.Push ModuleName, "VeryHideSelected"
    Dim Wks     As Object
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False

        On Error Resume Next
        For Each Wks In ActiveWindow.SelectedSheets
            Wks.Visible = xlSheetVeryHidden
        Next Wks
        On Error GoTo 0

        .ScreenUpdating = True
        .EnableEvents = True
    End With

End Sub

Public Sub Crop()
    FxCallStack.Push ModuleName, "Crop"
    On Error GoTo CleanFail

    If Workbooks.Count = 0 Then GoTo CleanExit

    If TypeName(ActiveSheet) = "Chart" Then
        Err.Raise 445, , fx.Error(445)
    ElseIf ActiveSheet.ProtectContents Then
        Err.Raise 70, , fx.Error(70)
    End If
    
    Call ApplyCrop(True)
  
CleanExit:
    FxCallStack.Pop
Exit Sub
CleanFail:
    FxCallStack.ShowExcInfoAndClear Err.Number, Err.Description
    Resume CleanExit
End Sub

Public Sub UnCrop()
    FxCallStack.Push ModuleName, "UnCrop"
    On Error GoTo CleanFail


    If TypeName(ActiveSheet) = "Chart" Then
        Err.Raise 5, , "Can't uncrop Chart sheet!"
    End If

    Call ApplyCrop(False)

CleanExit:
    FxCallStack.Pop
Exit Sub
CleanFail:
    FxCallStack.ShowExcInfoAndClear Err.Number, Err.Description
    Resume CleanExit
End Sub

Private Sub ApplyCrop(Value As Boolean)
    FxCallStack.Push ModuleName, "ApplyCrop"
    On Error GoTo UnhandledFail
    Dim lCol    As Long
    Dim lRow    As Long

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    With WorksheetFunction
        lCol = .Min(ActiveCell.Column + 1, Columns.Count)
        lRow = .Min(ActiveCell.Row + 1, Rows.Count)
    End With

    With ActiveSheet
        Range(.Columns(lCol), .Columns(Columns.Count)).Hidden = Value
        Range(.Rows(lRow), .Rows(Rows.Count)).Hidden = Value
    End With

    Application.ScreenUpdating = True
    Application.EnableEvents = True
CleanExit:
    FxCallStack.Pop
Exit Sub
UnhandledFail:
    Err.Raise Err.Number, , Err.Description
End Sub
