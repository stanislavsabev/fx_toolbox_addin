Attribute VB_Name = "mImportExport"
Option Explicit
Private Const ModuleName = "mImportExport"

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal ClassName As String, ByVal WindowName As String) As Long

Private Declare PtrSafe Function LockWindowUpdate Lib "user32" _
    (ByVal hWndLock As LongPtr) As Long

Public Const EXP_CONST = "Export"
Public Const IMP_CONST = "Import"
Public Const DEL_CONST = "Delete"
Private Const FAILED_ = "Failed: "

Public TerminateForm    As Boolean
Public HasProtected     As Boolean

Public LastPath         As String
Public FolderPath       As String

Private Type TThis
    Projects    As Object
    Modules     As Object
    Mode        As String
End Type
Dim this As TThis


Public Sub Export()
    FxCallStack.Push ModuleName, "Export"
    On Error GoTo CleanFail
    ExportModules EXP_CONST
CleanExit:
    FxCallStack.Pop
Exit Sub
CleanFail:
    FxCallStack.ShowExcInfoAndClear Err.Number, Err.Description
    Resume CleanExit
End Sub

Public Sub Import()
    FxCallStack.Push ModuleName, "Import"
    On Error GoTo CleanFail
    ImportModules
CleanExit:
    FxCallStack.Pop
Exit Sub
CleanFail:
    FxCallStack.ShowExcInfoAndClear Err.Number, Err.Description
    Resume CleanExit
End Sub

Public Sub Delete()
    FxCallStack.Push ModuleName, "Delete"
    On Error GoTo CleanFail
    ExportModules DEL_CONST
CleanExit:
    FxCallStack.Pop
Exit Sub
CleanFail:
    FxCallStack.ShowExcInfoAndClear Err.Number, Err.Description
    Resume CleanExit
End Sub

Public Function GetMode() As String
    GetMode = this.Mode
End Function

Public Function GetProjects() As Object
    Set GetProjects = this.Projects
End Function

Public Sub ExportModules(Mode As String)
    FxCallStack.Push ModuleName, "ExportModules"
    On Error GoTo UnhandledFail
    Dim Modules As Object
    Dim Results As Object
    Dim Msg     As String
    
    CleanUp
    this.Mode = Mode
    Call ReadVBProjects ' sets this.Projects
    If ShowForm() = False Then Exit Sub

    Set Modules = GetSelectedModules(frmImportExport)
    Set Results = AttemptExporting(Modules)
    Unload frmImportExport

    Msg = GetResultMessage(Results)
    CleanUp
    MsgBox Msg, vbInformation

CleanExit:
    FxCallStack.Pop
Exit Sub
UnhandledFail:
    Err.Raise Err.Number, , Err.Description
End Sub

Private Sub CleanUp()
    FxCallStack.Push ModuleName, "CleanUp"
    On Error GoTo UnhandledFail
    this.Mode = ""
    Set this.Projects = Nothing
    HasProtected = False
    TerminateForm = False
CleanExit:
    FxCallStack.Pop
Exit Sub
UnhandledFail:
    Err.Raise Err.Number, , Err.Description
End Sub

Public Function ExportComponent(Args() As Variant) As Boolean
    FxCallStack.Push ModuleName, "ExportComponent"
    On Error GoTo UnhandledFail
    Dim ext_        As String
    Dim proj        As VBProject ' Object
    Dim comp        As VBComponent '  Object

    Set proj = Args(0)
    Set comp = Args(1)

    Select Case comp.Type
        Case 1: ext_ = ".bas"
        Case 2: ext_ = ".cls"
        Case 3: ext_ = ".frm"
        Case Else
            Err.Raise 5, , "Unknown component type " & comp.Type
    End Select
    
    comp.Activate
    If this.Mode = DEL_CONST Then
        
        proj.VBComponents.Remove comp
    ElseIf this.Mode = EXP_CONST Then
        proj.VBE.SelectedVBComponent.Export LastPath & comp.Name & ext_
    End If

    ExportComponent = True
CleanExit:
    FxCallStack.Pop
Exit Function
UnhandledFail:
    Err.Raise Err.Number, , Err.Description
End Function

Sub ImportModules()
    FxCallStack.Push ModuleName, "ImportModules"
    On Error GoTo UnhandledFail
    Dim Msg         As String
    Dim proj        As Object ' VBProjects
    Dim Modules     As Object
    
    CleanUp
    this.Mode = IMP_CONST
    Call ReadVBProjects

    If ShowForm = False Then Exit Sub

    Set proj = this.Projects(frmImportExport.cboxProjFile.Value)
    Set Modules = GetSelectedModules(frmImportExport)

    Unload frmImportExport

    Dim Imported    As Object ' Scripting Dictionary
    Dim Removed     As Collection
    Dim Reuslts     As Object ' Scripting.Dictionary

    Set Removed = RemoveExistingAndBlacklisted(proj, Modules)
    Set Imported = AttemptImport(proj, Modules)
    Set Reuslts = JoinWithFailed(Removed, Imported)

    Msg = GetResultMessage(Reuslts)
    CleanUp
    MsgBox Msg, vbInformation
CleanExit:
    FxCallStack.Pop
Exit Sub
UnhandledFail:
    Err.Raise Err.Number, , Err.Description
End Sub

Private Function JoinWithFailed(Removed As Collection, Imported As Object) As Object
    FxCallStack.Push ModuleName, "JoinWithFailed"
    On Error GoTo UnhandledFail
    Dim v As Variant

    For Each v In Removed
        Imported(CStr(v)) = False
    Next v
    Set JoinWithFailed = Imported
CleanExit:
    FxCallStack.Pop
Exit Function
UnhandledFail:
    Err.Raise Err.Number, , Err.Description
End Function

Private Function RemoveExistingAndBlacklisted(proj As Object, Modules As Object) As Collection
    FxCallStack.Push ModuleName, "RemoveExistingAndBlacklisted"
    On Error GoTo UnhandledFail
    Dim comp    As Object ' VBComponent
    Dim Removed  As Collection
    Set Removed = New Collection

    ' remove existing ...
    For Each comp In proj.VBComponents
        Select Case comp.Type
            Case 1, 2, 3 ' ...is module, class or userform
                If Modules.Exists(comp.Name) Then ' ... self assignment not allowed
                    If IsBlacklisted(proj, comp.Name) Then
                        Removed.Add comp.Name       ' report fail
                        Modules.Remove comp.Name    ' removeing from the import list
                    Else
                        On Error Resume Next
                        Err.Clear
                        proj.VBComponents.Remove comp

                        If Err.Number <> 0 Then
                            Removed.Add comp.Name   ' report fail
                            Modules.Remove comp.Name ' removing from the import list
                        End If
                        On Error GoTo -1
                    End If
                End If
        End Select
    Next comp

CleanExit:
    Set RemoveExistingAndBlacklisted = Removed
    FxCallStack.Pop
Exit Function
UnhandledFail:
    Err.Raise Err.Number, , Err.Description
End Function

Private Function GetResultMessage(Results As Object) As String
    FxCallStack.Push ModuleName, "GetResultMessage"
    On Error GoTo UnhandledFail
    Dim Msg         As String
    Dim Succeded    As New Collection
    Dim Failed      As New Collection
    Dim v           As Variant

    For Each v In Results.Keys()
        If Results(v) = True Then
            Succeded.Add CStr(v)
        Else
            Failed.Add CStr(v)
        End If
    Next v

    Msg = Succeded.Count & " " & this.Mode & GetEnding(this.Mode) & vbCrLf
    If Succeded.Count Then
        Msg = Msg & vbTab & fx.CollectionToString(Succeded) & vbCrLf
    End If

    If Failed.Count Then
        Msg = Msg & Failed.Count & " " & FAILED_ & vbCrLf
        Msg = Msg & vbTab & fx.CollectionToString(Failed)
    End If
    GetResultMessage = Msg
CleanExit:
    FxCallStack.Pop
Exit Function
UnhandledFail:
    Err.Raise Err.Number, , Err.Description
End Function

Private Function AttemptImport(proj As Object, Modules As Object) As Object
    FxCallStack.Push ModuleName, "AttemptImport"
    On Error GoTo UnhandledFail
    Dim k       As Variant
    Dim Msg     As String
    Dim FName   As String
    Dim sz      As Long
    Dim Results As Object
    Set Results = CreateObject("Scripting.Dictionary")

    ' attemping import
    For Each k In Modules.Keys() ' Dict is 0 based
        FName = Modules(k)

        On Error Resume Next
        proj.VBComponents.Import LastPath & FName
        Results.Add k, (Err.Number = 0) ' report outcome
        On Error GoTo -1
    Next k
CleanExit:
    Set AttemptImport = Results
    FxCallStack.Pop
Exit Function
UnhandledFail:
    Err.Raise Err.Number, , Err.Description
End Function

Private Sub ReadVBProjects()
    FxCallStack.Push ModuleName, "ReadVBProjects"
    On Error GoTo UnhandledFail
    Dim env     As Object ' VBE
    Dim proj    As Object ' VBProject
    Dim comp    As Object ' VBComponent
    Dim Msg     As String

    Set env = GetEnv()
    If this.Projects Is Nothing Then
        Set this.Projects = CreateObject("Scripting.Dictionary")
    Else
        this.Projects.RemoveAll
    End If

    HasProtected = False
    For Each proj In env.VBProjects
        ' if not protected and Not Titus leak !
            ' 0 = vbext_pp_none
        If proj.Protection = 0 And Not proj.BuildFileName = "VBAProject.DLL" Then
            ' FileName: VBProject Object pair
            this.Projects.Add fx.DropFilePath(proj.FileName), proj
        Else
            HasProtected = True
        End If
    Next proj

    ' sorting projects here!!!
    If this.Projects.Count = 0 Then
        Msg = "There are no accessible projects." & vbNewLine
        Msg = Msg & "Unprotect and try again!"
        Err.Raise 5, , Msg
    End If

CleanExit:
    FxCallStack.Pop
Exit Sub
UnhandledFail:
    Err.Raise Err.Number, , Err.Description
End Sub

Public Function GetEnv() As Object
    FxCallStack.Push ModuleName, "GetEnv"
    Dim Msg     As String
    Dim Result  As Object

    On Error Resume Next
    Set Result = Application.VBE
    On Error GoTo 0
    
    On Error GoTo UnhandledFail
    If Result Is Nothing Then ' ..error if failed
        Msg = "Unable to connect to ThisWorkbook.VBProject" & vbNewLine
        Msg = Msg & "Possible reason: Trust Access to the VBA project model is disabled"
        Err.Raise 5, , Msg & Err.Description
    ElseIf Result.VBProjects.Count = 0 Then
        Msg = "No VBProjects detected" & vbNewLine
        Err.Raise 5, , Msg & Err.Description
    End If
    
CleanExit:
    Set GetEnv = Result
    FxCallStack.Pop
Exit Function
UnhandledFail:
    Err.Raise Err.Number, , Err.Description
End Function

Private Function IsBlacklisted(proj As Object, Name As String) As Boolean
    FxCallStack.Push ModuleName, "IsBlacklisted"
    On Error GoTo UnhandledFail
    Dim v   As Variant

    ' no blacklisted in external proj
    If proj.FileName <> Application.VBE.ActiveVBProject.FileName Then
        IsBlacklisted = False
        GoTo CleanExit
    End If

    For Each v In Array("mRibbon", "mVersionControl", "frmImportExport", "mHelpers", "mMouseMove")
        If CStr(v) = Name Then
            IsBlacklisted = True
            GoTo CleanExit
        End If
    Next v

    IsBlacklisted = False
CleanExit:
    FxCallStack.Pop
Exit Function
UnhandledFail:
    Err.Raise Err.Number, , Err.Description
End Function

Private Function GetEnding(Mode As String) As String
    GetEnding = (IIf(Right(this.Mode, 1) = "e", "d", "ed: "))
End Function

Public Sub Browse()
    FxCallStack.Push ModuleName, "Browse"
    On Error GoTo UnhandledFail

    FolderPath = fx.Browse(FolderPath, _
        this.Mode & " Folder Location...", _
        msoFileDialogFolderPicker)
CleanExit:
    FxCallStack.Pop
Exit Sub
UnhandledFail:
    Err.Raise Err.Number, , Err.Description
End Sub

Private Function GetSelectedModules(Frm As UserForm) As Object
    FxCallStack.Push ModuleName, "GetSelectedModules"
    On Error GoTo UnhandledFail
    Dim Modules As Object
    Dim i       As Long
    Set Modules = CreateObject("Scripting.Dictionary")
    With Frm
        ' setup
        LastPath = fx.TrailingSlash(FolderPath)

        ' get selected modules
        For i = 0 To .lboxModules.ListCount - 1
            If .lboxModules.Selected(i) = True Then
                ' Key is module name (-extension), Value is the filename
                Modules.Add Split(.lboxModules.List(i), ".")(0), .lboxModules.List(i)
            End If
        Next i
    End With

CleanExit:
    Set GetSelectedModules = Modules
    FxCallStack.Pop
Exit Function
UnhandledFail:
    Err.Raise Err.Number, , Err.Description
End Function

Private Function AttemptExporting(Modules As Object) As Object
    FxCallStack.Push ModuleName, "AttemptExporting"
    On Error GoTo UnhandledFail
    Dim comp        As Object ' VBComponent
    Dim proj        As Object ' VBProject
    Dim v           As Variant
    Dim Name        As String
    Dim Results     As Object
    Set Results = CreateObject("Scripting.Dictionary")

    Set proj = this.Projects(frmImportExport.cboxProjFile.Value)

    ' exporting if..
    For Each comp In proj.VBComponents
        Select Case comp.Type
            Case 1, 2, 3 ' ...is module, class or userform
                Name = comp.Name ' reading the name to use it for result
                If Modules.Exists(comp.Name) Then ' ...matches the criteria
                    v = StopScreenFlicker("ExportComponent", Array(proj, comp))
                    Results(Name) = v ' attempting exoprt
                End If
        End Select
    Next comp

CleanExit:
    Set AttemptExporting = Results
    FxCallStack.Pop
Exit Function
UnhandledFail:
    Err.Raise Err.Number, , Err.Description
End Function

Public Sub SetInitialExportPath(ProjFileName As String)
    FxCallStack.Push ModuleName, "SetInitialExportPath"
    On Error GoTo UnhandledFail
    Dim SubfoldersArr   As Variant
    Dim v               As Variant
    Dim fldr            As Object
    Dim fd              As Object

    SubfoldersArr = Array("repo", "*.repo", "src", "vba", "*?.git")

    FolderPath = LastPath
    If FolderPath = "" Then
        FolderPath = this.Projects(ProjFileName).FileName
    End If

    FolderPath = fx.GetFilePath(FolderPath)
    Set fldr = fx.Fso.GetFolder(FolderPath)

    For Each fd In fldr.SubFolders
        For Each v In SubfoldersArr
            If fx.Name Like v Then
                FolderPath = FolderPath & fd.Name
                GoTo CleanExit
            End If
        Next v
    Next fd
CleanExit:
    FxCallStack.Pop
Exit Sub
UnhandledFail:
    Err.Raise Err.Number, , Err.Description
End Sub

Private Function ShowForm() As Boolean
    FxCallStack.Push ModuleName, "ShowForm"
    On Error GoTo UnhandledFail
    TerminateForm = False
    frmImportExport.Show
    If TerminateForm Then
        Unload frmImportExport
        ShowForm = False
    Else
        ShowForm = True
    End If
CleanExit:
    FxCallStack.Pop
Exit Function
UnhandledFail:
    Err.Raise Err.Number, , Err.Description
End Function


Public Function StopScreenFlicker(ProcName As String, Args As Variant) As Variant
    FxCallStack.Push ModuleName, "StopScreenFlicker"
    On Error GoTo HandledFail
    Dim VBEHwnd As Long
    Dim Result  As Variant

    With Application.VBE.MainWindow
        .Visible = False
        VBEHwnd = FindWindow("wndclass_desked_gsk", .Caption)

        ' client code
        Result = Application.Run(ProcName, Args)
        .Visible = True
    End With

CleanExit:
    StopScreenFlicker = Result
    FxCallStack.Pop
Exit Function
HandledFail:
    '// Handle error locally, sync the call stack and resume
    LockWindowUpdate 0&
    FxCallStack.Sync ModuleName, "StopScreenFlicker"
    Resume CleanExit
End Function
