VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImportExport 
   Caption         =   "Export/Import Code"
   ClientHeight    =   8520.001
   ClientLeft      =   108
   ClientTop       =   384
   ClientWidth     =   12588
   OleObjectBlob   =   "frmImportExport.frx":0000
End
Attribute VB_Name = "frmImportExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type TThis
    IgnoreEvents        As Boolean
    Shift               As Boolean
    lboxIgnoreEvents    As Boolean
End Type
Private this As TThis


Private Sub UserForm_Initialize()
    On Error GoTo Finally
    this.IgnoreEvents = True
    
    ClearListElements
    
    Call AddProjects
    Call ShowStatus
    Call mImportExport.SetInitialExportPath(cboxProjFile.Value)
    Call SetupFolderPathUIElements
    UpdateListBoxContent
    
    btnExport.Caption = mImportExport.GetMode
    Me.Caption = btnExport.Caption & " Code"
    
    ' select all by default
    UpdateListBoxSelection 0, lboxModules.ListCount - 1
Finally:
    this.IgnoreEvents = False
End Sub

Private Sub ClearListElements()
    cboxProjFile.Clear
    ClearModules
End Sub

Private Sub ClearModules()
    lboxModules.Clear
    txtFilterModules.Value = ""
End Sub

Private Sub txtPath_Change()
    Dim IsFolder As Boolean
    Dim Diff     As Long
    
    mImportExport.FolderPath = txtPath.Value
    IsFolder = fx.FolderExists(mImportExport.FolderPath)
    
    If mImportExport.GetMode = IMP_CONST Then
        If IsFolder Then
            UpdateListBoxContent
        Else
            ClearModules
            lboxModules.AddItem "Not a folder..."
        End If
    End If
    
    If Not IsFolder Then
        txtPath.BorderColor = RGB(255, 0, 0)
        txtPath.BorderStyle = fmBorderStyleSingle
    Else
        txtPath.BorderStyle = fmBorderStyleNone
    End If
End Sub

Private Sub lboxModules_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
        ByVal x As Single, ByVal y As Single)
    this.Shift = (Shift And 1)
End Sub

''' be sure to include Error handling for any code that
''' might get called while the hook is running
Private Sub lboxModules_Change()
    If this.lboxIgnoreEvents Then Exit Sub
    
    Dim lMax        As Long
    Dim lMin        As Long
    Dim lCurrIdx    As Long
    Static lPriorIdx As Long
    
    lCurrIdx = Me.lboxModules.ListIndex
    
    On Error GoTo errExit
    If this.Shift And (lCurrIdx <> lPriorIdx) Then
        If lPriorIdx < lCurrIdx Then
            Call UpdateListBoxSelection(lPriorIdx, lCurrIdx)
        Else
            Call UpdateListBoxSelection(lCurrIdx, lPriorIdx)
        End If
    End If
    
    lPriorIdx = lCurrIdx
errExit:
End Sub

'Private Sub lboxModules_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
'        ByVal x As Single, ByVal y As Single)
'    HookListBoxScroll
'End Sub

Private Sub btnBrowse_Click()
    mImportExport.Browse
    txtPath.Value = mImportExport.FolderPath
End Sub

Private Sub btnExport_Click()
    If AnyModuleSelected Then
        Me.Hide
    Else
        MsgBox "No Modules selected!", vbExclamation
    End If
End Sub

Private Sub btnClose_Click()
    mImportExport.TerminateForm = True
    Me.Hide
End Sub

Private Sub OnSelectDeselectAllUpdateUI()
    txtFilterModules.Value = ""
    chkInvert.Value = False
End Sub

Private Sub btnDeselectAll_Click()
    UpdateListBoxSelection -1
    OnSelectDeselectAllUpdateUI
End Sub

Private Sub btnSelectAll_Click()
    UpdateListBoxSelection 0, lboxModules.ListCount - 1
    OnSelectDeselectAllUpdateUI
End Sub

Private Sub cboxProjFile_Change()
    If this.IgnoreEvents Then Exit Sub
    Call mImportExport.SetInitialExportPath(cboxProjFile.Value)
    txtPath.Value = mImportExport.FolderPath
    UpdateListBoxContent
End Sub

Private Sub chkClasses_Change()
    UpdateListBoxContent
End Sub

Private Sub chkForms_Change()
    UpdateListBoxContent
End Sub

Private Sub chkSheetModules_Change()
    UpdateListBoxContent
End Sub

Private Sub chkModules_Change()
    UpdateListBoxContent
End Sub

Private Sub chkInvert_Change()
    InvertSelected
End Sub

Private Sub UpdateListBoxContent()
    Dim proj    As Object ' VBProject
    Dim Itms    As New Collection
    Dim i       As Long
    Dim Mode    As String
    
    On Error GoTo failed
    lboxModules.Clear
    lboxModules.AddItem "Updating...", 0
    
    Mode = mImportExport.GetMode
    If Mode = EXP_CONST Or Mode = DEL_CONST Then
        Set proj = mImportExport.GetProjects(cboxProjFile.Value)
        Set Itms = GetFilteredModuleNames(proj)
    ElseIf Mode = IMP_CONST Then
        If mImportExport.FolderPath <> "" Then
            Set Itms = UpdateImportList(mImportExport.FolderPath)
        End If
    End If
    
    If Itms.Count = 0 Then GoTo failed
    
    Call fx.SortCollectionInplace(Itms)
    ClearModules
    
    For i = 1 To Itms.Count
        lboxModules.AddItem Itms(i)
    Next i
Exit Sub
failed:
    ClearModules
    lboxModules.AddItem "Empty..."
End Sub

Private Function GetFilteredModuleNames(proj As Object) As Collection
    Dim Itms As New Collection
    Dim comp As Object
    
    For Each comp In proj.VBComponents
        Select Case comp.Type ' ...is module, class or userform
            Case 1 And chkModules
                Itms.Add comp.Name
            Case 2 And chkClasses
                Itms.Add comp.Name
            Case 3 And chkForms
                Itms.Add comp.Name
            Case 100 And chkSheetModules
                Itms.Add comp.Name
        End Select
    Next comp
    Set GetFilteredModuleNames = Itms
End Function

Private Function UpdateImportList(FolderPath As String) As Collection
    Dim Itms    As New Collection
    Dim fldr    As Object
    Dim f       As Object
    
    Set fldr = CreateObject("Scripting.FileSystemObject").GetFolder(mImportExport.FolderPath)
    For Each f In fldr.Files
        If Not Left(f.Name, 1) = "~" Then
            Select Case Right(f.Name, 4)
                Case ".bas": If chkModules Then Itms.Add f.Name
                Case ".frm": If chkForms Then Itms.Add f.Name
                Case ".bas": If chkClasses Then Itms.Add f.Name
            End Select
        End If
    Next f
    Set UpdateImportList = Itms
End Function

Private Sub UpdateListBoxSelection(lStart As Long, Optional lEnd As Long = -1)
    Dim i   As Long
    Dim Val As Boolean
    
    this.lboxIgnoreEvents = True
    
    For i = 0 To lboxModules.ListCount - 1
        Val = (IIf(lStart = -1, False, (i >= lStart And i <= lEnd)))
        lboxModules.Selected(i) = Val
    Next i
    
    this.lboxIgnoreEvents = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'    UnhookListBoxScroll
    mImportExport.TerminateForm = True
End Sub

Private Sub AddProjects()
    Dim Projects    As Object
    Dim v           As Variant
    Dim i           As Long
    
    Set Projects = mImportExport.GetProjects
    
    If Application.Name = "Microsoft Excel" Then
        Call AddProjectsExcel(Projects)
    ElseIf Application.Name = "Microsoft Access" Then
        Call AddProjectsMSAccess(Projects)
    End If
    
    cboxProjFile.ListIndex = 0
End Sub

Private Sub AddProjectsExcel(Projects As Object)
    Dim v       As Variant
    Dim WbCount As Long
    
    WbCount = Workbooks.Count
    With cboxProjFile
        For Each v In Projects.Keys()
            If WbCount > 0 And CStr(v) = ActiveWorkbook.Name Then
                .AddItem CStr(v), 0
                WbCount = 0
            Else
                .AddItem CStr(v), .ListCount
            End If
        Next v
    End With
End Sub

Private Sub AddProjectsMSAccess(Projects As Object)
    Dim v       As Variant
    
    With cboxProjFile
        For Each v In Projects.Keys()
            .AddItem CStr(v), .ListCount
        Next v
    End With
End Sub

Private Sub ShowStatus()
    If mImportExport.HasProtected Then
        txtStatus.Value = "* Not showing protected VBProjects"
        txtStatus.BackStyle = fmbackStyleOpaque
    Else
        txtStatus.Value = ""
        txtStatus.BackStyle = fmbackStyleTransparent
    End If
End Sub

Private Sub SetupFolderPathUIElements()
    If mImportExport.GetMode = DEL_CONST Then
        txtPath.Visible = False
        btnBrowse.Visible = False
        lblPath.Visible = False
    Else
        txtPath.Visible = True
        btnBrowse.Visible = True
        lblPath.Visible = True
        
        txtPath.Value = mImportExport.FolderPath
    End If
End Sub

Private Sub txtFilterModules_Change()
    Call FilterModulesByText(txtFilterModules.Value)
End Sub

Private Sub FilterModulesByText(Txt As String)
    Dim i As Long
    Dim newListNdx As Long
    
    UpdateListBoxSelection -1
    
    With lboxModules
        .ListIndex = -1
        newListNdx = -1
        
        For i = 0 To .ListCount - 1
            If InStr(1, .List(i, 0), Txt, vbBinaryCompare) > 0 Then
                If newListNdx = -1 Then newListNdx = i
                .Selected(i) = True
            End If
        Next i
        .ListIndex = newListNdx
    End With
    
    If chkInvert.Value Then InvertSelected
End Sub

Private Sub InvertSelected()
    Dim i   As Long
    Dim Val As Boolean
    Dim newListNdx As Long
    
    If Not AnyModuleSelected Then Exit Sub
    
    newListNdx = -1
    With lboxModules
        For i = 0 To .ListCount - 1
            If newListNdx = -1 Then newListNdx = i
            .Selected(i) = (Not .Selected(i))
        Next i
    End With
End Sub

Private Function AnyModuleSelected() As Boolean
    Dim i As Long
    
    With lboxModules
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                AnyModuleSelected = True
                Exit Function
            End If
        Next i
    End With
    
    AnyModuleSelected = False
End Function
