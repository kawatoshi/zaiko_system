Attribute VB_Name = "InportExportModule"
''-----------------------------------------------------------------------
'' �S�v���W�F�N�g�t�@�C���G�N�X�|�[�g�i�u�b�N�E�V�[�g�ɕt������R�[�h�ȊO�j
'' ���O�Ƀ}�N���̃Z�L�����e�B��VBA�̃I�u�W�F�N�g���f���ւ̃A�N�Z�X�������鎖�i���s���G���[�ɂȂ�܂��B�j
''-----------------------------------------------------------------------
Private Sub Export_All()
 
    Dim Path As String
    Dim i As Integer
     
    Const cls As String = "\Class\"
    Const FRM As String = "\Form\"
    Const MODL As String = "\Module\"
     
    Const EXT_MODL As String = ".bas"
    Const EXT_CLS As String = ".cls"
    Const EXT_FRM As String = ".frm"
     
    Path = ThisWorkbook.Path
     
    '' �G�N�X�|�[�g�t�H���_
    If Dir(Path & cls) = "" Then MkDir (Path & cls)
    If Dir(Path & FRM) = "" Then MkDir (Path & FRM)
    If Dir(Path & MODL) = "" Then MkDir (Path & MODL)
     
    With ActiveWorkbook.VBProject
     
        For i = 1 To .VBComponents.Count
         
            Select Case .VBComponents(i).Type
            Case 1  '' vbCompTypeModul
                .VBComponents(i).Export Path & MODL & .VBComponents(i).name & EXT_MODL
            Case 2 '' vbCompTypeClassModul
                .VBComponents(i).Export Path & cls & .VBComponents(i).name & EXT_CLS
            Case 3 '' vbCompTypeUserform
                .VBComponents(i).Export Path & FRM & .VBComponents(i).name & EXT_FRM
            End Select
        Next
     
    End With
     
End Sub
 
 
''-----------------------------------------------
''--�v���W�F�N�g�t�@�C���􂢑ւ�-----------------
''-----------------------------------------------
Private Sub Refresh()
 
    Call Release_All
    Call Import_All
 
End Sub
 
 
'' �S�v���W�F�N�g�t�@�C�������[�X
Private Sub Release_All()
 
    Dim i As Integer
    Dim colComName As New Collection
     
    With ThisWorkbook.VBProject
    
        For i = 1 To .VBComponents.Count
            If .VBComponents(i).Type = 1 Or .VBComponents(i).Type = 2 Or .VBComponents(i).Type = 3 Then
                colComName.Add (.VBComponents(i).name)
            End If
        Next
     
        For i = 1 To colComName.Count
            .VBComponents.Remove .VBComponents(colComName(i))
        Next
     
    End With
     
    Set colComName = Nothing
     
 
End Sub
 
'' �S�v���W�F�N�g�t�@�C���C���|�[�g
Private Sub Import_All()
 
    Dim Path As String
    Dim i As Integer
     
    Const cls As String = "\Class\"
    Const FRM As String = "\Form\"
    Const MODL As String = "\Module\"
     
    Const EXT_MODL As String = ".bas"
    Const EXT_CLS As String = ".cls"
    Const EXT_FRM As String = ".frm"
     
    Path = ThisWorkbook.Path
     
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fileList As Object
     
    '' Cls
    Set fileList = fso.GetFolder(Path & cls).Files
    For Each file In fileList
        ActiveWorkbook.VBProject.VBComponents.Import Path & cls & file.name
    Next
     
    '' Form
    Set fileList = fso.GetFolder(Path & FRM).Files
    For Each file In fileList
        If Right(file.name, 4) = EXT_FRM Then
            ActiveWorkbook.VBProject.VBComponents.Import Path & FRM & file.name
        End If
    Next
     
    '' Module
    Set fileList = fso.GetFolder(Path & MODL).Files
    For Each file In fileList
        ActiveWorkbook.VBProject.VBComponents.Import Path & MODL & file.name
    Next
     
 
End Sub
