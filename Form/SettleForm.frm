VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettleForm 
   Caption         =   "����"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   OleObjectBlob   =   "SettleForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "SettleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnInventoryReport_Click()
'�I������
    Dim strState As String
    Dim strPathMy As String
    Dim strMode As String
        
    If MsgBox("�I�����񍐂�{���֑���܂��B" & vbCrLf & _
              "��낵���ł����H", vbOKCancel) <> vbOK Then
        MsgBox ("���~���܂���"): GoTo ending
    End If
    strState = CopyFile(ActiveWorkbook, "\\honbu\�c�ƕ���\����W�v\���Օi\�I����")
    If Not strState Like "ok" Then
        MsgBox ("�{���񍐏����Ɉُ킪����܂����B�{���ւ̕񍐂̓L�����Z������Ă��܂��B" & vbCrLf & _
               "�V�X�e���Ǘ��҂ɘA�����Ă��������B" & vbCrLf & _
               strState)
    Else
        MsgBox ("�{���ւ̕񍐂��������܂���")
    End If
ending:
End Sub

Private Sub btnMain_Click()
    Sheets("���C��").Select
    unload SettleForm
End Sub

Private Sub btnmakeApprovalList_Click()
    MsgBox (makeApprovalList)
    Range("a1").Select
    unload SettleForm
    unload MainForm
End Sub

Private Sub btnMonthDegreeProcess_Click()
    MonthDegreeProcess.Show
End Sub

Private Sub btnStockReport_Click()
'�݌ɕ�
    Dim strState As String
    Dim strPathMy As String
    Dim strMode As String
        
    If MsgBox("���݂̍݌ɏ󋵂�{���֕񍐂��܂��B" & vbCrLf & _
              "��낵���ł����H", vbOKCancel) <> vbOK Then
        MsgBox ("���~���܂���"): GoTo ending
    End If
    strState = CopyFile(ActiveWorkbook, "\\honbu\�c�ƕ���\����W�v\���Օi\�݌ɏ󋵕�")
    If Not strState Like "ok" Then
        MsgBox ("�{���񍐏����Ɉُ킪����܂����B�{���ւ̕񍐂̓L�����Z������Ă��܂��B" & vbCrLf & _
               "�V�X�e���Ǘ��҂ɘA�����Ă��������B" & vbCrLf & _
               strState)
    Else
        MsgBox ("�{���ւ̕񍐂��������܂���")
    End If
ending:

End Sub

Private Sub CommandButton2_Click()
    MsgBox ("���̋@�\�͎�������Ă��܂���")
End Sub


Private Sub UserForm_Activate()
    If Not ActiveWorkbook.name Like DataBaseName Then
        MsgBox ("���̃u�b�N��ł͎��s�o���܂���")
        unload SettleForm
    End If
End Sub

