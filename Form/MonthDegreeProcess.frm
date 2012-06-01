VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MonthDegreeProcess 
   Caption         =   "���x����"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5625
   OleObjectBlob   =   "MonthDegreeProcess.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "MonthDegreeProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnMonthDegreeProcess_Click()
'���x����

    Dim strState As String
    Dim strPathMy As String
    Dim strMode As String
    Dim tmpSht As Worksheet

    strPathMy = ActiveWorkbook.Path
    If Not strPathMy Like "*�{��*" Then
        If MsgBox("����������S�Ď����ōs���܂��B" & vbCrLf & _
               "�`�F�b�N���ꂽ���ނ��������܂��̂Ńv�����^�[�ɗp����p�ӂ��Ă�������" & vbCrLf & _
               "�������o������OK�{�^���������Ă��������B", vbOKCancel) <> vbOK Then
            MsgBox ("���~���܂���"): GoTo ending
        End If
        '���x�����쐬����ш��
        Call MonthDegreeProcessP2(chk_maruhiro_p.Value, _
                                  chk_tenant_p.Value, _
                                  chk_bill_p.Value, _
                                  chk_uriage_p.Value, _
                                  getBilldateOnStr(getClosingdate))
        If chk_uriage_p.Value = True Then
            '�ۑ�
            ActiveWorkbook.Save
            '�w��t�H���_�ւ̕�
            strState = CopyFile(ActiveWorkbook, "\\honbu\�c�ƕ���\����W�v\���Օi")
            If Not strState Like "ok" Then
                MsgBox ("�{���񍐏����Ɉُ킪����܂����B�{���ւ̕񍐂̓L�����Z������Ă��܂��B" & vbCrLf & _
                       "�V�X�e���Ǘ��҂ɘA�����Ă��������B" & vbCrLf & _
                       strState)
            Else
                MsgBox ("�{���ւ̕񍐂��������܂���")
            End If
        End If
    Else
        Call HonbuView
    End If
    Call PrintMode("put")
ending:
    unload MonthDegreeProcess
    unload SettleForm
    unload MainForm
End Sub

Private Sub btnMonthDegreeProcess2_Click()
'���x�����쐬����ш��
    If cmbBillDate.MatchFound = True Then
        Call MonthDegreeProcessP2(chk_maruhiro_p.Value, _
                                  chk_tenant_p.Value, _
                                  chk_bill_p.Value, _
                                  chk_uriage_p.Value, _
                                  cmbBillDate.text)
        unload MonthDegreeProcess
    Else
        MsgBox "�o�͂��������x��I�����Ă�������"
    End If
End Sub

Private Sub btnReturn_Click()
    unload MonthDegreeProcess
End Sub

Private Sub UserForm_Activate()
    Dim sht As Worksheet
    Dim list() As String
    
    If Not ActiveWorkbook.name Like DataBaseName Then
        MsgBox ("���̃u�b�N��ł͎��s�o���܂���")
        unload SettleForm
    Else
        Call PrintMode("get")
        If getBillDateFromSettleItems(list) = True Then
            cmbBillDate.list = list
            cmbBillDate.text = list(UBound(list))
        Else
            btnMonthDegreeProcess2.Enabled = False
        End If
    End If
End Sub

Private Sub PrintMode(mode As String)
    Dim sht As Worksheet
    Dim address(3) As String
    Set sht = ActiveWorkbook.Worksheets("tmp")
    address(0) = "h10"
    address(1) = "h11"
    address(2) = "h12"
    address(3) = "h13"
    
    Select Case mode
    Case "get"
        chk_maruhiro_p.Value = sht.Range(address(0))
        chk_tenant_p.Value = sht.Range(address(1))
        chk_bill_p.Value = sht.Range(address(2))
        chk_uriage_p.Value = sht.Range(address(3))
    Case "put"
        sht.Range(address(0)) = chk_maruhiro_p.Value
        sht.Range(address(0)).Offset(0, 1) = chk_maruhiro_p.Caption
        sht.Range(address(1)) = chk_tenant_p.Value
        sht.Range(address(1)).Offset(0, 1) = chk_tenant_p.Caption
        sht.Range(address(2)) = chk_bill_p.Value
        sht.Range(address(2)).Offset(0, 1) = chk_bill_p.Caption
        sht.Range(address(3)) = chk_uriage_p.Value
        sht.Range(address(3)).Offset(0, 1) = chk_uriage_p.Caption
    End Select
End Sub

