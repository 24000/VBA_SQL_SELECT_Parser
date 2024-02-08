VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf 
   Caption         =   "UserForm1"
   ClientHeight    =   2440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   11340
   OleObjectBlob   =   "uf.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "uf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Initialize()
    Me.cmb_sheet.AddItem "�I���N�G��"
    Me.cmb_sheet.AddItem "�A�N�V�����N�G��"
    
    Dim Wb As Workbook
    For Each Wb In Workbooks
        If Left(Wb.name, 1) = "�y" Then
            txt_toolnum.Value = Mid(Wb.name, 2, InStr(Wb.name, "�z") - 2)
            If Wb.ActiveSheet.name = "�I���N�G��" Then
                cmb_sheet = "�I���N�G��"
            ElseIf Wb.ActiveSheet.name = "�A�N�V�����N�G��" Then
                cmb_sheet = "�A�N�V�����N�G��"
            End If
            Exit Sub
        End If
    Next
End Sub

Private Sub spb_rownum_SpinUp()
    If txt_rownum = 1 Then Exit Sub
    If txt_rownum = "" Then Exit Sub
    
    txt_rownum.Value = txt_rownum.Value - 1
End Sub
Private Sub spb_rownum_SpinDown()
    If txt_rownum = "" Then
        txt_rownum.Value = 5
    Else
        txt_rownum.Value = txt_rownum.Value + 1
    End If
End Sub

Private Sub btn_sqlcopy_Click()
    If ExistsAllInput = False Then Exit Sub
    If ExistsTool = False Then Exit Sub
    
    Dim targetWb As xxx_TargetWB: Set targetWb = New xxx_TargetWB
    targetWb.Init txt_toolnum, cmb_sheet, txt_rownum
    Dim rawSql: rawSql = targetWb.GetSQL
    If rawSql = "" Then Exit Sub
    
    Dim parser As x_Parser: Set parser = New x_Parser
    Dim sql() As String: sql = Split(rawSql, vbLf)
    Dim returns As Object
    If cmb_sheet = "�I���N�G��" Then
        Set returns = parser.SelectSQL(sql)
    ElseIf cmb_sheet = "�A�N�V�����N�G��" Then
        Select Case targetWb.GetQueryType
            Case "�ǉ�"
                Set returns = parser.InsertSQL(sql)
            Case "�X�V"
                Set returns = parser.UpdateSQL(sql)
            Case "�V�K�쐬"
                Set returns = parser.SelectIntoSQL(sql)
            Case "�폜"
                Set returns = parser.DeleteSQL(sql)
        End Select
    End If
    
    txt_table = returns("�֘A�e�[�u��")
    txt_about = returns("�T�v")
End Sub

Private Function ExistsAllInput() As Boolean
    If txt_toolnum = "" Then GoTo FALSE_
    If cmb_sheet = "" Then GoTo FALSE_
    If txt_rownum = "" Then GoTo FALSE_
    
    ExistsAllInput = True
    Exit Function
FALSE_:
    ExistsAllInput = False
End Function

Private Function ExistsTool() As Boolean
    Dim wb_ As Workbook
    For Each wb_ In Workbooks
        If InStr(wb_.name, "�y" & txt_toolnum & "�z") > 0 Then
            ExistsTool = True
            Exit Function
        End If
    Next
    MsgBox "�Ώۃc�[���ԍ��̉�͓��e���̓t�@�C��(Excel)���J����Ă��܂���"
    ExistsTool = False
End Function


Private Sub txt_output_Click()
    If ExistsAllInput = False Then Exit Sub
    If ExistsTool = False Then Exit Sub
    
    Dim targetWb As xxx_TargetWB: Set targetWb = New xxx_TargetWB
    targetWb.Init txt_toolnum, cmb_sheet, txt_rownum
    targetWb.PasteTable txt_table
    targetWb.PasteAbout txt_about
End Sub

