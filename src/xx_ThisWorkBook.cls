VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xx_ThisWorkBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Class_Initialize()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.ClearContents
    Next
End Sub

Public Sub PasteSelect(sqlLines As Variant)
    With ThisWorkbook.Worksheets("SQL��")
        .Cells.ClearContents
        .Cells(1, 1) = "SQL��"
        Dim lineNum As Long
        For lineNum = LBound(sqlLines) To UBound(sqlLines)
            Sheet1.Cells(lineNum + 2, 1) = sqlLines(lineNum)
        Next
    End With
End Sub

Private Sub RowHeightAdjust()
    
End Sub

'SelectParser��GetParsedColumnPhrases�̃V�[�g�\��t�����������Ă��āA��������
Public Sub PasteParsedSelect(basePhrase As String, ColumnPhrases As Collection)
    With ThisWorkbook.Worksheets("Select")
        .Cells(1, 1) = "Select��"
        .Cells(2, 1) = basePhrase
        .Cells(1, 2) = "�񖼂̂݉��s�\��"
        .Cells(3, 2) = "�\����"
        .Cells(3, 3) = "�e�[�u����"
        .Cells(3, 4) = "����"
        .Cells(3, 6) = "�W�v"
        Dim phrase As Variant
        Dim row As Long: row = 4
        For Each phrase In ColumnPhrases
            .Cells(row, 2) = phrase("displayName")
            .Cells(row, 3) = phrase("tableName")
            .Cells(row, 4) = phrase("columnName")
            .Cells(row, 6) = phrase("aggregate")
            If Right(phrase("sourcePhrase"), 1) <> vbCrLf Or Right(phrase("sourcePhrase"), 1) <> vbLf Then
                phrase("sourcePhrase") = phrase("sourcePhrase") & vbCrLf
            End If
            .Range("B2") = .Range("B2") & phrase("sourcePhrase")
            row = row + 1
        Next
    End With
End Sub

Public Sub PasteParsedFrom(base As String, wrapped As String, tableName As String, condition As String)
        With ThisWorkbook.Worksheets("From")
            .Range("A1") = "���t���[�Y"
            .Range("A2") = base
            .Range("A3") = "From��i�܂�Ԃ��ς݁j"
            .Range("A4") = wrapped
            .Range("A5") = "�֘A�e�[�u����"
            .Range("A6") = tableName
            .Range("A7") = "Join����"
            .Range("A8") = condition
        End With
End Sub

Public Sub PasteWherePhrase(base As String, japanese As String)
        With ThisWorkbook.Worksheets("Where")
            .Range("A1") = "���t���[�Y"
            .Range("A2") = base
            .Range("A3") = "���{�ꉻ"
            .Range("A4") = japanese
        End With

End Sub



Public Sub PasteInsert(sqlLines As Variant)
    With ThisWorkbook.Worksheets("SQL��")
        .Cells(1, 1) = "SQL��"
        Dim lineNum As Long
        For lineNum = LBound(sqlLines) To UBound(sqlLines)
            Sheet1.Cells(lineNum + 2, 1) = sqlLines(lineNum)
        Next
    End With
End Sub
Public Sub PasteUpdate(sqlLines As Variant)
    With ThisWorkbook.Worksheets("SQL��")
        .Cells(1, 1) = "SQL��"
        Dim lineNum As Long
        For lineNum = LBound(sqlLines) To UBound(sqlLines)
            Sheet1.Cells(lineNum + 2, 1) = sqlLines(lineNum)
        Next
    End With
End Sub
Public Sub PasteSelectInto(sqlLines As Variant)
    With ThisWorkbook.Worksheets("SQL��")
        .Cells(1, 1) = "SQL��"
        Dim lineNum As Long
        For lineNum = LBound(sqlLines) To UBound(sqlLines)
            Sheet1.Cells(lineNum + 2, 1) = sqlLines(lineNum)
        Next
    End With
End Sub
Public Sub PasteDelete(sqlLines As Variant)
    With ThisWorkbook.Worksheets("SQL��")
        .Cells.ClearContents
        .Cells(1, 1) = "SQL��"
        Dim lineNum As Long
        For lineNum = LBound(sqlLines) To UBound(sqlLines)
            Sheet1.Cells(lineNum + 2, 1) = sqlLines(lineNum)
        Next
    End With
End Sub

