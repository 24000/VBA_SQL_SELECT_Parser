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
    With ThisWorkbook.Worksheets("SQL文")
        .Cells.ClearContents
        .Cells(1, 1) = "SQL文"
        Dim lineNum As Long
        For lineNum = LBound(sqlLines) To UBound(sqlLines)
            Sheet1.Cells(lineNum + 2, 1) = sqlLines(lineNum)
        Next
    End With
End Sub

Private Sub RowHeightAdjust()
    
End Sub

'SelectParserのGetParsedColumnPhrasesのシート貼り付け部を持ってきて、少し改変
Public Sub PasteParsedSelect(basePhrase As String, ColumnPhrases As Collection)
    With ThisWorkbook.Worksheets("Select")
        .Cells(1, 1) = "Select句"
        .Cells(2, 1) = basePhrase
        .Cells(1, 2) = "列名のみ改行表示"
        .Cells(3, 2) = "表示名"
        .Cells(3, 3) = "テーブル名"
        .Cells(3, 4) = "元列名"
        .Cells(3, 6) = "集計"
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
            .Range("A1") = "元フレーズ"
            .Range("A2") = base
            .Range("A3") = "From句（折り返し済み）"
            .Range("A4") = wrapped
            .Range("A5") = "関連テーブル名"
            .Range("A6") = tableName
            .Range("A7") = "Join条件"
            .Range("A8") = condition
        End With
End Sub

Public Sub PasteWherePhrase(base As String, japanese As String)
        With ThisWorkbook.Worksheets("Where")
            .Range("A1") = "元フレーズ"
            .Range("A2") = base
            .Range("A3") = "日本語化"
            .Range("A4") = japanese
        End With

End Sub



Public Sub PasteInsert(sqlLines As Variant)
    With ThisWorkbook.Worksheets("SQL文")
        .Cells(1, 1) = "SQL文"
        Dim lineNum As Long
        For lineNum = LBound(sqlLines) To UBound(sqlLines)
            Sheet1.Cells(lineNum + 2, 1) = sqlLines(lineNum)
        Next
    End With
End Sub
Public Sub PasteUpdate(sqlLines As Variant)
    With ThisWorkbook.Worksheets("SQL文")
        .Cells(1, 1) = "SQL文"
        Dim lineNum As Long
        For lineNum = LBound(sqlLines) To UBound(sqlLines)
            Sheet1.Cells(lineNum + 2, 1) = sqlLines(lineNum)
        Next
    End With
End Sub
Public Sub PasteSelectInto(sqlLines As Variant)
    With ThisWorkbook.Worksheets("SQL文")
        .Cells(1, 1) = "SQL文"
        Dim lineNum As Long
        For lineNum = LBound(sqlLines) To UBound(sqlLines)
            Sheet1.Cells(lineNum + 2, 1) = sqlLines(lineNum)
        Next
    End With
End Sub
Public Sub PasteDelete(sqlLines As Variant)
    With ThisWorkbook.Worksheets("SQL文")
        .Cells.ClearContents
        .Cells(1, 1) = "SQL文"
        Dim lineNum As Long
        For lineNum = LBound(sqlLines) To UBound(sqlLines)
            Sheet1.Cells(lineNum + 2, 1) = sqlLines(lineNum)
        Next
    End With
End Sub


