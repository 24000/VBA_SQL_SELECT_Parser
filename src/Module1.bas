Attribute VB_Name = "Module1"
Option Explicit

Sub selectãÂâêÕ()
    Dim selectPhrase As String
    selectPhrase = Sheet1.Range("A1")
    
    Dim parser As SelectParser: Set parser = New SelectParser
    Dim ColumnPhrases As Collection
    Set ColumnPhrases = parser.GetParsedColumnPhrases(selectPhrase)
    
    Sheet1.Range("A2:D100").ClearContents
    Dim phrase As Variant
    Dim row As Long: row = 4
    For Each phrase In ColumnPhrases
        Sheet1.Cells(row, 2) = phrase("displayName")
        Sheet1.Cells(row, 3) = phrase("tableName")
        Sheet1.Cells(row, 4) = phrase("columnName")
        If Right(phrase("sourcePhrase"), 1) <> vbCrLf Or Right(phrase("sourcePhrase"), 1) <> vbLf Then
            phrase("sourcePhrase") = phrase("sourcePhrase") & vbCrLf
        End If
        Sheet1.Range("A2") = Sheet1.Range("A2") & phrase("sourcePhrase")
        row = row + 1
    Next
End Sub


Sub èåèÇì˙ñ{åÍïœä∑()
    
    Dim s As String
    s = Selection.Value
    
    Dim converter As x_ConditionConverter: Set converter = New x_ConditionConverter
    Dim newPhrase As String
    newPhrase = converter.Replacecomparisons(s)
    newPhrase = converter.SimpleReplace(newPhrase)
    Sheet2.Range("B2") = newPhrase
        
End Sub

Sub FromãÂâêÕ()
    
    Dim s As String
    s = Selection.Value
    
    Dim parser As b_FromParser: Set parser = New b_FromParser
    
    Dim wrappedPhrase As String
    wrappedPhrase = parser.GetWappedPhrase(s)
    Sheet3.Range("B2") = wrappedPhrase
    
    Dim tableNames As String
    tableNames = parser.GetTableNames(s)
    Sheet3.Range("B3") = tableNames
    
    Dim rows As Variant: rows = Split(Sheet3.Range("B2").Value, vbCrLf)
    Dim converter As x_ConditionConverter: Set converter = New x_ConditionConverter
    Dim newPhrase As String
    Dim i As Long
    For i = LBound(rows) To UBound(rows)
        If InStr(rows(i), "  ON ") > 0 Or InStr(rows(i), "    AND ") > 0 Or InStr(rows(i), "  OR ") > 0 Then
            rows(i) = converter.Replacecomparisons(rows(i))
            rows(i) = converter.SimpleReplace(rows(i))
            newPhrase = newPhrase & rows(i) & vbCrLf
        End If
    Next
    newPhrase = Replace(newPhrase, "  ON", "ÅE")
    
    If Len(newPhrase) <> 0 Then
        Sheet3.Range("B4") = Trim(Left(newPhrase, Len(newPhrase) - 1))
    End If
End Sub



