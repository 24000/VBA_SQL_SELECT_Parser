Attribute VB_Name = "Module1"
Option Explicit


Sub selectãÂâêÕ()
    Dim selectPhrase As String
    selectPhrase = Sheet1.Range("A1")
    
    Dim parser As SelectParser: Set parser = New SelectParser
    Dim words As Collection
    Set words = parser.GetParsedPhrases(selectPhrase)
    
    Sheet2.Cells.ClearContents
    Dim word As Variant
    Dim row As Long: row = 2
    For Each word In words
        Sheet2.Cells(row, 1) = word("displayName")
        Sheet2.Cells(row, 2) = word("tableName")
        Sheet2.Cells(row, 3) = word("columnName")
        row = row + 1
    Next
End Sub


