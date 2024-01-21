VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "b_FromParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private replaceTargets As Collection
Private translateTable As Object

Private Sub Class_Initialize()
    Set replaceTargets = New Collection
    replaceTargets.Add " LEFT JOIN "
    replaceTargets.Add " RIGHT JOIN "
    replaceTargets.Add " INNER JOIN "
    replaceTargets.Add " JOIN "
    replaceTargets.Add " ON "
    replaceTargets.Add " AND "
    replaceTargets.Add " OR "
    
    Set translateTable = CreateObject("Scripting.Dictionary")
    translateTable.Add " LEFT JOIN ", vbCrLf & "LEFTJOIN "
    translateTable.Add " RIGHT JOIN ", vbCrLf & "RIGHTJOIN "
    translateTable.Add " INNER JOIN ", vbCrLf & "INNERJOIN "
    translateTable.Add " JOIN ", vbCrLf & "JOIN "
    translateTable.Add " ON ", vbCrLf & "  ON "
    translateTable.Add " AND ", vbCrLf & "    AND "
    translateTable.Add " OR ", vbCrLf & "    OR "

End Sub

Public Function GetTableNames(phrase As String)
    Dim target As Variant
    For Each target In replaceTargets
        If InStr(target, "JOIN") > 0 Then
            phrase = Replace(phrase, target, translateTable(target))
        End If
    Next
    
    Dim rows As Variant: rows = Split(phrase, vbCrLf)
    Dim i As Long, firstSpacePos As Long, secondSpacePos As Long
    Dim tablename As String
    Dim tableNames As Collection: Set tableNames = New Collection
    For i = LBound(rows) To UBound(rows)
        If InStr(rows(i), "  ON ") > 0 Or InStr(rows(i), "    AND ") > 0 Or InStr(rows(i), "  OR ") > 0 Then GoTo CONTINUE
        firstSpacePos = InStr(rows(i), " ")
        secondSpacePos = InStr(firstSpacePos + 1, rows(i), " ")
        If secondSpacePos <> 0 Then
            tablename = Trim(Mid(rows(i), firstSpacePos + 1, secondSpacePos - firstSpacePos))
        Else
            tablename = Trim(Mid(rows(i), firstSpacePos))
        End If
        tableNames.Add tablename
CONTINUE:
    Next
    
    Dim name As Variant
    Dim concatTableName As Variant
    For Each name In tableNames
        concatTableName = concatTableName & ", " & name
    Next
    GetTableNames = Mid(concatTableName, 3)
End Function

Public Function GetWappedPhrase(phrase As String)
    Dim target As Variant
    For Each target In replaceTargets
            phrase = Replace(phrase, target, translateTable(target))
    Next
    GetWappedPhrase = phrase
End Function
