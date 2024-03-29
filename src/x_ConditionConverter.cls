VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "x_ConditionConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private comparisons As Collection
Private translateTable As Object

Private replaceTargets As Collection
Private translateTable2 As Object

Private Sub Class_Initialize()
    Set comparisons = New Collection
    comparisons.Add " = "
    comparisons.Add "<>"
    comparisons.Add " < "
    comparisons.Add " > "
    comparisons.Add "<="
    comparisons.Add ">="
    comparisons.Add " LIKE "
    comparisons.Add " BETWEEN "
    
    Set translateTable = CreateObject("Scripting.Dictionary")
    translateTable.Add " = ", Array(" と ", " が一致 ")
    translateTable.Add "<>", Array(" と ", "が一致しない ")
    translateTable.Add " < ", Array(" が ", " より小さい ")
    translateTable.Add " > ", Array(" が ", " より大きい ")
    translateTable.Add "<=", Array(" が ", " 以下 ")
    translateTable.Add ">=", Array(" が ", " 以上 ")
    translateTable.Add " LIKE ", Array(" に ", " が含まれる ")
    translateTable.Add " BETWEEN ", Array(" が ", " 〜 ")


    Set replaceTargets = New Collection
    replaceTargets.Add " AND "
    replaceTargets.Add " OR "
    replaceTargets.Add "."
    replaceTargets.Add "["
    replaceTargets.Add "]"
    replaceTargets.Add ";"
    replaceTargets.Add "("
    replaceTargets.Add ")"
    replaceTargets.Add "WHERE "
    replaceTargets.Add "Is Null"
    replaceTargets.Add "Is Not Null"
    
    Set translateTable2 = CreateObject("Scripting.Dictionary")
    translateTable2.Add " AND ", " かつ "
    translateTable2.Add " OR ", " 又は "
    translateTable2.Add ".", " の "
    translateTable2.Add "[", ""
    translateTable2.Add "]", ""
    translateTable2.Add "(", ""
    translateTable2.Add ")", ""
    translateTable2.Add "WHERE ", ""
    translateTable2.Add "Is Null", " が未入力 "
    translateTable2.Add "Is Not Null", " が未入力 "
    
    
End Sub

Public Function Replacecomparisons(wherePhrase As Variant) As Variant
    Dim operator As Variant
    Dim targetPos As Long, spacePos As Long, charNum As Long
    Dim leftPart As String, rightPart As String, restPart As String, newPhrase As String, newPhraseAll As String
    For Each operator In comparisons
        targetPos = InStr(wherePhrase, operator)
        If targetPos <> 0 Then newPhraseAll = ""
        Do While targetPos <> 0
            leftPart = Left(wherePhrase, targetPos - 1)
            charNum = Len(operator)
            wherePhrase = Trim(Mid(wherePhrase, targetPos + charNum))
            spacePos = InStr(wherePhrase, " ")
            If spacePos = 0 Then
                restPart = wherePhrase
                wherePhrase = leftPart & translateTable(operator)(0) & restPart & translateTable(operator)(1)
            Else
                rightPart = Left(wherePhrase, spacePos - 1)
                restPart = Mid(wherePhrase, spacePos + 1)
                wherePhrase = leftPart & translateTable(operator)(0) & rightPart & translateTable(operator)(1) & restPart
            End If
            targetPos = InStr(wherePhrase, operator)
        Loop
    Next
    
    Replacecomparisons = wherePhrase
End Function

Public Function SimpleReplace(phrase As Variant) As Variant
    Dim target As Variant
    For Each target In replaceTargets
        phrase = Replace(phrase, target, translateTable2(target), , , vbTextCompare)
    Next
    SimpleReplace = phrase
End Function
