VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "x_Parser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
Private uf_txt_table As MSForms.TextBox
Private uf_txt_about As MSForms.TextBox

Private selectParser As a_SelectParser
Private fromParser As b_FromParser
Private converter As x_ConditionConverter
Private returns As Object
Private thisWb As xx_ThisWorkBook

Private sql_ As Variant
Private ColumnPhrases As Collection
Private singleTable As Boolean



Private Sub Class_Initialize()
    Set selectParser = New a_SelectParser
    Set fromParser = New b_FromParser
    Set converter = New x_ConditionConverter
    Set returns = CreateObject("Scripting.Dictionary")
    Set thisWb = New xx_ThisWorkBook
End Sub


Public Function SelectSQL(sql() As String) As Object
    thisWb.PasteSelect sql
    
    ParseSelect sql(0)
    ParseFrom sql(1)
    
    Dim about As String
    If singleTable Then
        about = "関連テーブルより" & vbCrLf
    Else
        about = "関連テーブル記載の結合テーブルより" & vbCrLf
    End If
    MakeAboutText sql, about
    Set SelectSQL = returns
End Function

' Select句解析 & selectシートへ解析内容貼り付け
Private Sub ParseSelect(phrase As String)
    Set ColumnPhrases = selectParser.GetParsedColumnPhrases(phrase)
    thisWb.PasteParsedSelect phrase, ColumnPhrases
End Sub

' From句解析 ＆ Fromシートへ解析内容貼り付け
Private Sub ParseFrom(phrase As String)
    Dim wrappedFrom As String: wrappedFrom = fromParser.GetWappedPhrase(phrase)
    Dim tableName As String: tableName = fromParser.GetTableNames(phrase)
    
    Dim for_uf_txt_table As String
    Dim condition As String
    If InStr(tableName, ", ") = 0 Then
        for_uf_txt_table = tableName
        singleTable = True
    Else
        for_uf_txt_table = tableName & vbCrLf & "以下で結合" & vbCrLf
        condition = GetJoinCondition(wrappedFrom)
        for_uf_txt_table = for_uf_txt_table & condition
        singleTable = False
    End If
    
    returns("関連テーブル") = for_uf_txt_table
    thisWb.PasteParsedFrom phrase, wrappedFrom, tableName, condition

End Sub

Private Function GetJoinCondition(wrappedFrom As String) As String
    Dim rows As Variant: rows = Split(wrappedFrom, vbCrLf)
    
    Dim conditionPhrase As String
    Dim i As Long
    For i = LBound(rows) To UBound(rows)
        If InStr(rows(i), "  ON ") > 0 Or InStr(rows(i), "    AND ") > 0 Or InStr(rows(i), "  OR ") > 0 Then
            rows(i) = converter.Replacecomparisons(rows(i))
            rows(i) = converter.SimpleReplace(rows(i))
            conditionPhrase = conditionPhrase & rows(i) & vbCrLf
        End If
    Next
    conditionPhrase = Replace(conditionPhrase, "  ON", "・")
    conditionPhrase = Trim(Left(conditionPhrase, Len(conditionPhrase) - 1))
    GetJoinCondition = conditionPhrase
End Function

Private Sub MakeAboutText(sql() As String, about As String)

    Dim num As Long, i As Long
    Dim orderby As String
    For i = LBound(sql) To UBound(sql)
        If InStr(sql(i), "WHERE") > 0 Then
            num = num + 1
            about = about & num & GetWherePhrase(sql(i))
            
        ElseIf InStr(sql(i), "GROUP BY ") > 0 Then
            num = num + 1
            about = about & num & GetGroupByPhrase(Replace(sql(i), "GROUP BY ", ""))
            
        ElseIf InStr(sql(i), "HAVING ") > 0 Then
            num = num + 1
            about = about & num & GetHavingPhrase(sql(i))
            
        ElseIf InStr(sql(i), "ORDER BY ") > 0 Then
            orderby = GetOrderByPhrase(sql(i))
        End If
    Next
    
    If Not ColumnPhrases Is Nothing Then
        Dim aggregatePhrase As String
        aggregatePhrase = GetAggregatePhrase
        If aggregatePhrase <> "" Then
            num = num + 1
            about = about & num & ". 各グループで以下集計を実施" & vbCrLf & " ・" & aggregatePhrase & vbCrLf & vbCrLf
        End If
    End If
    
    If orderby <> "" Then
        num = num + 1
        about = about & num & orderby
    End If
    
    returns("概要") = about
End Sub

Private Function GetWherePhrase(phrase As String) As String
    Dim condition As String
    condition = GetConditionPhrase(phrase)
    GetWherePhrase = ". 以下条件を満たすレコードを抽出" & vbCrLf & " ・" & condition & vbCrLf & vbCrLf
    
    thisWb.PasteWherePhrase phrase, condition
End Function

Private Function GetConditionPhrase(phrase As String) As String
    Dim newPhrase As String
    newPhrase = converter.Replacecomparisons(phrase)
    newPhrase = converter.SimpleReplace(newPhrase)
    GetConditionPhrase = newPhrase
End Function

Private Function GetGroupByPhrase(phrase As String) As String
    Dim cols As Collection
    Set cols = selectParser.GetParsedColumnPhrases(phrase)
    
    Dim s As String
    Dim i As Long
    For i = 1 To cols.count
        If singleTable Then
            s = s & cols(i)("columnName") & ", "
        Else
            s = s & cols(i)("tableName") & " の " & cols(i)("columnName") & ", "
        End If
    Next
    s = Replace(s, ";", "")
    GetGroupByPhrase = ". 以下列の値でグループ化" & vbCrLf & "  " & s & vbCrLf & vbCrLf
End Function

Private Function GetHavingPhrase(phrase As String) As String
    Dim condition As String
    condition = GetConditionPhrase(phrase)
    GetHavingPhrase = ". 以下条件を満たすグループを抽出" & vbCrLf & " ・" & condition & vbCrLf & vbCrLf
End Function

Private Function GetAggregatePhrase() As String
    Dim i As Long
    Dim colName As String
    Dim phrase As String
    For i = 1 To ColumnPhrases.count
        If ColumnPhrases(i)("aggregate")(0) <> "" Then
            colName = Replace(ColumnPhrases(i)("columnName"), ColumnPhrases(i)("aggregate")(0), "")
            colName = Mid(colName, 2, Len(colName) - 2)
            GetAggregatePhrase = GetAggregatePhrase & colName & "：" & ColumnPhrases(i)("aggregate")(1) & ", "
        End If
    Next
End Function

Private Function GetOrderByPhrase(phrase As String) As String
    Dim orderby As String
    orderby = Replace(phrase, "ORDER BY ", "")
    orderby = Replace(orderby, ";", "")
    
    If InStr(orderby, "DESC") > 0 Then
        orderby = Replace(orderby, "DESC", "")
        GetOrderByPhrase = ". 列:" & orderby & " で　降順並べ替え"
    Else
        GetOrderByPhrase = ". 列:" & orderby & " で　昇順並べ替え"
    End If
End Function

'SelectPaserに記述し、ColumnPhrasesにaggregateを持たせる
'Private Function CheckAggregateFunc(phrase As String) As String()
'    Dim aggregate(1) As String
'    If InStr(phrase, "SUM") > 0 Then
'        aggregate(0) = "SUM"
'        aggregate(1) = "合計"
'        CheckAggregateFunc = aggregate
'    ElseIf InStr(phrase, "COUNT") > 0 Then
'        aggregate(0) = "COUNT"
'        aggregate(1) = "カウント"
'        CheckAggregateFunc = aggregate
'    ElseIf InStr(phrase, "AVG") > 0 Then
'        aggregate(0) = "AVG"
'        aggregate(1) = "平均"
'        CheckAggregateFunc = aggregate
'    ElseIf InStr(phrase, "MIN") > 0 Then
'        aggregate(0) = "MIN"
'        aggregate(1) = "最小"
'        CheckAggregateFunc = aggregate
'    ElseIf InStr(phrase, "MAX") > 0 Then
'        aggregate(0) = "MAX"
'        aggregate(1) = "最大"
'        CheckAggregateFunc = aggregate
'    Else
'        aggregate(0) = ""
'        aggregate(1) = ""
'        CheckAggregateFunc = aggregate
'    End If
'
'End Function



Public Function InsertSQL(sql() As String)
    
    thisWb.PasteInsert sql
    
    ParseSelect sql(1)
    ParseFrom sql(2)
    
    Dim about As String
    about = "データ追加対象テーブル：" & vbCrLf
    If singleTable Then
        about = about & "データ取得元テーブル：" & returns("関連テーブル") & vbCrLf & vbCrLf
    Else
        about = about & "データ取得元テーブル：" & "関連テーブル記載の結合テーブル" & vbCrLf & vbCrLf
    End If
    about = about & "データ取得元テーブルより" & vbCrLf
    
    MakeAboutText sql, about
    returns("概要") = returns("概要") & "結果をデータ追加対象テーブルへ追加する"
    Set InsertSQL = returns
End Function



Public Function UpdateSQL(sql() As String)
    thisWb.PasteUpdate sql
    
    
'    ParseFrom sql(2)
'    MakeUpdateAboutText sql
    Set UpdateSQL = returns
End Function

Public Sub MakeUpdateAboutText(sql() As String, about As String)

    returns("概要") = about
End Sub

Public Function SelectIntoSQL(sql() As String)
    thisWb.PasteSelectInto sql
    SelectSQL sql 'select解析をそのまま代用
    
    Dim into As String: into = Split(sql(0), "INTO ")(1)
    returns("概要") = "以下処理の結果を取得し、新テーブル「" & into & "」を作成" & vbCrLf & vbCrLf & returns("概要")
    Set SelectIntoSQL = returns
End Function

Public Function DeleteSQL(sql() As String)
    thisWb.PasteDelete sql
    
    Dim about As String
    about = ""
    ParseFrom sql(1)
    MakeAboutText sql, about
    Set DeleteSQL = returns
End Function
