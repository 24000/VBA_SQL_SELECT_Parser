VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xxx_TargetWB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum wsSelectCol
    sql = 4
    Table
    about
End Enum

Private Enum wsActionCol
    actiontype = 4
    sql
    Table
    about
End Enum

Private wb_ As Workbook
Private ws_ As Worksheet
Private targetRow_ As Long

Public Sub Init(targetWbNum As Long, targetWsName As String, targetRow As Long)
    For Each wb_ In Workbooks
        If InStr(wb_.name, "【" & targetWbNum & "】") > 0 Then Exit For
    Next
    Set ws_ = wb_.Worksheets(targetWsName)
    targetRow_ = targetRow
End Sub


Public Function GetSQL() As String
    Dim sql As String
    If ws_.name = "選択クエリ" Then
        sql = ws_.Cells(targetRow_, wsSelectCol.sql)
    Else
        sql = ws_.Cells(targetRow_, wsActionCol.sql)
    End If
    If sql = "" Then
        MsgBox "指定のシート、行にSQL文がありません"
        GetSQL = sql
    Else
        GetSQL = sql
    End If
End Function

Public Sub PasteTable(s As String)
    With ws_
        If .name = "選択クエリ" Then
            .Cells(targetRow_, wsSelectCol.Table) = s
            .Cells(targetRow_, wsSelectCol.Table).WrapText = False
        ElseIf .name = "アクションクエリ" Then
            .Cells(targetRow_, wsActionCol.Table) = s
            .Cells(targetRow_, wsActionCol.Table).WrapText = False
        End If
    End With
End Sub

Public Sub PasteAbout(s As String)
    With ws_
        If .name = "選択クエリ" Then
            .Cells(targetRow_, wsSelectCol.about) = s
            .Cells(targetRow_, wsSelectCol.about).WrapText = False
        ElseIf .name = "アクションクエリ" Then
            .Cells(targetRow_, wsActionCol.about) = s
            .Cells(targetRow_, wsActionCol.about).WrapText = False
        End If
    End With
End Sub

Public Function GetQueryType() As String
    With ws_
        If .name = "選択" Then
            GetQueryType = "選択クエリ"
        ElseIf .name = "アクションクエリ" Then
            GetQueryType = .Cells(targetRow_, wsActionCol.actiontype)
        End If
    End With
End Function
