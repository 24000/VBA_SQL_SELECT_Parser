VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function GetParsedPhrases(selectPhrase As String) As Object
    Dim columnPhrases As Collection
    Set columnPhrases = DivideSelectPhraseToColumnPhrases(selectPhrase)
    
    Dim parsedPhrases As Collection: Set parsedPhrases = New Collection
    Dim ColumnPhrase As Variant
    Dim parsedphrase As Object
    For Each ColumnPhrase In columnPhrases
        Set parsedphrase = Parse(ColumnPhrase)
        parsedPhrases.Add parsedphrase
    Next
    
    Set GetParsedPhrases = parsedPhrases
End Function


Private Function DivideSelectPhraseToColumnPhrases(selectPhrase As String) As Object
    selectPhrase = Replace(selectPhrase, "SELECT", "")
    
    Dim words As Collection: Set words = New Collection
    
    Dim comma_separate As Variant
    comma_separate = Split(selectPhrase, ",")
    
    
    Dim word As Variant
    Dim funcMode As Boolean
    Dim lParenthesisCount As Long, rParenthesisCount As Long
    Dim funcStr As String
    For Each word In comma_separate
        If funcMode Then  ' 関数文字列解析中の場合
            If InStr(word, "(") > 0 Or InStr(word, ")") > 0 Then  '関数ネストか関数終了文字列かを判定
                lParenthesisCount = lParenthesisCount + CharCount(word, "(")
                rParenthesisCount = rParenthesisCount + CharCount(word, ")")
                
                If lParenthesisCount = rParenthesisCount Then  '　関数の文字列の最終か確認
                    funcStr = funcStr & word
                    words.Add funcStr
                    funcStr = ""
                    funcMode = False
                Else
                    funcStr = funcStr & word & ","
                End If
            Else  '関数の引数
                funcStr = funcStr & word & ","
            End If
            
        ElseIf InStr(word, "(") > 0 Then  ' 関数文字列解析中ではない場合。まず関数開始文字列かチェック。
            lParenthesisCount = CharCount(word, "(")
            rParenthesisCount = CharCount(word, ")")
            If lParenthesisCount = rParenthesisCount Then    ' ()の数が同じ関数なら問題なし
                words.Add word
            Else  '異なるのであれば、関数解析モードに設定し、"(" と ")"のカウント
                funcMode = True
                funcStr = word & ","
            End If
        Else
             words.Add word
        End If
    
    Next
    
    Set DivideSelectPhraseToColumnPhrases = words
End Function

Private Function CharCount(word As Variant, targetChar As String) As Long
    CharCount = Len(word) - (Len(Replace(word, targetChar, "")))
End Function


Private Function Parse(phrase As Variant) As Object
    
    Dim asDivided As Variant, periodDevided As Variant
    Dim displayName As String, columnName As String, tableName As String
    
    
    If InStr(1, phrase, " AS ", vbTextCompare) > 0 Then
        asDivided = Split(phrase, " AS ", compare:=vbTextCompare)
        displayName = asDivided(1)
        columnName = asDivided(0)
    Else
        columnName = phrase
    End If
    
    If InStr(columnName, "(") > 0 Then '関数の場合：テーブル名は空欄にして解析者に判断させる
        tableName = ""
    ElseIf InStr(columnName, ".") > 0 Then
        periodDevided = Split(columnName, ".")
        columnName = periodDevided(1)
        tableName = periodDevided(0)
    Else
        tableName = ""
    End If
    
    If displayName = "" Then displayName = columnName
    
    Dim parsedphrase As Object: Set parsedphrase = CreateObject("Scripting.Dictionary")
    parsedphrase.Add "displayName", TrimAndRemoveBrackets(displayName)
    parsedphrase.Add "columnName", TrimAndRemoveBrackets(columnName)
    parsedphrase.Add "tableName", TrimAndRemoveBrackets(tableName)
    
    Set Parse = parsedphrase
End Function

Private Function TrimAndRemoveBrackets(word As String) As String
    word = Replace(word, "[", "")
    word = Replace(word, "]", "")
    word = Trim(word)
    TrimAndRemoveBrackets = word
End Function
