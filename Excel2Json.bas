Attribute VB_Name = "Excel2Json"
Option Explicit

' ヘッダ行を収集
Private Function GetHeaders(ByVal sheet As Object, ByVal row As Integer, ByVal col As Integer) As Collection
    Dim maxCol As Integer, headers As Collection, i As Integer
    Set headers = New Collection
    For i = col To sheet.UsedRange.Columns.Count
        Dim val As Variant
        val = sheet.Cells(row, i)
        If val = "" Then Exit For
        headers.Add val
    Next
    Set GetHeaders = headers
End Function

' 再帰的に辞書要素を追加
Private Sub AddDictItem(ByVal parent As Object, ByVal keys As Collection, ByVal value As Variant)
    If keys.Count = 1 Then
        parent.Add keys(1), value
    Else
        Dim key As String
        key = keys(1)
        keys.Remove 1
        If parent.Exists(key) Then
            AddDictItem parent(key), keys, value
        Else
            Dim child As Object
            Set child = CreateObject("Scripting.Dictionary")
            AddDictItem child, keys, value
            parent.Add key, child
        End If
    End If
End Sub

Private Function SplitAddress(ref As String)
    SplitAddress = Split(Mid(ref, 2, Len(ref) - 2), "!")
End Function

Private Function GetReferencedValue(ByVal sheet As Worksheet, ref As String)
    Dim adr, row As Integer, col As Integer
    adr = SplitAddress(ref)
    If UBound(adr) >= 1 Then
        If adr(0) <> "" Then
            Set sheet = Worksheets(adr(0))
        End If
        If adr(1) <> "" Then
            Dim r As Object
            Set r = sheet.Range(adr(1))
            row = r.row: col = r.Column
        Else
            row = 1: col = 1
        End If
    Else
        Set sheet = Worksheets(adr(0))
        row = 1: col = 1
   End If
    Set GetReferencedValue = GetDictArray(sheet, row, col)
End Function

Private Function GetIndex(ByVal ref As String)
    Dim adr
    adr = SplitAddress(ref)
    If UBound(adr) >= 2 Then
        GetIndex = adr(2)
    Else
        GetIndex = "1"
    End If
End Function

Private Function GetValue(ByVal sheet As Worksheet, ByVal val As String) As Variant
    If Left(val, 1) = "[" And Right(val, 1) = "]" Then
        Set GetValue = GetReferencedValue(sheet, val)
    ElseIf Left(val, 1) = "{" And Right(val, 1) = "}" Then
        Set GetValue = GetReferencedValue(sheet, val)(GetIndex(val))
    Else
        GetValue = val
    End If
End Function

' １行分の辞書データを取得
Private Function GetDict(ByVal sheet As Worksheet, ByVal row As Integer, ByVal col As Integer, ByVal headers As Collection) As Object
    Dim result As Object, hasValue As Boolean, head As Variant, i As Integer
    Set result = CreateObject("Scripting.Dictionary")
    hasValue = False
    i = 0
    For Each head In headers
        Dim val As Variant, keys As Collection, key As Variant
        Set keys = New Collection
        For Each key In Split(headers(i + 1), ".")
            keys.Add key
        Next
        val = sheet.Cells(row, col + i)
        AddDictItem result, keys, GetValue(sheet, val)
        If val <> "" Then hasValue = True
        i = i + 1
    Next
    If hasValue Then
        Set GetDict = result
    Else
        Set GetDict = Nothing
    End If
End Function

' 辞書の配列を取得
Private Function GetDictArray(ByVal sheet As Worksheet, ByVal row As Integer, ByVal col As Integer) As Object
    Dim headers As Collection, result As Object, values As Object, i As Integer
    Set result = CreateObject("Scripting.Dictionary")
    Set headers = GetHeaders(sheet, row, col)
    row = row + 1
    i = 1
    Set values = GetDict(sheet, row, col, headers)
    Do Until values Is Nothing
        result.Add CStr(i), values
        row = row + 1
        i = i + 1
        Set values = GetDict(sheet, row, col, headers)
    Loop
    Set GetDictArray = result
End Function

Private Function JoinStringCollection(ByVal col As Collection) As String
    Dim result() As String, i As Integer
    ReDim result(col.Count - 1)
    For i = 0 To col.Count - 1
        result(i) = col.item(i + 1)
    Next
    JoinStringCollection = Join(result, ", ")
End Function

Private Function IsNumericArray(ByVal obj As Variant) As Boolean
    If IsObject(obj) Then
        IsNumericArray = True
        Dim key As Variant
        For Each key In obj
            If Not IsNumeric(key) Then
                IsNumericArray = False
                Exit For
            End If
        Next
    Else
        IsNumericArray = False
    End If
End Function

' 辞書のキーがすべて数値であれば，インデックスの最大値を返す。
Private Function GetMaxIndex(ByVal obj As Variant) As Integer
    If IsObject(obj) Then
        Dim maxValue As Integer, key As Variant
         maxValue = 0
        For Each key In obj
            If Not IsNumeric(key) Then
                maxValue = 0
                Exit For
            End If
            If key > maxValue Then maxValue = key
        Next
        GetMaxIndex = maxValue
    Else
        GetMaxIndex = 0
    End If
End Function

Private Function ToString(ByVal obj As Variant) As String
    Dim items As Object, item As Variant, key As Variant
    Dim maxIndex As Integer, i As Integer
    Set items = New Collection
    maxIndex = GetMaxIndex(obj)
    If maxIndex > 0 Then
        For i = 1 To maxIndex
            key = CStr(i)
            If obj.Exists(key) Then
                items.Add ToString(obj.item(key))
            Else
                items.Add "null"
            End If
        Next
        ToString = "[" & JoinStringCollection(items) & "]"
    ElseIf IsObject(obj) Then
        For Each key In obj
            items.Add ("""" & key & """: " & ToString(obj.item(key)))
        Next
        ToString = "{" & JoinStringCollection(items) & "}"
    ElseIf obj = Empty Then
        ToString = "null"
    Else
        ToString = """" & CStr(obj) & """"
    End If
End Function

' アクティブシートのA1を起点として表をJSON形式に変換する。
Public Function Convert() As String
    Dim result As Object
    Set result = GetDictArray(ActiveSheet, 1, 1)
    Convert = ToString(result.item("1"))
End Function


