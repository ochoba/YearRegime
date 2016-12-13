Attribute VB_Name = "Il_funCheakID"
'Function Il_funCheakID(Il_strID)
'    Dim Il_arrPart() As String
'    Dim Il_strSlash As String
'    Dim Il_strErr As String
'
'    Dim Il_strPartHead As String
'    Dim Il_strPartTail As String
'
'    If Il_strID <> "" Then
'        'Разрез кода проверка
'        Il_arrPart = Split(Il_strID, "/")
'        Il_strSlash = UBound(Il_arrPart)
'
'        If Il_strSlash = 0 Then
'            Il_strErr = "нет символа ""/"" "
'
'        Else 'Если есть разделитель
'
'            'Проверка ГОСБ
'            Il_strPartHead = Il_arrPart(0)
'            If Len(Il_strPartHead) < 4 Then
'                Il_strErr = Il_strErr & "номер ГОСБ меньше 4 символов "
'            End If
'
'            'Проверка ВСП
'            Il_intPartTail = Il_arrPart(1)
'
'        End If
'    Else
'        Il_strErr = "Пусто ВСП"
'    End If
'        'установка ошибки
'        If Il_strErr <> "" Then
'            Range("Il_errCODE") = "Номер ВСП, ошибки: " & Il_strErr & "! "
'        Else
'            Range("Il_errCODE") = ""
'        End If
'        'MsgBox "сработало"
'End Function

