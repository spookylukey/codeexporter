Attribute VB_Name = "modExcel"
' This modules contains low level code useful for any Excel VBA project

Function GetCustomProperty(wkbk As Workbook, pname As String) As Variant
    On Error GoTo DoesntExist
    GetCustomProperty = wkbk.CustomDocumentProperties(pname)
    Exit Function
    
DoesntExist:
    GetCustomProperty = Null
    
End Function
    
Sub SetCustomProperty(wkbk As Workbook, pname As String, ptype, val As Variant)
    On Error GoTo DoesntExist
    wkbk.CustomDocumentProperties(pname) = val
    Exit Sub
    
DoesntExist:
    On Error Resume Next
    wkbk.CustomDocumentProperties.Add pname, False, ptype, val
End Sub


