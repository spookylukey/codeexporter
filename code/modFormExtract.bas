Attribute VB_Name = "modFormExtract"
Sub Test1()
    Dim frm As UserForm
    Dim comp As VBComponent
    Dim wkbk As Workbook
    Dim i As Integer
    Dim valstring As String
    Dim cntls As Controls
    Set wkbk = ThisWorkbook
    wkbk.Activate
    For Each comp In wkbk.VBProject.VBComponents
        If comp.Type = vbext_ct_MSForm Then
            'Set frm = wkbk.VBProject.UserForms
            'Debug.Print comp.Name
            'Load comp.Name
            Debug.Print "<UserForm Name=""" & comp.Name & """>"
            wkbk.VBProject.VBE.MainWindow.Visible = True
            comp.Activate
            Debug.Print "  <Name>" & comp.Name & "</Name>"
            For i = 1 To comp.Properties.Count
                On Error GoTo NoValue
                If VarType(comp.Properties(i).Value) <> vbObject Then
                    valstring = CStr(comp.Properties(i).Value)
'                    Debug.Print "  <Property Name=""" & comp.Properties(i).Name & """>" & valstring & "</Property>"
                Else
                    If comp.Properties(i).Name = "Controls" Then
                        Debug.Print "  <Controls>"
                        On Error GoTo 0
                        If TypeOf comp.Properties(i).Object Is Controls Then
                            printcontrols comp.Properties(i).Object
                        End If
                        Debug.Print "  </Controls>"
                    Else
                        If comp.Properties(i).Object Is Nothing Then
                            Debug.Print "  <Property Name=""" & comp.Properties(i).Name & """>[Nothing]</Property>"
                        Else
                            Debug.Print "  <Property Name=""" & comp.Properties(i).Name & """>[Object]</Property>"
                        End If
                    End If
                End If
                GoTo Skip1
NoValue:
                valstring = "[No value]"
'                Debug.Print "  <Property Name=""" & comp.Properties(i).Name & """>" & valstring & "</Property>"
Skip1:
                
            Next i
            Debug.Print "</UserForm>"
        End If
    Next
End Sub

Sub printcontrols(ByRef cntls As Controls)
    Dim cntl As Control
    Dim obj As OLEObject
    Dim i As Integer, j As Integer
    Dim props As Properties
    For i = 0 To cntls.Count - 1
        Set cntl = cntls(i)
        Debug.Print "    <Control>"
        Debug.Print "      <Name>" & cntl.Name & "</Name>"
        cntl.SetFocus

        Debug.Print Application.VBE.SelectedVBComponent.Name
        Set wnd = Application.VBE.SelectedVBComponent.Designer
        
        Debug.Print "    </Control>"
    Next
End Sub


