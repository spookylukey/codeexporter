Attribute VB_Name = "modExport"
Option Explicit

' Export all (or selected) modules in a project to
' a directory

' called from menus/toolbars
Public Sub ExportCode()
    ExportWorkbookCode ActiveWorkbook
End Sub

' use this directly if you need to save an add-in
' (an add-in can never be the ActiveWorkbook)
Sub ExportWorkbookCode(wkbk As Workbook)
    Dim dumpdir As String, defdir As Variant
    
    If wkbk.VBProject.Protection = vbext_pp_locked Then
        MsgBox "You must unlock the project before it can be exported.", , "CodeExporter"
        wkbk.VBProject.VBE.MainWindow.Visible = True
        Exit Sub
    End If
    
    ' if the workbook has a custom property 'CodeExporterSavePath' then use that
    defdir = GetCustomProperty(wkbk, "CodeExporterSavePath")
    If IsNull(defdir) Then
        defdir = wkbk.Path
    End If
    If Left$(defdir, 1) = "\" Then
        ' relative dir - add workbook path
        defdir = ThisWorkbook.Path & defdir
    End If
    dumpdir = GetDirectory(CStr(defdir), 0, True, "Select folder to export modules to:")
    If dumpdir <> "" Then
        ExportModules dumpdir, wkbk
        If Not wkbk.ReadOnly Then
            ' save the directory, but only if it is different from
            ' what is was before.  Also don't save it if it is the
            ' default
            If CStr(defdir) <> dumpdir Then
                ' see if we can set a relative path
                Debug.Print dumpdir, wkbk.Path, InStr(1, dumpdir, wkbk.Path)
                If InStr(1, dumpdir, wkbk.Path) = 1 Then
                    dumpdir = VBA.Mid(dumpdir, Len(wkbk.Path) + 1)
                End If
                SetCustomProperty wkbk, "CodeExporterSavePath", msoPropertyTypeString, dumpdir
            End If
        End If
    End If
End Sub

Public Sub ExportModules(dirname As String, wkbk As Workbook)
    Dim VBComp As VBComponent
    Dim fname As String
    
    For Each VBComp In wkbk.VBProject.VBComponents
        fname = VBComp.Name
        If VBComp.Type = vbext_ct_ClassModule Then
            fname = fname & ".cls"
        ElseIf VBComp.Type = vbext_ct_Document Then
            fname = fname & ".cls"
        ElseIf VBComp.Type = vbext_ct_MSForm Then
            fname = fname & ".frm"
        ElseIf VBComp.Type = vbext_ct_StdModule Then
            fname = fname & ".bas"
        End If
        ' remove it first
'        If Dir(CStr(dirname & "\" & fname), vbNormal) <> "" Then
'            On Error GoTo CantDelete
'            Kill dirname & "\" & fname
'        End If
        VBComp.Export dirname & "\" & fname
'        GoTo Skip1
'CantDelete:
'        MsgBox "Couldn't delete the existing file: " & VBA.Chr(10) & fname
'Skip1:
    Next
End Sub

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

' To automate exporting of code, use a routine like this.
' This requires codeexporter to be open or installed as an add-in
Sub ExportMe()
    Run "codeexporter.xla!ExportWorkbookCode", ThisWorkbook
End Sub
