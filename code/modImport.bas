Attribute VB_Name = "modImport"
' Do the reverse of modExport
' It is not anticipated that this will be needed normally.
' It can be useful for restoring just the code to a workbook
' or starting a workbook from 'scratch' and adding all the code.

Option Explicit

' called from menus/toolbars
Public Sub CodeExporter_Import()
    CodeExporter_ImportCodeToWorkbook ActiveWorkbook
End Sub

Public Sub CodeExporter_ImportCodeToWorkbook(wkbk As Workbook)
    Dim dumpdir As String, defdir As Variant
    
    If wkbk.VBProject.Protection = vbext_pp_locked Then
        MsgBox "You must unlock the project before code can be imported.", , "CodeExporter"
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
        defdir = wkbk.Path & defdir
    End If
    dumpdir = GetDirectory(CStr(defdir), 0, True, "Select folder to import modules from:")
    If dumpdir <> "" Then
        ImportModules dumpdir, wkbk
        
        If Not wkbk.ReadOnly Then
            ' save the directory, but only if it is different from
            ' what is was before.  Also don't save it if it is the
            ' default
            If CStr(defdir) <> dumpdir Then
                ' see if we can set a relative path
                If InStr(1, dumpdir, wkbk.Path) = 1 Then
                    dumpdir = VBA.Mid(dumpdir, Len(wkbk.Path) + 1)
                End If
                SetCustomProperty wkbk, "CodeExporterSavePath", msoPropertyTypeString, dumpdir
            End If
        End If
    End If
End Sub

Private Sub ImportModules(dirname As String, wkbk As Workbook)
    Dim VBComp As VBComponent, i As Integer, vbdocs As Variant
    Dim fname As String, files As Variant, codemod As CodeModule
    Dim tmp As Variant
    
    ' Find all *.frm *.cls and *.bas files
    ' We cannot tell till we import them whether
    ' there is a name clash or not, so we just have to try
    ' and report failures.
    
    ' We do ask the user if all modules should be removed first
    
    ' Also, we want all code that belongs in
    ' sheets or ThisWorkbook module to end up
    ' there - so we insert into those modules instead
    ' of importing it.
    
    If MsgBox("Do you want to remove all modules first? (This cannot be undone)", vbYesNo) = vbYes Then
        For Each VBComp In wkbk.VBProject.VBComponents
          If VBComp.Type = vbext_ct_ClassModule Or _
            VBComp.Type = vbext_ct_MSForm Or _
            VBComp.Type = vbext_ct_StdModule Then
                wkbk.VBProject.VBComponents.Remove VBComp
            End If
        Next
    End If
    
    vbdocs = Array()
    For Each VBComp In wkbk.VBProject.VBComponents
        If VBComp.Type = vbext_ct_Document Then
            ArrayPush vbdocs, VBComp.Name & ".cls"
        End If
    Next
    
    files = GetDirList(dirname & "\*.bas", vbNormal)
    files = ArrayJoin(files, GetDirList(dirname & "\*.cls", vbNormal))
    files = ArrayJoin(files, GetDirList(dirname & "\*.frm", vbNormal))
    
    For i = 0 To UBound(files)
        If InArray(vbdocs, files(i)) Then
            ' don't import - insert the text into the relevant module
            ' and then delete the header lines
            Set codemod = wkbk.VBProject.VBComponents(Left(CStr(files(i)), Len(files(i)) - 4)).CodeModule
            If codemod.CountOfLines > 0 Then
                codemod.DeleteLines 1, codemod.CountOfLines
            End If
            codemod.AddFromFile dirname & "\" & CStr(files(i))
            If Left(codemod.Lines(1, 1), 7) = "VERSION" Then
                codemod.DeleteLines 1, 4
            End If
        Else
            On Error GoTo ImportFailed
            wkbk.VBProject.VBComponents.Import dirname & "\" & CStr(files(i))
            On Error GoTo 0
            ' there seems to be a bug in import/export - repeated imports/exports
            ' result in leading new lines appearing in modules.  So we trim
            ' all leading newlines
            Set codemod = wkbk.VBProject.VBComponents(Left(CStr(files(i)), Len(files(i)) - 4)).CodeModule
            Do
                If codemod.Lines(1, 1) = "" Then
                    codemod.DeleteLines 1, 1
                Else
                    Exit Do
                End If
            Loop
            GoTo Skip1
ImportFailed:
            MsgBox "Importing file " & files(i) & " failed: " & VBA.Chr(10) & _
            "Err " & Err.Number & ": " & Err.Description
            On Error GoTo 0
Skip1:
        End If
            
    Next i
End Sub

Private Function GetDirList(filepath As String, Optional attributes = vbNormal) As Variant
    Dim tmp As String
    GetDirList = Array()
    tmp = DirS(filepath, attributes)
    Do Until tmp = ""
        ArrayPush GetDirList, tmp
        tmp = Dir()
    Loop
End Function


