Attribute VB_Name = "modExport"
' modExport.bas  Copyright 2003 Luke Plant
' L.Plant.98@cantab.net

' * This program is free software; you can redistribute it and/or modify
' * it under the terms of the GNU General Public License as published by
' * the Free Software Foundation; either version 2 of the License, or
' * (at your option) any later version.
' *
' * This program is distributed in the hope that it will be useful,
' * but WITHOUT ANY WARRANTY; without even the implied warranty of
' * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' * GNU General Public License for more details.
' *
' * You should have received a copy of the GNU General Public License
' * along with this program; if not, write to the Free Software
' * Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA 02111-1307, USA.

' The modules basAddrOf and modBrowseForFolder are taken from
' Jim Rech's BrowseForFolder.xls demo file.

Option Explicit

' Export all (or selected) modules in a project to
' a directory

' called from menus/toolbars
Public Sub CodeExporter_Export()
    CodeExporter_ExportCodeFromWorkbook ActiveWorkbook
End Sub

' To automate exporting of code, use a routine like this.
' This requires codeexporter to be open or installed as an add-in
'Sub ExportMe()
'    Run "codeexporter.xla!CodeExporter_ExportCodeFromWorkbook", ThisWorkbook
'End Sub

Public Sub CodeExporter_ExportCodeFromWorkbook(wkbk As Workbook)
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
        defdir = wkbk.Path & defdir
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
                If InStr(1, dumpdir, wkbk.Path) = 1 Then
                    dumpdir = VBA.Mid(dumpdir, Len(wkbk.Path) + 1)
                End If
                SetCustomProperty wkbk, "CodeExporterSavePath", msoPropertyTypeString, dumpdir
            End If
        End If
    End If
End Sub

Private Sub ExportModules(dirname As String, wkbk As Workbook)
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
        ' contrary to documentation, VBComp.Export seems to
        ' simply overwrite any existing files, which is what we
        ' want.
        VBComp.Export dirname & "\" & fname
    Next
End Sub
