Attribute VB_Name = "basAddrOf"
Option Explicit

'NOTE: The brilliant AddrOf function herein contained is the work of Ken Getz and
'Michael Kaplan.  Published in the May 1998 issue of
'Microsoft Office & Visual Basic for Applications Developer (page 46).

'Office 97 does not support the "AddressOf" operator which is needed to tell Windows
'where our "call back" function is.  Getz and Kaplan figured out a workaround.

'The rest of this module is entirely their work.

'-------------------------------------------------------------------------------------------------------------------
'   Declarations
'
'   These function names were puzzled out by using DUMPBIN /exports
'   with VBA332.DLL and then puzzling out parameter names and types
'   through a lot of trial and error and over 100 IPFs in MSACCESS.EXE
'   and VBA332.DLL.
'
'   These parameters may not be named properly but seem to be correct in
'   light of the function names and what each parameter does.
'
'   EbGetExecutingProj: Gives you a handle to the current VBA project
'   TipGetFunctionId: Gives you a function ID given a function name
'   TipGetLpfnOfFunctionId: Gives you a pointer a function given its function ID
'
'-------------------------------------------------------------------------------------------------------------------
Private Declare Function GetCurrentVbaProject _
 Lib "vba332.dll" Alias "EbGetExecutingProj" _
 (hProject As Long) As Long
Private Declare Function GetFuncID _
 Lib "vba332.dll" Alias "TipGetFunctionId" _
 (ByVal hProject As Long, ByVal strFunctionName As String, _
 ByRef strFunctionId As String) As Long
Private Declare Function GetAddr _
 Lib "vba332.dll" Alias "TipGetLpfnOfFunctionId" _
 (ByVal hProject As Long, ByVal strFunctionId As String, _
 ByRef lpfn As Long) As Long

'-------------------------------------------------------------------------------------------------------------------
'   AddrOf
'
'   Returns a function pointer of a VBA public function given its name. This function
'   gives similar functionality to VBA as VB5 has with the AddressOf param type.
'
'   NOTE: This function only seems to work if the proc you are trying to get a pointer
'       to is in the current project. This makes sense, since we are using a function
'       named EbGetExecutingProj.
'-------------------------------------------------------------------------------------------------------------------
Public Function AddrOf(strFuncName As String) As Long
    Dim hProject As Long
    Dim lngResult As Long
    Dim strID As String
    Dim lpfn As Long
    Dim strFuncNameUnicode As String
    
    Const NO_ERROR = 0
    
    ' The function name must be in Unicode, so convert it.
    strFuncNameUnicode = StrConv(strFuncName, vbUnicode)
    
    ' Get the current VBA project
    ' The results of GetCurrentVBAProject seemed inconsistent, in our tests,
    ' so now we just check the project handle when the function returns.
    Call GetCurrentVbaProject(hProject)
    
    ' Make sure we got a project handle... we always should, but you never know!
    If hProject <> 0 Then
        ' Get the VBA function ID (whatever that is!)
        lngResult = GetFuncID( _
         hProject, strFuncNameUnicode, strID)
        
        ' We have to check this because we GPF if we try to get a function pointer
        ' of a non-existent function.
        If lngResult = NO_ERROR Then
            ' Get the function pointer.
            lngResult = GetAddr(hProject, strID, lpfn)
            
            If lngResult = NO_ERROR Then
                AddrOf = lpfn
            End If
        End If
    End If
End Function


