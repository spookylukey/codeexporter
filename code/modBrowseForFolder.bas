Attribute VB_Name = "modBrowseForFolder"
Option Explicit



Public Type BROWSEINFO
    hOwner As Long                 'Handle to the owner window for the dialog box.
    
    pidlRoot As Long                'Address of an ITEMIDLIST structure specifying the location
                                                'of the root folder from which to browse. Only the specified
                                                'folder and its subfolders appear in the dialog box.
                                                'This member can be NULL; in that case, the namespace
                                                'root (the desktop folder) is used.
                                                
    pszDisplayName As String    'Address of a buffer to receive the display name of the folder
                                                'selected by the user. The size of this buffer is assumed to
                                                'be MAX_PATH bytes.
    
    lpszTitle As String                 'Address of a null-terminated string that is displayed above
                                                'the tree view control in the dialog box.  This string can be
                                                'used to specify instructions to the user.

    ulFlags As Long                    'Flags specifying the options for the dialog box.
                                                'See constants below
                                                
    lpfn As Long                        'Address of an application-defined function that the dialog box calls
                                                'when an event occurs. For more information, see the
                                                'BrowseCallbackProc function. This member can be NULL.
                                                
    lParam As Long                  'Application-defined value that the dialog box passes to the
                                                'callback function (in pData), if one is specified
                                                
    iImage As Long                  'Variable to receive the image associated with the selected folder.
                                                'The image is specified as an index to the system image list.
End Type

Public Const WM_USER = &H400
Public Const MAX_PATH = 260

'ulFlag constants
Public Const BIF_RETURNONLYFSDIRS = &H1      'Only return file system directories.
                                                                  'If the user selects folders that are not
                                                                  'part of the file system, the OK button is grayed.
                                                                
Public Const BIF_DONTGOBELOWDOMAIN = &H2  'Do not include network folders below the
                                                                    'domain level in the tree view control

Public Const BIF_STATUSTEXT = &H4                    'Include a status area in the dialog box.
                                                                    'The callback function can set the status text
                                                                    'by sending messages to the dialog box.

Public Const BIF_RETURNFSANCESTORS = &H8      'Only return file system ancestors. If the user selects
                                                                    'anything other than a file system ancestor, the OK button is grayed
                                                                    
Public Const BIF_EDITBOX = &H10                        'Version 4.71. The browse dialog includes an edit control
                                                                    'in which the user can type the name of an item.

Public Const BIF_VALIDATE = &H20                       'Version 4.71. If the user types an invalid name into the
                                                                    'edit box, the browse dialog will call the application's
                                                                    'BrowseCallbackProc with the BFFM_VALIDATEFAILED
                                                                    ' message. This flag is ignored if BIF_EDITBOX is not specified
                                                                    
Public Const BIF_NEWDIALOGSTYLE = &H40      'Version 5.0. New dialog style with context menu and resizability

Public Const BIF_BROWSEINCLUDEURLS = &H80   'Version 5.0. Allow URLs to be displayed or entered. Requires BIF_USENEWUI.

Public Const BIF_BROWSEFORCOMPUTER = &H1000 'Only return computers. If the user selects anything
                                                                            'other than a computer, the OK button is grayed

Public Const BIF_BROWSEFORPRINTER = &H2000    'Only return printers. If the user selects anything
                                                                            'other than a printer, the OK button is grayed.

Public Const BIF_BROWSEINCLUDEFILES = &H4000   'The browse dialog will display files as well as folders

Public Const BIF_SHAREABLE = &H8000   'Version 5.0.  Allow display of remote shareable resources.  Requires BIF_USENEWUI.

'Message from browser to callback function constants

Public Const BFFM_INITIALIZED = 1   'Indicates the browse dialog box has finished initializing.
                                                           'The lParam parameter is NULL.
                                                    
Public Const BFFM_SELCHANGED = 2    'Indicates the selection has changed. The lParam parameter
                                                    'contains the address of the item identifier list for the newly selected folder.
                                                    
Public Const BFFM_VALIDATEFAILED = 3  'Version 4.71. Indicates the user typed an invalid name into the edit
                                                        'box of the browse dialog. The lParam parameter is the address of
                                                        'a character buffer that contains the invalid name.
                                                        'An application can use this message to inform the user that the
                                                        'name entered was not valid. Return zero to allow the dialog to be
                                                        'dismissed or nonzero to keep the dialog displayed.

' messages to browser from callback function
Public Const BFFM_SETSTATUSTEXTA = WM_USER + 100
Public Const BFFM_ENABLEOK = WM_USER + 101
Public Const BFFM_SETSELECTIONA = WM_USER + 102
Public Const BFFM_SETSELECTIONW = WM_USER + 103
Public Const BFFM_SETSTATUSTEXTW = WM_USER + 104

Public Const LMEM_FIXED = &H0
Public Const LMEM_ZEROINIT = &H40
Public Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

'Main Browse for directory function
Declare Function SHBrowseForFolder Lib "shell32.dll" _
 Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
'Gets path from pidl
Declare Function SHGetPathFromIDList Lib "shell32.dll" _
  Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
'Used by callback function to communicate with the browser
Declare Function SendMessage Lib "user32" _
 Alias "SendMessageA" (ByVal hwnd As Long, _
  ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, _
      hpvSource As Any, ByVal cbCopy As Long)

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Public Declare Function LocalAlloc Lib "kernel32" _
   (ByVal uFlags As Long, _
    ByVal uBytes As Long) As Long
    
Public Declare Function LocalFree Lib "kernel32" _
   (ByVal hMem As Long) As Long


''The following declarations for the option to center the dialog in the user's screen
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Const SM_CXFULLSCREEN = 16
Public Const SM_CYFULLSCREEN = 17

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, _
 ByVal x As Long, _
   ByVal y As Long, _
    ByVal nWidth As Long, _
     ByVal nHeight As Long, ByVal bRepaint As Long) As Long
''End of dialog centering declarations




Dim CntrDialog As Boolean

Function GetDirectory(InitDir As String, Flags As Long, CntrDlg As Boolean, Msg) As String
    Dim bInfo As BROWSEINFO
    Dim pidl As Long, lpInitDir As Long
    
    CntrDialog = CntrDlg ''Copy dialog centering setting to module level variable so callback function can see it
    With bInfo
        .pidlRoot = 0 'Root folder = Desktop
        .lpszTitle = Msg
        .ulFlags = Flags

        lpInitDir = LocalAlloc(LPTR, Len(InitDir) + 1)
        CopyMemory ByVal lpInitDir, ByVal InitDir, Len(InitDir) + 1
        .lParam = lpInitDir
        
        If val(Application.Version) > 8 Then 'Establish the callback function
            .lpfn = BrowseCallBackFuncAddress
        Else
            .lpfn = AddrOf("BrowseCallBackFunc")
        End If
    End With
    'Display the dialog
    pidl = SHBrowseForFolder(bInfo)
    'Get path string from pidl
    GetDirectory = GetPathFromID(pidl)
    CoTaskMemFree pidl
    LocalFree lpInitDir
End Function

'Windows calls this function when the dialog events occur
Function BrowseCallBackFunc(ByVal hwnd As Long, ByVal Msg As Long, ByVal lParam As Long, ByVal pData As Long) As Long
    Select Case Msg
        Case BFFM_INITIALIZED
            'Dialog is being initialized. I use this to set the initial directory and to center the dialog if the requested
            SendMessage hwnd, BFFM_SETSELECTIONA, 1, pData 'Send message to dialog
            If CntrDialog Then CenterDialog hwnd
        Case BFFM_SELCHANGED
            'User selected a folder - change status text ("show status text" option must be set to see this)
            SendMessage hwnd, BFFM_SETSTATUSTEXTA, 0, GetPathFromID(lParam)
        Case BFFM_VALIDATEFAILED
            'This message is sent  to the callback function only if "Allow direct entry" and
            '"Validate direct entry" have been be set on the Demo worksheet
            'and the user's direct entry is not valid.
            '"Show status text" must be set on to see error message we send back to the dialog
            Beep
            SendMessage hwnd, BFFM_SETSTATUSTEXTA, 0, "Bad Directory"
            BrowseCallBackFunc = 1 'Block dialog closing
            Exit Function
    End Select
    BrowseCallBackFunc = 0 'Allow dialog to close
End Function

'Converts a PIDL to a string
Function GetPathFromID(ID As Long) As String
    Dim Result As Boolean, Path As String * MAX_PATH
    Result = SHGetPathFromIDList(ID, Path)
    If Result Then
        GetPathFromID = Left(Path, InStr(Path, Chr$(0)) - 1)
    Else
        GetPathFromID = ""
    End If
End Function

'XL8 is very unhappy about using Excel 9's AddressOf operator, but as long as it is in a
' function that is not called when run on XL8, it seems to allow it to exist.
Function BrowseCallBackFuncAddress() As Long
    'BrowseCallBackFuncAddress = Long2Long(AddressOf BrowseCallBackFunc)
End Function

'It is not possible to assign the result of AddressOf (which is a Long) directly to a member
'of a user defined data type.  This explicitly "converts" it to a Long and
'allows the assignment
Function Long2Long(x As Long) As Long
    Long2Long = x
End Function

'Centers dialog on desktop
Sub CenterDialog(hwnd As Long)
     Dim WinRect As RECT, ScrWidth As Integer, ScrHeight As Integer
    Dim DlgWidth As Integer, DlgHeight As Integer
    GetWindowRect hwnd, WinRect
    DlgWidth = WinRect.Right - WinRect.Left
    DlgHeight = WinRect.Bottom - WinRect.Top
    ScrWidth = GetSystemMetrics(SM_CXFULLSCREEN)
    ScrHeight = GetSystemMetrics(SM_CYFULLSCREEN)
    MoveWindow hwnd, (ScrWidth - DlgWidth) / 2, (ScrHeight - DlgHeight) / 2, DlgWidth, DlgHeight, 1
End Sub

