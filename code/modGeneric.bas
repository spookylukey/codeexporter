Attribute VB_Name = "modGeneric"
Option Explicit

' Module modGeneric
' Contains misc functions useful for any VB project
' Current version: 0.4
'---------------------------------------------------------------------------------------
' ChangeLog

' v 0.4
' - added ArrayPushO()
' - added ArrayRemoveC()
' - added ArrayCopy()
' - removed GeneralMsgBox

' v 0.3
' - added ItoS()

' v 0.2
' - removed function Max()
' - added DirS()
' - added DisableEvents(), EnableEvents()
' - added InStrR()

' v 0.1

'---------------------------------------------------------------------------------------
' this Module contains code that could be useful for any VBA project
Private ScreenLock As Integer
Private EventLock As Integer

Function Replace$(ByVal subject As String, f As String, r As String)
    Replace$ = WorksheetFunction.Substitute(subject, f, r)
End Function

' need to avoid use of Error statements in this function,
' as Printf is often used in statements like
' MsgBox Printf("Err %s: %s", Array(Err.Number, Err.Description))

Function Printf(str As String, sub1 As Variant) As String
    ' accepts only %s as replacement marker. It can be
    ' a string or a number - the CStr() conversion is done
    ' sub1 can also be an array of values
    Dim i As Integer, temp As String
    If (VarType(sub1) And vbArray) = vbArray Then
        temp = str
        For i = 0 To UBound(sub1)
            temp = Printf(temp, sub1(i))
        Next i
        Printf = temp
    Else
        Printf = replaceStr1("%s", CStr(sub1), str)
    End If
End Function

' only replace first instance, the string to replace might not exist
Function replaceStr1(findStr As String, rStr As String, subj As String) As String
    Dim pos As Integer
    replaceStr1 = ""
    pos = VBA.InStr(1, subj, findStr, 1)
    If pos <> 0 Then
        replaceStr1 = VBA.Left$(subj, pos - 1) & rStr + VBA.Mid$(subj, pos + Len(findStr))
    Else
        replaceStr1 = subj
    End If
End Function

' no checking is done here
Function InArray(arr As Variant, search As Variant) As Boolean
    Dim i As Integer
    If EmptyArray(arr) Then
        InArray = False
        Exit Function
    End If
    For i = LBound(arr) To UBound(arr)
        If IsNull(arr(i)) And IsNull(search) Then
            InArray = True
            Exit Function
        ElseIf arr(i) = search Then
            InArray = True
            Exit Function
        End If
    Next i
    InArray = False
End Function

' no checking here either
Function ArraySeach(ByRef arr As Variant, ByRef search As Variant) As Integer
    Dim i As Integer
    ArraySeach = -1
    For i = LBound(arr) To UBound(arr)
        If search = arr(i) Then
            ArraySeach = i
            Exit Function
        End If
    Next i
End Function

' arr must be a variant containing an array
Sub ArrayPush(ByRef arr As Variant, item As Variant)
    Dim i As Integer
    If Not IsArray(arr) Then
        arr = Array()
    End If
    i = UBound(arr)
    ReDim Preserve arr(i + 1)
    arr(i + 1) = item
End Sub

' arr must be a variant containing an array
Sub ArrayPop(ByRef arr As Variant)
    If Not EmptyArray(arr) Then
        If UBound(arr) = 0 Then
            arr = Array()
        Else
            ReDim Preserve arr(UBound(arr) - 1)
        End If
    End If
End Sub

' arr must be a variant containing an array
' add element at beginning
Sub ArrayUnshift(ByRef arr As Variant, item As Variant)
    Dim i As Integer
    If Not IsArray(arr) Then
        arr = Array()
    End If
    ReDim Preserve arr(UBound(arr) + 1)
    For i = UBound(arr) To 1 Step -1
        arr(i) = arr(i - 1)
    Next i
    arr(0) = item
End Sub

Sub ArrayShift(ByRef arr As Variant)
    Dim i As Integer
    If Not EmptyArray(arr) Then
        For i = 0 To UBound(arr) - 1
            arr(i) = arr(i + 1)
        Next i
        ReDim Preserve arr(UBound(arr) - 1)
    End If
End Sub


Function EmptyArray(ByRef arr As Variant) As Boolean
    If Not IsArray(arr) Then
        EmptyArray = True
        Exit Function
    End If
    If UBound(arr) = -1 Then
        EmptyArray = True
        Exit Function
    End If
    EmptyArray = False
End Function

Function ArrayJoin(arr1 As Variant, arr2 As Variant)
    Dim t As Integer, a As Variant, i As Integer
    
    If EmptyArray(arr1) Then
        ArrayJoin = arr2
    ElseIf EmptyArray(arr2) Then
        ArrayJoin = arr1
    Else
        t = UBound(arr1) + UBound(arr2) + 1
        
        a = Array()
        ReDim a(t)
        For i = 0 To UBound(arr1)
            a(i) = arr1(i)
        Next i
        For i = 0 To UBound(arr2)
            a(i + UBound(arr1) + 1) = arr2(i)
        Next i
        ArrayJoin = a
    End If
        
End Function

Sub ArrayRemove(ByRef arr As Variant, starti As Integer, length As Integer)
    ' length is the number of items to remove
    Dim i As Integer
    If starti < 0 Or starti > UBound(arr) Then
        Exit Sub
    ElseIf length = 0 Then
        Exit Sub
    Else
        If starti + length > UBound(arr) Then
            If starti = 0 Then
                arr = Array()
                Exit Sub
            Else
                ReDim Preserve arr(starti - 1)
                Exit Sub
            End If
        Else
            For i = starti To UBound(arr) - length
                arr(i) = arr(i + length)
            Next i
            ReDim Preserve arr(UBound(arr) - length)
        End If
    End If
End Sub

Function ArrayRemoveC(arr As Variant, starti As Integer, length As Integer) As Variant
    ArrayRemoveC = ArrayCopy(arr)
    ArrayRemove ArrayRemoveC, starti, length
End Function

Function ArrayCopy(arr As Variant)
' Assigning arrays directly e.g arr2 = arr1 seems to be buggy in VBA
' Modifiying arr2 in this example will corrupt/modify arr1, depending on
' how arr1 was constructed and what data it stores
' [e.g.s
'   arr1 = array(array(1,2),array(3,4))
'   arr2 = arr1
'   ' arr2 will work fine and completely independently of arr1
'
'   arr1 = array()
'   ArrayPush arr1, array(1,2)
'   ArrayPush arr1, array(3,4)
'   arr2 = arr1
'   ' modifying arr2 will modify/corrupt arr1.
'
' The same behaviour demonstrated using only VBA functions
'Sub VBAArrayErrorExample()
'    Dim arr1, arr2, arr3, arr4, arr5, arr6
'    arr1 = Array(Array(1, 2), Array(3, 4))
'    arr2 = arr1
'    arr2(0)(0) = 100
'    Debug.Print arr1(0)(0), arr2(0)(0)  ' prints:    1       100
'
'    arr3 = Array()
'    ReDim Preserve arr3(0)
'    arr3(0) = Array(1, 2)
'    ReDim Preserve arr3(1)
'    arr3(1) = Array(3, 4)
'    arr4 = arr3
'    arr4(0)(0) = 100
'    Debug.Print arr3(0)(0), arr4(0)(0)  ' prints:    100     100
'
'    arr5 = Array(0, 0)
'    arr5(0) = Array(1, 2)
'    arr5(1) = Array(3, 4)
'    arr6 = arr5
'    arr6(0)(0) = 100
'    Debug.Print arr5(0)(0), arr6(0)(0)  ' prints:    1       100
'End Sub

'   Hence the need for this function that will produce a clean copy

    Dim i As Integer, arr2 As Variant
    arr2 = Array()
    For i = 0 To UBound(arr)
        If (VarType(arr(i)) And vbArray) = vbArray Then
            ArrayPush arr2, ArrayCopy(arr(i))
        Else
            ArrayPush arr2, arr(i)
        End If
    Next
    ArrayCopy = arr2
End Function

Function Join(ByRef arr As Variant, glue As String) As String
    Dim i As Integer
    Join = ""
    On Error GoTo EmptyArray
    For i = LBound(arr) To UBound(arr)
        If i <> LBound(arr) Then
            Join = Join & glue
        End If
        Join = Join & CStr(arr(i))
    Next i
EmptyArray:

End Function
Function Split(txt As String, boundary As String, Optional limit As Integer = 0) As Variant
    ' limit is the maximum number of parts to split into
    ' If txt is zero length, an empty array is returned
    '
    Dim a As Variant, pos As Integer, pos2 As Integer, done As Boolean
    a = Array()
    pos = 1
    pos2 = 0
    Do Until pos2 >= Len(txt) Or (UBound(a) > limit And limit > 0)
        pos2 = InStr(pos, txt, boundary)
        If pos2 = 0 Or (UBound(a) = limit - 2 And limit > 0) Then pos2 = Len(txt) + 1
        ArrayPush a, VBA.Mid$(txt, pos, pos2 - pos)
        pos = pos2 + Len(boundary)
    Loop
    Split = a
End Function

' requires arr is a variant, not an array
Sub ArrayClear(ByRef arr As Variant)
    arr = Array()
End Sub

' Quick sort algorithm.  Sorts array directly, not a copy
Sub SortArray(ByRef arr As Variant, f As Integer, l As Integer)
    Dim index As Integer
    If (f < l) Then
        index = PartitionArray(arr, f, l)
        SortArray arr, index + 1, l
        SortArray arr, f, index - 1
    End If
End Sub

Function PartitionArray(ByRef arr As Variant, f As Integer, l As Integer) As Integer
    Dim pivot As Variant, i As Integer, firstUnknown As Integer, _
    lastS1 As Integer, tmp As Variant
    pivot = arr(f)
    firstUnknown = f + 1
    lastS1 = f
    Do While firstUnknown <= l
        If arr(firstUnknown) < pivot Then
            lastS1 = lastS1 + 1
            tmp = arr(lastS1)
            arr(lastS1) = arr(firstUnknown)
            arr(firstUnknown) = tmp
        End If
        firstUnknown = firstUnknown + 1
    Loop
    tmp = arr(f)
    arr(f) = arr(lastS1)
    arr(lastS1) = tmp
    PartitionArray = lastS1
End Function

' sorting for an array of arrays (NB, not 2D arrays really) where the column to compare is cmp
Sub SortArray2D(ByRef arr As Variant, f As Integer, l As Integer, cmp As Integer)
    Dim index As Integer
    If (f < l) Then
        index = PartitionArray2D(arr, f, l, cmp)
        SortArray2D arr, index + 1, l, cmp
        SortArray2D arr, f, index - 1, cmp
    End If
End Sub

Function PartitionArray2D(ByRef arr As Variant, f As Integer, l As Integer, cmp As Integer) As Integer
    Dim pivot As Variant, i As Integer, firstUnknown As Integer, _
    lastS1 As Integer, tmp As Variant
    pivot = arr(f)(cmp)
    firstUnknown = f + 1
    lastS1 = f
    Do While firstUnknown <= l
        If arr(firstUnknown)(cmp) < pivot Then
            lastS1 = lastS1 + 1
            tmp = arr(lastS1)
            arr(lastS1) = arr(firstUnknown)
            arr(firstUnknown) = tmp
        End If
        firstUnknown = firstUnknown + 1
    Loop
    tmp = arr(f)
    arr(f) = arr(lastS1)
    arr(lastS1) = tmp
    PartitionArray2D = lastS1
End Function

Function ArrayConv1(a As Variant) As Variant
    ' convert an array of arrays into a 2D array
    Dim temp(), i As Integer, j As Integer
    ReDim temp(UBound(a), UBound(a(0)))
    For i = 0 To UBound(a)
        For j = 0 To UBound(a(0))
            temp(i, j) = a(i)(j)
        Next j
    Next i
    ArrayConv1 = temp
End Function

' type casting functions, with extra 'F'orce and no errors

Function CIntF(v As Variant) As Integer
    On Error GoTo converterror
    CIntF = CInt(v)
    On Error GoTo 0
    Exit Function
converterror:
    CIntF = 0
End Function

Function CLngF(v As Variant) As Long
    On Error GoTo converterror
    CLngF = CLng(v)
    On Error GoTo 0
    Exit Function
converterror:
    CLngF = 0
End Function

Function CBoolF(v As Variant) As Boolean
    On Error GoTo converterror
    CBoolF = CBool(v)
    On Error GoTo 0
    Exit Function
converterror:
    CBoolF = False
End Function

Function CStrF(v As Variant) As String
    On Error GoTo converterror
    CStrF = CStr(v)
    On Error GoTo 0
    Exit Function
converterror:
    CStrF = ""
End Function


Function CDblF(v As Variant) As Long
    On Error GoTo converterror
    CDblF = CDbl(v)
    On Error GoTo 0
    Exit Function
converterror:
    CDblF = 0
End Function

Function CDateF(v As Variant) As Date
    On Error GoTo converterror
    CDateF = CDate(v)
    On Error GoTo 0
    Exit Function
converterror:
    CDateF = CDate(0)
End Function


Sub ScreenFreeze()
    If ScreenLock = 0 Or Application.ScreenUpdating Then
        Application.ScreenUpdating = False
    End If
    ScreenLock = ScreenLock + 1
End Sub

Sub ScreenThaw()
    If ScreenLock <> 0 Then
        ScreenLock = ScreenLock - 1
        If ScreenLock = 0 Then
            Application.ScreenUpdating = True
        End If
    End If
End Sub


Sub ScreenThawForce()
    ScreenLock = 0
    If Not Application.ScreenUpdating Then
        Application.ScreenUpdating = True
    End If
End Sub


Sub DisableEvents()
    If EventLock = 0 Or Application.EnableEvents Then
        Application.EnableEvents = False
    End If
    EventLock = EventLock + 1
End Sub

Sub EnableEvents()
    If EventLock <> 0 Then
        EventLock = EventLock - 1
        If EventLock = 0 Then
            Application.EnableEvents = True
        End If
    End If
End Sub

' a safe version of Dir, to be used as replacement for first
' call to Dir.  This will not produce errors if the drive
' does not exist
Public Function DirS(PathName As String, Optional attributes = vbNormal) As String
    On Error GoTo BadPath
    DirS = Dir(PathName, attributes)
    Exit Function
    
BadPath:
    DirS = ""
End Function

Function InStrR(haystack As String, needle As String, Optional Start As Integer = 0) As Integer
    ' search a string from the right, return the position
    If Start = 0 Then Start = Len(haystack)
    Dim i As Integer, k As Integer
    InStrR = 0
    k = Len(needle)
    If k = 0 Then Exit Function
    For i = Start To 1 Step -1
        If VBA.Mid(haystack, i, k) = needle Then
            InStrR = i
            Exit Function
        End If
    Next i
End Function

Function ItoS(i As Integer) As String
    ItoS = VBA.Trim(VBA.str(i))
End Function

' object safe version of the above. It uses Set for assignment
' arr must be a variant containing an array
Sub ArrayPushO(ByRef arr As Variant, item As Object)
    Dim i As Integer
    If Not IsArray(arr) Then
        arr = Array()
    End If
    i = UBound(arr)
    ReDim Preserve arr(i + 1)
    Set arr(i + 1) = item
End Sub
