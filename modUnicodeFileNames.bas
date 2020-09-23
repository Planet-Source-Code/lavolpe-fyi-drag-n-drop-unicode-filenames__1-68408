Attribute VB_Name = "modUnicodeFileNames"
Option Explicit

' ////////////////////////////////////////////////////////////////
' Kernel32/User32 APIs for Unicode Filename Support
' ////////////////////////////////////////////////////////////////
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
' ////////////////////////////////////////////////////////////////

' ////////////////////////////////////////////////////////////////
' used to create a stdPicture from a byte array
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
' ////////////////////////////////////////////////////////////////

' ////////////////////////////////////////////////////////////////
' Unicode-capable Drag and Drop of file names with wide characters
' ////////////////////////////////////////////////////////////////
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, _
    ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As VbVarType, _
    ByVal paCNT As Long, ByRef paTypes As Integer, _
    ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (lpString As Any) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long

' ////////////////////////////////////////////////////////////////
' Unicode-capable Pasting of file names with wide characters
' ////////////////////////////////////////////////////////////////
Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
' ////////////////////////////////////////////////////////////////

Private Type FORMATETC
    cfFormat As Long
    pDVTARGETDEVICE As Long
    dwAspect As Long
    lindex As Long
    TYMED As Long
End Type

Private Type DROPFILES
    pFiles As Long
    ptX As Long
    ptY As Long
    fNC As Long
    fWide As Long
End Type

Private Type STGMEDIUM
    TYMED As Long
    data As Long
    pUnkForRelease As IUnknown
End Type

Private Const TYMED_HGLOBAL = 1
Private Const DVASPECT_CONTENT = 1


Public Function isStringANSI(inText As String) As Boolean

    ' simple test to determine if passed string is ANSI-like or not.
    ' In other words, does it contain unicode characters.
    
    Dim tArray() As Byte
    Dim X As Long
    
    If inText = vbNullString Then
    
        isStringANSI = True
    
    Else
    
        tArray = inText
        For X = LBound(tArray) + 1 To UBound(tArray) Step 2
            If Not tArray(X) = 0 Then Exit For
        Next
        
        isStringANSI = (X > UBound(tArray))
        
    End If

End Function

Public Function LoadPictureW(ByVal FileName As String) As StdPicture

    If FileName = vbNullString Then Exit Function
    
    If isStringANSI(FileName) Then
        ' use VB's LoadPicture if filename has no unicode characters
        On Error Resume Next
        Set LoadPictureW = LoadPicture(FileName)
        If Err Then
            Err.Raise Err.Number, "modUnicodeFileNames:LoadPictureW", Err.Description
        End If
    
    Else
        ' otherwise use APIs to open file, cache contents, then send
        ' to API to create stdPicture
        Dim hFile As Long
        
        On Error Resume Next
        hFile = OpenFileW(FileName)
        If Err Then
            Err.Raise Err.Number, "modUnicodeFileNames:LoadPictureW", Err.Description
        ElseIf hFile = 0& Then
            Err.Raise 53, "modUnicodeFileNames:LoadPictureW"
        Else
            
            Dim lLen As Long
            Dim o_hMem  As Long
            Dim o_lpMem  As Long
            Dim aSize As Long
            Dim aGUID(0 To 3) As Long
            Dim IIStream As IUnknown
    
            lLen = GetFileSize(hFile, 0&)
            If lLen > 1 Then
                
                o_hMem = GlobalAlloc(&H2&, lLen)
                If Not o_hMem = 0& Then
                    o_lpMem = GlobalLock(o_hMem)
                    If Not o_lpMem = 0& Then
                        ReadFile hFile, ByVal o_lpMem, lLen, aSize, ByVal 0&
                        CloseHandle hFile
                        Call GlobalUnlock(o_hMem)
                        hFile = 0&
                        If aSize = lLen Then
                            If CreateStreamOnHGlobal(o_hMem, 1&, IIStream) = 0& Then
                                aGUID(0) = &H7BF80980    ' GUID for stdPicture
                                aGUID(1) = &H101ABF32
                                aGUID(2) = &HAA00BB8B
                                aGUID(3) = &HAB0C3000
                                If OleLoadPicture(ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0), LoadPictureW) <> 0 Then
                                    Err.Raise 75, "modUnicodeFileNames:LoadPictureW"
                                End If
                            End If
                        End If
                    End If
                End If
                If Not hFile = 0& Then CloseHandle hFile
            End If
        End If
        
    End If
End Function

Public Function GetDroppedFiles(oData As DataObject, ListOfFiles() As String) As Long

    ' Caution: Editing this routine after it has been called may crash the IDE
    ' I believe I have fixed that issue but am not 100% positive
    
    ' See posting by John Kleinen for more information regarding this method of calling GetData
    ' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=49268&lngWId=1
    
    If oData.GetFormat(vbCFFiles) = False Then Exit Function
    
    Dim fmtEtc As FORMATETC, pMedium As STGMEDIUM
    Dim dFiles As DROPFILES
    Dim Vars(0 To 1) As Variant, pVars(0 To 1) As Long, pVartypes(0 To 1) As Integer
    Dim varRtn As Variant
    Dim iFiles As Long, iCount As Long, hDrop As Long
    Dim lLen As Long, sFile As String
    
    Dim IID_IDataObject As Long ' IDataObject Interface ID
    Const IDataObjVTable_GetData As Long = 12 ' 4th vtable entry
    Const CC_STDCALL As Long = 4&

    With fmtEtc
        .cfFormat = vbCFFiles         ' same as CF_DROP
        .lindex = -1                    ' want all data
        .TYMED = TYMED_HGLOBAL        ' want global ptr to files
        .dwAspect = DVASPECT_CONTENT  ' no rendering
    End With

    ' The IDataObject pointer is 16 bytes after VBs DataObject
    CopyMemory IID_IDataObject, ByVal ObjPtr(oData) + 16, 4&
    
    ' Here we are going to do something very new to me and kinda cool
    ' Since we know the objPtr of the IDataObject interface, we therefore know
    ' the beginning of the interface's VTable
    
    ' So, if we know the VTable address and we know which function index we want
    ' to call, we can call it directly using the following OLE API. Otherwise we
    ' would need to use a TLB to define the IDataObject interface since VB doesn't
    ' 't expose it. This has some really neat implications if you think about it.
    ' The IDataObject function we want is GetData which is the 4th function in
    ' the VTable... http://msdn2.microsoft.com/en-us/library/ms688421.aspx
    
'////////////////////////////////////////////////////////////////////////////////
' Testing only. Want to know if ref count goes up after I call DispCallFunc API
' By getting the refcount before it is called and again later, we can find out
' The answer is no, it doesn't. Left this in should you be interested
'    Call DispCallFunc(IID_IDataObject, 4, CC_STDCALL, vbLong, 0, 0, 0, varRtn)
'    Debug.Print "AddRef before calling dispcallfunc: "; varRtn - 1
'////////////////////////////////////////////////////////////////////////////////

    pVartypes(0) = vbLong: Vars(0) = VarPtr(fmtEtc): pVars(0) = VarPtr(Vars(0))
    pVartypes(1) = vbLong: Vars(1) = VarPtr(pMedium): pVars(1) = VarPtr(Vars(1))
    
    ' The variants are required by the OLE API: http://msdn2.microsoft.com/en-us/library/ms221473.aspx
    If DispCallFunc(IID_IDataObject, IDataObjVTable_GetData, CC_STDCALL, _
                        vbLong, 2, pVartypes(0), pVars(0), varRtn) = 0 Then
        
        If Not pMedium.data = 0 Then
            ' we have a pointer to the files, kinda sorta
            CopyMemory hDrop, ByVal pMedium.data, 4&
            If Not hDrop = 0 Then
                ' the hDrop is a pointer to a DROPFILES structure
                ' copy the 20-byte structure for our use
                CopyMemory dFiles, ByVal hDrop, 20&
                ' use the pFiles member to track offsets for file names
                dFiles.pFiles = dFiles.pFiles + hDrop
            End If
        End If
        
        ReDim ListOfFiles(1 To oData.Files.Count)
        
        For iCount = 1 To oData.Files.Count
            If dFiles.fWide = 0 Then    ' non-unicode
                ListOfFiles(iCount) = oData.Files(iCount) ' simply copy VB's ANSI DataObject.Files list
            Else
                ' get the length of the current file & multiply by 2 because it is unicode
                ' lstrLenW is supported in Win9x
                lLen = lstrlenW(ByVal dFiles.pFiles) * 2
                sFile = String$(lLen \ 2, 0)    ' build a buffer to hold the file name
                CopyMemory ByVal StrPtr(sFile), ByVal dFiles.pFiles, lLen ' populate the buffer
                ' move the pointer to location for next file, adding 2 because of a double null separator/delimiter btwn file names
                dFiles.pFiles = dFiles.pFiles + lLen + 2
                ' add our file name to the list.
                ListOfFiles(iCount) = sFile ' this may contain unicode characters if your system supports it
            End If
        Next
        
        GlobalFree pMedium.data
        
'////////////////////////////////////////////////////////////////////////////////
' Finish testing. Get refcount after DispFuncCall is called:
'        Call DispCallFunc(IID_IDataObject, 8, CC_STDCALL, vbLong, 0, 0, 0, varRtn)
'        Debug.Print "refcount after calling dispfunccall "; varRtn
'////////////////////////////////////////////////////////////////////////////////

        GetDroppedFiles = iCount - 1
        
    End If
    
End Function

Public Function GetPastedFiles(ListOfFiles() As String) As Long

    ' SPECIAL NOTES:
    ' 1. The DragQueryFileW API can be used to get the unicode filename instead of
    '    parsing the hDrop object like we are going to do here.
    ' 2. However, when do you use DragQueryW or DragQueryA?  The answer is
    '    probably when in NT use W else use A versions.
    ' 3. This method doesn't care which operating system is used and is therefore generic
    
    Dim hDrop As Long
    Dim sFile As String
    Dim lLen As Long
    Dim iCount As Long
    Dim dFiles As DROPFILES

   ' Get handle to CF_HDROP if any:
   If OpenClipboard(0&) = 0 Then Exit Function
        
    hDrop = GetClipboardData(vbCFFiles)
    If Not hDrop = 0 Then   ' then copied/cut files exist in memory
        iCount = DragQueryFile(hDrop, -1&, vbNullString, 0)
        ' the hDrop is a pointer to a DROPFILES structure
        ' copy the 20-byte structure for our use
        CopyMemory dFiles, ByVal hDrop, 20&
        ' use the pFiles member to track offsets for file names
        dFiles.pFiles = dFiles.pFiles + hDrop
    
        ReDim ListOfFiles(1 To iCount)
    
        For iCount = 1 To iCount
            If dFiles.fWide = 0 Then   ' ANSI text, use API to get file name
               lLen = DragQueryFile(hDrop, iCount - 1, vbNullString, 0&)       ' query length
               ListOfFiles(iCount) = String$(lLen, 0)                          ' set up buffer
               DragQueryFile hDrop, iCount - 1, ListOfFiles(iCount), lLen + 1  ' populate buffer
            Else
               ' get the length of the current file & multiply by 2 because it is unicode
               ' lstrLenW is supported in Win9x
               lLen = lstrlenW(ByVal dFiles.pFiles) * 2
               sFile = String$(lLen \ 2, 0)    ' build a buffer to hold the file name
               CopyMemory ByVal StrPtr(sFile), ByVal dFiles.pFiles, lLen ' populate the buffer
               ' move the pointer to location for next file, adding 2 because of a double null separator/delimiter btwn file names
               dFiles.pFiles = dFiles.pFiles + lLen + 2
               ' add our file name to the list.
               ListOfFiles(iCount) = sFile ' this may contain unicode characters if your system supports it
           End If
        Next
        
        GetPastedFiles = iCount - 1
        
    End If
    CloseClipboard

End Function


Private Function OpenFileW(FileName As String)
    ' Function uses APIs to read file name with unicode characters

    Const GENERIC_READ As Long = &H80000000
    Const OPEN_EXISTING = &H3
    Const FILE_SHARE_READ = &H1
    Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
    Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
    Const FILE_ATTRIBUTE_READONLY As Long = &H1
    Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
    Const FILE_ATTRIBUTE_NORMAL = &H80&

    Dim Flags As Long, Access As Long
    Dim Disposition As Long, Share As Long
    Dim hFile As Long

    Access = GENERIC_READ
    Share = FILE_SHARE_READ
    Disposition = OPEN_EXISTING
    Flags = FILE_ATTRIBUTE_ARCHIVE Or FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_NORMAL _
            Or FILE_ATTRIBUTE_READONLY Or FILE_ATTRIBUTE_SYSTEM

    hFile = CreateFileW(StrPtr(FileName), Access, Share, 0&, Disposition, Flags, 0&)
    If hFile = 0 Then
        ' hFile should never be zero. It should be -1 (error) or a valid handle
        ' when hFile is zero, most likely API was called on a Win9x system
        ' so we will call the ANSI version and see if that returns a handle
        hFile = CreateFile(FileName, Access, Share, 0&, Disposition, Flags, 0&)
    End If
    If Not hFile = -1 Then OpenFileW = hFile

End Function
