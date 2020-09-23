VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim hFile As Long, hFileWnd As Long, IOS As IOSup, Max As Long
Dim vFName As String, wndCAP As String, NStruct As NotifyInfo
Dim ExistFlag As Boolean, NoEvent As Boolean
Public Event OnVFileCreate(ByVal VFileName As String)
Attribute OnVFileCreate.VB_Description = "Event fires when the associated Virtual File is created."
Public Event OnVFileDestroy(ByVal VFileName As String)
Attribute OnVFileDestroy.VB_Description = "Event fires when the associated Virtual file is deleted."

Public Event OnVFileInitError(ByVal VFileName As String, ByVal ErrDesc As String)
Attribute OnVFileInitError.VB_Description = "Event fires in case of an error occuring during Initialization of a Virtual File or a connection to it."
Public Event OnVFileInitSuccess(ByVal VFileName As String)
Attribute OnVFileInitSuccess.VB_Description = "Event fires after successful initialization of a connection to a virtual file."
Public Event OnVFileChange(ByVal VFileName As String, ByVal Offset As Long, ByVal BytesWritten As Long)
Attribute OnVFileChange.VB_Description = "Event Fires when a process writes to the associated Virtual File"
Public Event OnVFileReadError(ByVal VFileName As String, ByVal Reason As String)
Attribute OnVFileReadError.VB_Description = "Event fires when an error occurs while attempting to read from a virtual file."
Public Event OnVFileReadSuccess(ByVal VFileName As String, ByVal Offset As Long, ByVal BytesRead As Long)
Attribute OnVFileReadSuccess.VB_Description = "Event fires when a read operation on the virtual file is successful."
Public Event OnVFileWriteError(ByVal VFileName As String, ByVal Reason As String)
Attribute OnVFileWriteError.VB_Description = "Event fires when an error occurs while writing to a virtual file."
Public Event OnVFileWriteSuccess(ByVal VFileName As String, ByVal Offset As Long, ByVal BytesWritten As Long)
Attribute OnVFileWriteSuccess.VB_Description = "Event fires when a write operation on a virtual file has been completed."

Friend Sub RaiseChangeEvent(Off As Long, Size As Long)
    RaiseEvent OnVFileChange(vFName, Off, Size)
End Sub

Function InitializeVirtualFile(VFileName As String, MaxSize As Long) As Boolean
Attribute InitializeVirtualFile.VB_Description = "Initializes a connection to the specified Virtual File.Also creates the Virtual File, if it does not exist."
    Dim T As SECURITY_ATTRIBUTES, Wnd As Long
    VFileName = UCase$(VFileName)
    T.nLength = Len(T)
    T.bInheritHandle = 1

    If hFile <> 0 Then
        RaiseEvent OnVFileInitError(VFileName, "Cannot open multiple virtural files with the same vFile object instance")
        Exit Function
    End If
        
    
    hFile = CreateFileMapping(&HFFFFFFFF, T, PAGE_READWRITE, 0, MaxSize, VFileName)
    
    If hFile = 0 Then
            RaiseEvent OnVFileInitError(VFileName, "Unable to create virtual file")
            Exit Function
    End If
    
    If hFile <> 0 Then
        ExistFlag = False
        Wnd = FindWindow(WndClass, "Virtual File I/O Supervisor Window:[" & VFileName & "]")
        If GetLastError() = ERROR_ALREADY_EXISTS Then
            If Wnd = 0 Then
                RaiseEvent OnVFileInitError(VFileName, "Virtual file name conflicts with that of some Win32 Object")
                CloseHandle hFile
                Exit Function
            End If
        End If
        
        If Wnd = 0 Then RaiseEvent OnVFileCreate(VFileName)
        
        Set IOS = New IOSup
        Load IOS
        IOS.Caption = "Virtual File I/O Supervisor Window:[" & VFileName & "]"
        wndCAP = IOS.Caption
        IOS.SetParams Me
        If Wnd <> 0 Then
            Max = GetProp(Wnd, "Max")
            SetProp IOS.hwnd, "Max", Max
        Else
            Max = MaxSize
            SetProp IOS.hwnd, "Max", Max
                Dim TStr As String, Addr As Long
                TStr = Space$(Max)
                Addr = MapViewOfFile(hFile, FILE_MAP_WRITE, 0, 0, Max)
                WriteMem TStr, Addr
                UnmapViewOfFile Addr
        End If
        NStruct.wndCAP = VFileName
        vFName = VFileName
        InitializeVirtualFile = True
        RaiseEvent OnVFileInitSuccess(VFileName)
    End If
End Function

Function ReadVirtualFile(Optional AutoTrim As Boolean = True) As String
Attribute ReadVirtualFile.VB_Description = "Reads the contents of the associated virtual file."
    Dim Temp As String
    
    If hFile = 0 Then
        RaiseEvent OnVFileReadError(vFName, "Virtual File or Virtual File Interface not initialized")
        Exit Function
    End If
    
    If AutoTrim = True Then
        NoEvent = True
        ReadVirtualFile = Trim$(ReadVirtualFileEx(0, Max))
        If ReadVirtualFile = "" Then Exit Function
        If Asc(Right$(ReadVirtualFile, 1)) = 0 Then
            ReadVirtualFile = Left$(ReadVirtualFile, Len(ReadVirtualFile) - 1)
        End If
        RaiseEvent OnVFileReadSuccess(vFName, 0, Len(ReadVirtualFile))
    Else
        NoEvent = False
        ReadVirtualFile = ReadVirtualFileEx(0, Max)
    End If
    NoEvent = False
End Function

Function ReadVirtualFileEx(Offset As Long, nBytes As Long) As String
Attribute ReadVirtualFileEx.VB_Description = "Reads the contents of the specified virtual file (as per specified parameters)"
    If hFile = 0 Then
        RaiseEvent OnVFileReadError(vFName, "Virtual File or Virtual File Interface not initialized")
        Exit Function
    End If
    
    
    Dim Addr As Long
    If (Offset + nBytes) > Max Then
        RaiseEvent OnVFileReadError(vFName, "Invalid size and offset combination")
        Exit Function
    End If
    Addr = MapViewOfFile(hFile, FILE_MAP_WRITE, 0, 0, Max)
    If Addr = 0 Then
        RaiseEvent OnVFileReadError(vFName, "Unable to access shared memory area")
        Exit Function
    End If
    
    Dim Temp As String
    Temp = Space$(nBytes)
    ReadMem Temp, Addr + Offset
    ReadVirtualFileEx = Temp
    UnmapViewOfFile Addr
    If NoEvent = False Then RaiseEvent OnVFileReadSuccess(vFName, Offset, Len(Temp))
End Function

Function WriteFile(Offset As Long, bData As String, Optional ErasePrevData As Boolean = False) As Boolean
Attribute WriteFile.VB_Description = "Writes to the associated Virtual File"
    If hFile = 0 Then
        RaiseEvent OnVFileReadError(vFName, "Virtual File or Virtual File Interface not initialized")
        Exit Function
    End If


Dim Addr As Long
    If (Offset + Len(bData)) > Max Then
        RaiseEvent OnVFileWriteError(vFName, "Invalid data chunk size and offset combination")
        Exit Function
    End If
    If bData = "" Then
        RaiseEvent OnVFileWriteError(vFName, "Invalid data chunk")
        Exit Function
    End If
    Addr = MapViewOfFile(hFile, FILE_MAP_WRITE, 0, 0, Max)
    If Addr = 0 Then
        RaiseEvent OnVFileWriteError(vFName, "Unable to access shared memory area")
        Exit Function
    End If
    
    If ErasePrevData = True Then
        Dim Temp As String
        Temp = Space$(Max)
        WriteMem Temp, Addr
    End If
    
    WriteMem bData, Addr + Offset
    NStruct.nOffset = Offset
    NStruct.nSize = Len(bData)
    NStruct.wndCAP = wndCAP
    WriteFile = True
    UnmapViewOfFile Addr
    Call SendChangeSignal
    RaiseEvent OnVFileWriteSuccess(vFName, Offset, Len(bData))
End Function


Sub CleanUp()
Attribute CleanUp.VB_Description = "Closes the connection to the associated Virtual File.Also deletes the associated Virtual File if no other processes are using it."
    If hFile = 0 Then
        Exit Sub
    End If
    
    Dim Temp As String, Wnd As Long
    Temp = IOS.Caption
    Unload IOS
    
    Wnd = FindWindow(WndClass, Temp)
    If Wnd = 0 Then RaiseEvent OnVFileDestroy(vFName)
    Set IOS = Nothing
    wndCAP = ""
    CloseHandle hFile
    hFile = 0
    Max = 0
    vFName = ""
End Sub


Friend Function GetNotifyParams() As NotifyInfo
    GetNotifyParams = NStruct
End Function

Friend Sub SendChangeSignal()
    EnumWindows AddressOf EnumWndProc, ObjectPtr(Me)
End Sub

Friend Sub SetExistFlag()
    ExistFlag = True
End Sub

Function DoesFileExist(VFileName As String) As Boolean
Attribute DoesFileExist.VB_Description = "Returns True if a specified Virtual File Exists"
    VFileName = UCase$(VFileName)
    ExistFlag = False
    Wnd = FindWindow(WndClass, "Virtual File I/O Supervisor Window:[" & VFileName & "]")
    DoesFileExist = CBool(Wnd)
End Function

Public Property Get VirtualFileName() As String
Attribute VirtualFileName.VB_Description = "Read only property that returns the name of the currently associated virtual file name"
    VirtualFileName = NStruct.wndCAP
End Property

Public Property Get FileHandle() As Long
Attribute FileHandle.VB_Description = "Returns the Win32 Handle of the associated Virtual File"
    FileHandle = hFile
End Property

Private Sub Class_Initialize()
    Dim DForm As New IOSup
    Load IOSup
    WndClass = Space$(255)
    GetClassName DForm.hwnd, WndClass, Len(WndClass)
    WndClass = Trim$(WndClass)
    Unload DForm
    Set DForm = Nothing
End Sub

Private Sub Class_Terminate()
    Call CleanUp
End Sub
