Attribute VB_Name = "vFileRT"

Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Sub ReadMem Lib "MemHlp.dll" (ByVal St As String, ByVal Addr As Long)
Public Declare Sub WriteMem Lib "MemHlp.dll" (ByVal St As String, ByVal Addr As Long)
Public Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes As SECURITY_ATTRIBUTES, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Public Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Public Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Public Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
Public Declare Function FlushViewOfFile Lib "kernel32" (lpBaseAddress As Any, ByVal dwNumberOfBytesToFlush As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long

Public Const ERROR_ALREADY_EXISTS = 183&


Public Const WM_SIZE = &H5

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const PAGE_READWRITE = &H4
Public Const SECTION_MAP_WRITE = &H2
Public Const FILE_MAP_WRITE = SECTION_MAP_WRITE

Public Type NotifyInfo
    wndCAP As String
    nOffset As Long
    nSize As Long
End Type

Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type


Public Function GetObj(Ptr As Long) As Object
'Retrieves an Object from its pointer
Dim TObj As Object
CopyMemory TObj, Ptr, 4
Set GetObj = TObj
CopyMemory TObj, 0&, 4
End Function

Public Function ObjectPtr(Obj As Object) As Long
'Returns a pointer to an object
Dim lpObj As Long
CopyMemory lpObj, Obj, 4
ObjectPtr = lpObj
End Function


Function EnumWndProc(ByVal hwnd As Long, ByVal Ptr As Long) As Boolean
    Dim pObj As vFile, aWndName As String
    Dim WLen As Long, NInfo As NotifyInfo
    Set pObj = GetObj(Ptr)
    
    NInfo = pObj.GetNotifyParams()
    Set pObj = Nothing
    
    WLen = GetWindowTextLength(hwnd)
    If WLen = Len(NInfo.wndCAP) Then
        aWndName = Space$(WLen + 1)
        GetWindowText hwnd, aWndName, WLen + 1
        If Left$(aWndName, Len(aWndName) - 1) = NInfo.wndCAP Then
            SetProp hwnd, "Offset", NInfo.nOffset
            SetProp hwnd, "Size", NInfo.nSize
            PostMessage hwnd, WM_SIZE, 0, 0
        End If
    End If
    EnumWndProc = True
End Function


Function EnumWndChkProc(ByVal hwnd As Long, ByVal Ptr As Long) As Boolean
    Dim pObj As vFile, aWndName As String
    Dim WLen As Long, NInfo As NotifyInfo
    Set pObj = GetObj(Ptr)
    
    NInfo = pObj.GetNotifyParams()
    
    WLen = GetWindowTextLength(hwnd)
    If WLen = Len(NInfo.wndCAP) Then
        aWndName = Space$(WLen + 1)
        GetWindowText hwnd, aWndName, WLen + 1
        If Left$(aWndName, Len(aWndName) - 1) = NInfo.wndCAP Then
            Call pObj.SetExistFlag
        End If
    End If
    
    Set pObj = Nothing
    EnumWndChkProc = True
End Function
