VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Go to article page.."
      Height          =   300
      Left            =   105
      TabIndex        =   12
      Top             =   5940
      Width           =   6735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Status:"
      Height          =   675
      Left            =   120
      TabIndex        =   8
      Top             =   4980
      Width           =   6720
      Begin VB.Label lWrite 
         Caption         =   "Last Write Access:"
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   375
         Width           =   6300
      End
      Begin VB.Label VName 
         AutoSize        =   -1  'True
         Caption         =   "Virtual Filename Name:"
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   180
         Width           =   1620
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1725
      Left            =   105
      TabIndex        =   4
      Top             =   2745
      Width           =   6735
      Begin VB.Label Label4 
         Caption         =   "This can be easily verified by running another instance of this app, and writing to the virtual file from it !"
         Height          =   390
         Left            =   75
         TabIndex        =   7
         Top             =   1155
         Width           =   6495
      End
      Begin VB.Label Label3 
         Caption         =   "Any writes to the current Virtual File by any other app (or other instances of this app) will be detected automaticslly"
         Height          =   390
         Left            =   75
         TabIndex        =   6
         Top             =   750
         Width           =   6555
      End
      Begin VB.Label Label2 
         Caption         =   $"dForm.frx":0000
         Height          =   600
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   6600
      End
   End
   Begin VB.CommandButton Write 
      Caption         =   "&Write To Virtual File"
      Height          =   300
      Left            =   2145
      TabIndex        =   2
      Top             =   4590
      Width           =   1935
   End
   Begin VB.CommandButton Read 
      Caption         =   "&Read from Virtual File"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   4590
      Width           =   1920
   End
   Begin VB.TextBox T1 
      Height          =   2190
      Left            =   165
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   465
      Width           =   6675
   End
   Begin VB.Label Label5 
      Caption         =   "If you found the code useful or interesting, please vote !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   5700
      Width           =   6735
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   180
      Top             =   270
      Width           =   6705
   End
   Begin VB.Label Label1 
      Caption         =   "Welcome to the vFile32 Virtual File Component Demonstration !"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   180
      TabIndex        =   3
      Top             =   45
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Const SW_SHOWNORMAL = 1
Const CodeID = 34324

Dim WithEvents VF As vFile
Attribute VF.VB_VarHelpID = -1

Private Sub Command1_Click()
    GotoURL ("http://planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=34324")
End Sub

Private Sub Form_Load()
Set VF = New vFile
Call VF.InitializeVirtualFile("A Trial 2.txt", 1024)
VName.Caption = "Virtual File Name:" & VF.VirtualFileName
End Sub

Private Sub Form_Unload(Cancel As Integer)
VF.CleanUp
Set VF = Nothing
End Sub

Private Sub Read_Click()
Dim T As String
T = VF.ReadVirtualFile()
MsgBox T, , "Virtual File Contents"
End Sub

Private Sub VF_OnVFileChange(ByVal VFileName As String, ByVal Offset As Long, ByVal Size As Long)
lWrite.Caption = "Last Write Access at Offset " & CStr(Offset) & " (" & CStr(Size) & " Bytes Written )"
End Sub

Private Sub VF_OnVFileCreate(ByVal VFileName As String)
    MsgBox "Virtual File successfully created", , VFileName
End Sub

Private Sub VF_OnVFileDestroy(ByVal VFileName As String)
    MsgBox "Virtual File successfully destroyed", , VFileName
End Sub

Private Sub VF_OnVFileInitError(ByVal VFileName As String, ByVal ErrDesc As String)
    MsgBox "Error Initializing Virtual file", , VFileName
    VF.CleanUp
    Set VF = Nothing
    Unload Me
End Sub

Private Sub VF_OnVFileInitSuccess(ByVal VFileName As String)
    MsgBox "Successfully initialized Virtual File Interface", , VFileName
End Sub

Private Sub VF_OnVFileReadError(ByVal VFileName As String, ByVal Reason As String)
    MsgBox "Error reading virtual file" & Chr$(13) & "Reason:" & Reason, , VFileName
End Sub

Private Sub VF_OnVFileReadSuccess(ByVal VFileName As String, ByVal Offset As Long, ByVal BytesRead As Long)
    MsgBox "Successfully read from virtual file" & Chr$(13) & "Offset:" & Offset & Chr$(13) & "Bytes Read:" & BytesRead, , VFileName
End Sub

Private Sub VF_OnVFileWriteError(ByVal VFileName As String, ByVal Reason As String)
    MsgBox "Error writing to virtual file" & Chr$(13) & "Reason:" & Reason, , VFileName
End Sub

Private Sub VF_OnVFileWriteSuccess(ByVal VFileName As String, ByVal Offset As Long, ByVal BytesWritten As Long)
    MsgBox "Successfully wrote to virtual file" & Chr$(13) & "Offset:" & Offset & Chr$(13) & "Bytes written:" & BytesWritten, , VFileName
End Sub

Private Sub Write_Click()
VF.WriteFile 0, T1.Text
End Sub

Sub GotoURL(URL As String)
    Dim Res As Long
    Dim TFile As String, Browser As String, Dum As String
    
    TFile = App.Path + "\test.htm"
    Open TFile For Output As #1
    Close
    Browser = String(255, " ")
    Res = FindExecutable(TFile, Dum, Browser)
    Browser = Trim$(Browser)
    
    If Len(Browser) = 0 Then
        MsgBox "Cannot find browser"
        Exit Sub
    End If
    
    Res = ShellExecute(Me.hwnd, "open", Browser, URL, Dum, SW_SHOWNORMAL)
    If Res <= 32 Then
        MsgBox "Cannot open web page"
        Exit Sub
    End If
End Sub



