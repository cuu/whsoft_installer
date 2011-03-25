Attribute VB_Name = "Module2"

Private Declare Function CreatePipe Lib "kernel32.dll" (ByRef phReadPipe As Long, ByRef phWritePipe As Long, ByRef lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Declare Function CreateProcess Lib "kernel32.dll" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByRef lpProcessAttributes As Long, ByRef lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByRef lpEnvironment As Any, ByVal lpCurrentDriectory As String, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Long) As Long

Private Const STARTF_USESHOWWINDOW As Long = &H1
Private Const STARTF_USESTDHANDLES As Long = &H100
Private Const SW_HIDE As Long = 0

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Public Function ExecAndCapture(ByVal sCommandLine As String, Optional ByVal sStartInFolder As String = vbNullString) As String


Const BUFSIZE As Long = 1024 * 10
Dim hPipeRead As Long
Dim hPipeWrite As Long
Dim sa As SECURITY_ATTRIBUTES
Dim si As STARTUPINFO
Dim pi As PROCESS_INFORMATION
Dim baOutput(BUFSIZE) As Byte
Dim sOutput As String
Dim lBytesRead As Long

    With sa
        .nLength = Len(sa)
        .bInheritHandle = 1    ' get inheritable pipe handles
    End With
    
    If CreatePipe(hPipeRead, hPipeWrite, sa, 0) = 0 Then
        Exit Function
    End If
    
    With si
        .cb = Len(si)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = SW_HIDE          ' hide the window
        .hStdOutput = hPipeWrite
        .hStdError = hPipeWrite
    End With
    
    If CreateProcess(vbNullString, sCommandLine, ByVal 0&, ByVal 0&, 1, 0&, ByVal 0&, sStartInFolder, si, pi) Then
        Call CloseHandle(hPipeWrite)
        Call CloseHandle(pi.hThread)
        hPipeWrite = 0
        Do
            DoEvents
            If ReadFile(hPipeRead, baOutput(0), BUFSIZE, lBytesRead, ByVal 0&) = 0 Then
                Exit Do
            End If
            sOutput = Left$(StrConv(baOutput(), vbUnicode), lBytesRead)
        Loop
        

        Call CloseHandle(pi.hProcess)
    End If
    
    Call CloseHandle(hPipeRead)
    Call CloseHandle(hPipeWrite)
    
        '--
        ExecAndCapture = (StrConv(baOutput(), vbUnicode))
        '--
    
End Function
Public Function RunCommand(ByVal str As String)

      RunCommand = ExecAndCapture(str)
    
End Function

