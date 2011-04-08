Attribute VB_Name = "Module4"

Public Type PROCESSENTRY32
  size As Long
  usage As Long
  processId As Long
  defaultHeapId As Long
  moduleId As Long
  cntThreads As Long
  parentProcessId As Long
  classBase As Long
  flags As Long
  exeFile As String * 260
End Type

Const INVALID_HANDLE_VALUE As Long = -1
Const TH32CS_SNAPPROCESS As Long = 2

Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal flags As Long, ByVal processId As Long) As Long
Public Declare Function Process32First Lib "kernel32.dll" (ByVal hSnapshot As Long, processEntry As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32.dll" (ByVal hSnapshot As Long, processEntry As PROCESSENTRY32) As Long


Public Function GetPID(ByVal pname As String) As Long

    Dim lReturnID     As Long
    Dim hSnap  As Long
    Dim proc32 As PROCESSENTRY32
    Dim iProcesses    As Long
    Dim vs As String
    iProcesses = 0

    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If hSnap = INVALID_HANDLE_VALUE Then
        GetPID = -1
    Else
        proc32.size = Len(proc32)
        lReturnID = Process32First(hSnap, proc32)
        Do While lReturnID
                If StrComp(getNullTerminatedString(proc32.exeFile), pname, 0) = 0 Then
                    iProcesses = proc32.processId
                    Exit Do
                End If
            lReturnID = Process32Next(hSnap, proc32)
        Loop
        Call CloseHandle(hSnap)
        GetPID = iProcesses
        
    End If
End Function

Private Function getNullTerminatedString(ByRef str As String) As String
Dim i As Long
  Let i = InStr(str, vbNullChar)
  Let getNullTerminatedString = Left$(str, IIf(i = 0, Len(str), i - 1))
End Function

