Attribute VB_Name = "Module1"
Option Explicit

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
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
    lpReserved2 As Long
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
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Const NORMAL_PRIORITY_CLASS As Long = &H20&
Private Const STARTF_USESTDHANDLES As Long = &H100&
Private Const STARTF_USESHOWWINDOW As Long = &H1&
Private Const SW_HIDE As Long = 0&
Private Const INFINITE As Long = &HFFFF&

Private Type OPENFILENAME
   lStructSize As Long
   hwndOwner As Long
   hInstance As Long
   lpstrFilter As String
   lpstrCustomFilter As String
   nMaxCustFilter As Long
   nFilterIndex As Long
   lpstrFile As String
   nMaxFile As Long
   lpstrFileTitle As String
   nMaxFileTitle As Long
   lpstrInitialDir As String
   lpstrTitle As String
   flags As Long
   nFileOffset As Integer
   nFileExtension As Integer
   lpstrDefExt As String
   lCustData As Long
   lpfnHook As Long
   lpTemplateName As String
End Type
Public Declare Function GetLastError Lib "kernel32" () As Long

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long


'---------------------------------------------------------------------------------------------------------------------------------

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
   (dest As Any, source As Any, ByVal numBytes As Long)

Public Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" _
   (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
   
''''//ע��� API ��������
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal lpReserved As Long, lpType As Long, lpData As Any, _
    lpcbData As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" _
   (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, _
   ByVal lpbData As Any, ByVal cbData As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" _
   (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
   ByVal lpClass As String, ByVal dwOptions As Long, _
   ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, _
   phkResult As Long, lpdwDisposition As Long) As Long

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
   (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
   lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, _
   lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
   (ByVal hKey As Long, ByVal dwIndex As Long, _
   ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, _
   lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
   (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
   (ByVal hKey As Long, ByVal lpValueName As String) As Long

Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" _
   (ByVal hKey As Long, ByVal ipValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, _
   ByVal lpValue As String, ByVal cbData As Long) As Long

Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" _
   (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, _
   lpValue As Long, ByVal cbData As Long) As Long

Private Declare Function RegSetValueExByte Lib "advapi32.dll" Alias "RegSetValueExA" _
   (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, _
   lpValue As Byte, ByVal cbData As Long) As Long

Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" _
   (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, _
   ByVal lpReserved As Long, lpcSubKeys As Long, _
   lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, _
   lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, _
   lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long

Private Declare Function RegEnumValueInt Lib "advapi32.dll" Alias "RegEnumValueA" _
   (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
   lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
   lpData As Byte, lpcbData As Long) As Long

Private Declare Function RegEnumValueStr Lib "advapi32.dll" Alias "RegEnumValueA" _
   (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
   lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
   lpData As Byte, lpcbData As Long) As Long

Private Declare Function RegEnumValueByte Lib "advapi32.dll" Alias "RegEnumValueA" _
   (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
   lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
   lpData As Byte, lpcbData As Long) As Long
   
Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

''''//ע������Ȩ
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = &H3F

''''//��/����ѡ��
Const REG_OPTION_NON_VOLATILE = 0&
Const REG_OPTION_VOLATILE = &H1

''''//Key ����/��
Const REG_CREATED_NEW_KEY = &H1
Const REG_OPENED_EXISTING_KEY = &H2

''''//Ԥ�����ȡ����
Const STANDARD_RIGHTS_ALL = &H1F0000
Const SPECIFIC_RIGHTS_ALL = &HFFFF

''''//�ϸ���붨��
Const ERROR_SUCCESS = 0&
Const ERROR_ACCESS_DENIED = 5
Const ERROR_NO_MORE_ITEMS = 259
Const ERROR_MORE_DATA = 234 ''''//  ����

''''//ע���ֵ�����о�
Private Enum RegDataTypeEnum
''''   REG_NONE = (0)                         ''''// No value type
   REG_SZ = (1)                           ''''// Unicode nul terminated string
   REG_EXPAND_SZ = (2)                    ''''// Unicode nul terminated string w/enviornment var
   REG_BINARY = (3)                       ''''// Free form binary
   REG_DWORD = (4)                        ''''// 32-bit number
   REG_DWORD_LITTLE_ENDIAN = (4)          ''''// 32-bit number (same as REG_DWORD)
   REG_DWORD_BIG_ENDIAN = (5)             ''''// 32-bit number
''''   REG_LINK = (6)                         ''''// Symbolic Link (unicode)
   REG_MULTI_SZ = (7)                     ''''// Multiple, null-delimited, double-null-terminated Unicode strings
''''   REG_RESOURCE_LIST = (8)                ''''// Resource list in the resource map
''''   REG_FULL_RESOURCE_DESCRIPTOR = (9)     ''''// Resource list in the hardware description
''''   REG_RESOURCE_REQUIREMENTS_LIST = (10)
End Enum
   
''''//ע��������ֵ�б�
Public Enum RootKeyEnum
   HKEY_CLASSES_ROOT = &H80000000
   HKEY_CURRENT_USER = &H80000001
   HKEY_LOCAL_MACHINE = &H80000002
   HKEY_USERS = &H80000003
   HKEY_PERFORMANCE_DATA_WIN2K_ONLY = &H80000004 ''''//��Win2k
   HKEY_CURRENT_CONFIG = &H80000005
   HKEY_DYN_DATA = &H80000006
End Enum

''''// for specifying the type of data to save
Public Enum RegValueTypes
   eInteger = vbInteger
   eLong = vbLong
   eString = vbString
   eByteArray = vbArray + vbByte
End Enum

''''//����ʱָ������
Public Enum RegFlags
   IsExpandableString = 1
   IsMultiString = 2
   ''''IsBigEndian = 3 ''''// ��ָ��ͬ����Ҫ���ô�Endianֵ
End Enum

'
'---------------------------------------------------------------------------
' Used to get the MAC address.
'---------------------------------------------------------------------------
'
Private Const NCBNAMSZ As Long = 16
Private Const NCBENUM As Long = &H37
Public Const NCBRESET As Long = &H32
Public Const NCBASTAT As Long = &H33
Public Const HEAP_ZERO_MEMORY As Long = &H8
Public Const HEAP_GENERATE_EXCEPTIONS As Long = &H4

Public Type NET_CONTROL_BLOCK  'NCB
    ncb_command    As Byte
    ncb_retcode    As Byte
    ncb_lsn        As Byte
    ncb_num        As Byte
    ncb_buffer     As Long
    ncb_length     As Integer
    ncb_callname   As String * NCBNAMSZ
    ncb_name       As String * NCBNAMSZ
    ncb_rto        As Byte
    ncb_sto        As Byte
    ncb_post       As Long
    ncb_lana_num   As Byte
    ncb_cmd_cplt   As Byte
    ncb_reserve(9) As Byte 'Reserved, must be 0
    ncb_event      As Long
End Type

Public Type ADAPTER_STATUS
    adapter_address(5) As Byte
    rev_major          As Byte
    reserved0          As Byte
    adapter_type       As Byte
    rev_minor          As Byte
    duration           As Integer
    frmr_recv          As Integer
    frmr_xmit          As Integer
    iframe_recv_err    As Integer
    xmit_aborts        As Integer
    xmit_success       As Long
    recv_success       As Long
    iframe_xmit_err    As Integer
    recv_buff_unavail  As Integer
    t1_timeouts        As Integer
    ti_timeouts        As Integer
    Reserved1          As Long
    free_ncbs          As Integer
    max_cfg_ncbs       As Integer
    max_ncbs           As Integer
    xmit_buf_unavail   As Integer
    max_dgram_size     As Integer
    pending_sess       As Integer
    max_cfg_sess       As Integer
    max_sess           As Integer
    max_sess_pkt_size  As Integer
    name_count         As Integer
End Type

Public Type NAME_BUFFER
    name_(0 To NCBNAMSZ - 1) As Byte
    name_num                 As Byte
    name_flags               As Byte
End Type

Public Type ASTAT
    adapt             As ADAPTER_STATUS
    NameBuff(0 To 29) As NAME_BUFFER
End Type

Public Declare Function Netbios Lib "netapi32" _
        (pncb As NET_CONTROL_BLOCK) As Byte

Public Declare Function GetProcessHeap Lib "kernel32" () As Long

Public Declare Function HeapAlloc Lib "kernel32" _
        (ByVal hHeap As Long, ByVal dwFlags As Long, _
        ByVal dwBytes As Long) As Long
     
Public Declare Function HeapFree Lib "kernel32" _
        (ByVal hHeap As Long, ByVal dwFlags As Long, _
        lpMem As Any) As Long

Public Const ERR_NONE = 0
Public Files(8) As String
Public Vers(8) As String
Public Server_ip(1) As String
Public Rjbb(4) As Integer





Public Function OpenFile(WinHwnd As Long, _
   Optional BoxLabel As String = "", _
   Optional StartPath As String = "", _
   Optional FilterStr = "*.*|*.*", _
   Optional Flag As Variant = &H8 Or &H200000) As String
   
   
   Dim Rc As Long
   Dim pOpenfilename As OPENFILENAME
   Dim Fstr1() As String
   Dim Fstr As String
   Dim I As Long
   Const MAX_Buffer_LENGTH = 256
   
   On Error Resume Next
   
   If Len(Trim$(StartPath)) > 0 Then
    If Right$(StartPath, 1) <> "\" Then StartPath = StartPath & "\"
        If Dir$(StartPath, vbDirectory + vbHidden) = "" Then
            StartPath = App.path
        End If
        Else
            StartPath = App.path
        End If
    If Len(Trim$(FilterStr)) = 0 Then
        Fstr = "*.*|*.*"
   End If
   Fstr = ""
   Fstr1 = Split(FilterStr, "|")
   For I = 0 To UBound(Fstr1)
   Fstr = Fstr & Fstr1(I) & vbNullChar
   Next
   With pOpenfilename
   .hwndOwner = WinHwnd
   .hInstance = App.hInstance
   .lpstrTitle = BoxLabel
   .lpstrInitialDir = StartPath
   .lpstrFilter = Fstr
   .nFilterIndex = 1
   .lpstrDefExt = vbNullChar & vbNullChar
   .lpstrFile = String(MAX_Buffer_LENGTH, 0)
   .nMaxFile = MAX_Buffer_LENGTH - 1
   .lpstrFileTitle = .lpstrFile
   .nMaxFileTitle = MAX_Buffer_LENGTH
   .lStructSize = Len(pOpenfilename)
   .flags = Flag
   End With
   Rc = GetOpenFileName(pOpenfilename)
   If Rc Then
   OpenFile = Left$(pOpenfilename.lpstrFile, pOpenfilename.nMaxFile)
   Else
   OpenFile = ""
   End If
End Function

Public Function SetRegistryValue(ByVal hKey As RootKeyEnum, ByVal KeyName As String, _
   ByVal ValueName As String, ByVal Value As Variant, valueType As RegValueTypes, _
   Optional Flag As RegFlags = 0) As Boolean
   
   Dim handle As Long
   Dim lngValue As Long
   Dim strValue As String
   Dim binValue() As Byte
   Dim length As Long
   Dim RetVal As Long
   
   Dim SecAttr As SECURITY_ATTRIBUTES ''''//���İ�ȫ����
   ''''//�����¼�ֵ�����ƺ�Ĭ�ϰ�ȫ����
   SecAttr.nLength = Len(SecAttr) ''''//�ṹ��С
   SecAttr.lpSecurityDescriptor = 0 ''''//Ĭ�ϰ�ȫȨ��
   SecAttr.bInheritHandle = True ''''//���õ�Ĭ��ֵ

   ''''// �򿪻򴴽���
   ''''If RegOpenKeyEx(hKey, KeyName, 0, KEY_ALL_ACCESS, handle) Then Exit Function
   RetVal = RegCreateKeyEx(hKey, KeyName, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SecAttr, handle, RetVal)
   If RetVal Then Exit Function

   ''''//3����������
   Select Case VarType(Value)
      Case vbByte, vbInteger, vbLong ''''// �����ֽ�, Integerֵ��Longֵ...
         lngValue = Value
         RetVal = RegSetValueExLong(handle, ValueName, 0, REG_DWORD, lngValue, Len(lngValue))
      
      Case vbString ''''// �ַ���, ��չ�����ַ��������ַ���...
         strValue = Value
         Select Case Flag
            Case IsExpandableString
               RetVal = RegSetValueEx(handle, ValueName, 0, REG_EXPAND_SZ, ByVal strValue, 255)
            Case IsMultiString
               RetVal = RegSetValueEx(handle, ValueName, 0, REG_MULTI_SZ, ByVal strValue, 255)
            Case Else ''''// ���� REG_SZ �ַ���
               RetVal = RegSetValueEx(handle, ValueName, 0, REG_SZ, ByVal strValue, 255)
         End Select
      
      Case vbArray + vbByte ''''// ������ֽ�����...
         binValue = Value
         length = UBound(binValue) - LBound(binValue) + 1
         RetVal = RegSetValueExByte(handle, ValueName, 0, REG_BINARY, binValue(0), length)
      
      Case Else ''''// �����������
         RegCloseKey handle
         ''''Err.Raise 1001, , "��֧�ֵ�ֵ����"
   
   End Select

   ''''// ���عرս��
   RegCloseKey handle
   
   ''''// ����д��ɹ����
   SetRegistryValue = (RetVal = 0)

End Function


Public Function GetRegistryValue(ByVal hKey As RootKeyEnum, ByVal KeyName As String, _
   ByVal ValueName As String, Optional DefaultValue As Variant) As Variant
   
   Dim handle As Long
   Dim resLong As Long
   Dim resString As String
   Dim resBinary() As Byte
   Dim length As Long
   Dim RetVal As Long
   Dim valueType As Long

   Const KEY_READ = &H20019
   
   ''''// Ĭ�Ͻ��
   GetRegistryValue = IIf(IsMissing(DefaultValue), Empty, DefaultValue)
   
   ''''// �򿪼�, ���������˳�
   If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
   
   ''''// ׼�� 1K  resBinary ���ڽ���
   length = 1024
   ReDim resBinary(0 To length - 1) As Byte
   
   ''''// ��ע���ֵ
   RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), length)
   
   ''''// ��resBinary ̫С���ض�
   If RetVal = ERROR_MORE_DATA Then
      ''''// resBinary�Ŵ�,�����¶�ȡ
      ReDim resBinary(0 To length - 1) As Byte
      RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
      length)
   End If
   
   ''''// ������Ӧֵ����
   Select Case valueType
      Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN
         ''''// REG_DWORD �� REG_DWORD_LITTLE_ENDIAN ��ͬ
         CopyMemory resLong, resBinary(0), 4
         GetRegistryValue = resLong
      
      Case REG_DWORD_BIG_ENDIAN
         ''''// Big Endian''''s ���ڷ�-Windows����, ��Unixϵͳ, ���ؼ����Զ�̷���
         CopyMemory resLong, resBinary(0), 4
         GetRegistryValue = SwapEndian(resLong)
      
      Case REG_SZ, REG_EXPAND_SZ
         resString = Space$(length - 1)
         CopyMemory ByVal resString, resBinary(0), length - 1
         If valueType = REG_EXPAND_SZ Then
            ''''// ��ѯ��Ӧ�Ļ�������
            GetRegistryValue = ExpandEnvStr(resString)
         Else
            GetRegistryValue = resString
         End If

      Case REG_MULTI_SZ
         ''''// ����ʱ��ָ��2���ո��
         resString = Space$(length - 2)
         CopyMemory ByVal resString, resBinary(0), length - 2
         GetRegistryValue = resString

      Case Else '''' ���� REG_BINARY
         ''''// resBinary ����
         If length <> UBound(resBinary) + 1 Then
            ReDim Preserve resBinary(0 To length - 1) As Byte
         End If
      GetRegistryValue = resBinary()
   
   End Select
   
   ''''// �ر�
   RegCloseKey handle

End Function


Public Function DeleteRegistryValueOrKey(ByVal hKey As RootKeyEnum, RegKeyName As String, _
   ValueName As String) As Boolean
''''//ɾ��ע���ֵ�ͼ�,����ɹ�����True

   Dim lRetval As Long      ''''//�򿪺����ע�����ķ���ֵ
   Dim lRegHWND As Long     ''''//��ע�����ľ��
   Dim sREGSZData As String ''''//�ѻ�ȡֵ���뻺����
   Dim lSLength As Long     ''''//��������С.  �ı仺������СҪ�ڵ���֮��
   
   ''''//�򿪼�
   lRetval = RegOpenKeyEx(hKey, RegKeyName, 0, KEY_ALL_ACCESS, lRegHWND)
   
   ''''//�ɹ���
   If lRetval = ERR_NONE Then
      ''''//ɾ��ָ��ֵ
      lRetval = RegDeleteValue(lRegHWND, ValueName)  ''''//����Ѵ�������ɾ��
      
      ''''//����ִ�����ɾ��ֵ������False
      If lRetval <> ERR_NONE Then Exit Function
      
      ''''//ע��: ����ɹ��򿪽��ر�ע����
      lRetval = RegCloseKey(lRegHWND)
     
      ''''//��ɹ��ر��򷵻� True ������������
      If lRetval = ERR_NONE Then DeleteRegistryValueOrKey = True
      
   End If

End Function


Public Function ExpandEnvStr(sData As String) As String
''''// ��ѯ���������ͷ��ض���ֵ
''''// �磺 %PATH% �򷵻� "c:\;c:\windows;"

   Dim c As Long, s As String
   
   s = "" ''''// ��֧��Windows 95
   
   ''''// get the length
   c = ExpandEnvironmentStrings(sData, s, c)
   
   ''''// չ���ַ���
   s = String$(c - 1, 0)
   c = ExpandEnvironmentStrings(sData, s, c)
   
   ''''// ���ػ�������
   ExpandEnvStr = s
   
End Function


Public Function SwapEndian(ByVal dw As Long) As Long
''''// ת����DWord ��С DWord
   
   CopyMemory ByVal VarPtr(SwapEndian) + 3, dw, 1
   CopyMemory ByVal VarPtr(SwapEndian) + 2, ByVal VarPtr(dw) + 1, 1
   CopyMemory ByVal VarPtr(SwapEndian) + 1, ByVal VarPtr(dw) + 2, 1
   CopyMemory SwapEndian, ByVal VarPtr(dw) + 3, 1

End Function
''HKLM\SOFTWARE\GOLDROCKFX\RGNUMBER



