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
   
''''//注册表 API 函数声明
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

''''//注册表访问权
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = &H3F

''''//打开/建立选项
Const REG_OPTION_NON_VOLATILE = 0&
Const REG_OPTION_VOLATILE = &H1

''''//Key 创建/打开
Const REG_CREATED_NEW_KEY = &H1
Const REG_OPENED_EXISTING_KEY = &H2

''''//预定义存取类型
Const STANDARD_RIGHTS_ALL = &H1F0000
Const SPECIFIC_RIGHTS_ALL = &HFFFF

''''//严格代码定义
Const ERROR_SUCCESS = 0&
Const ERROR_ACCESS_DENIED = 5
Const ERROR_NO_MORE_ITEMS = 259
Const ERROR_MORE_DATA = 234 ''''//  错误

''''//注册表值类型列举
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
   
''''//注册表基本键值列表
Public Enum RootKeyEnum
   HKEY_CLASSES_ROOT = &H80000000
   HKEY_CURRENT_USER = &H80000001
   HKEY_LOCAL_MACHINE = &H80000002
   HKEY_USERS = &H80000003
   HKEY_PERFORMANCE_DATA_WIN2K_ONLY = &H80000004 ''''//仅Win2k
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

''''//保存时指定类型
Public Enum RegFlags
   IsExpandableString = 1
   IsMultiString = 2
   ''''IsBigEndian = 3 ''''// 无指针同样不要设置大Endian值
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
   
   Dim SecAttr As SECURITY_ATTRIBUTES ''''//键的安全设置
   ''''//设置新键值的名称和默认安全设置
   SecAttr.nLength = Len(SecAttr) ''''//结构大小
   SecAttr.lpSecurityDescriptor = 0 ''''//默认安全权限
   SecAttr.bInheritHandle = True ''''//设置的默认值

   ''''// 打开或创建键
   ''''If RegOpenKeyEx(hKey, KeyName, 0, KEY_ALL_ACCESS, handle) Then Exit Function
   RetVal = RegCreateKeyEx(hKey, KeyName, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SecAttr, handle, RetVal)
   If RetVal Then Exit Function

   ''''//3种数据类型
   Select Case VarType(Value)
      Case vbByte, vbInteger, vbLong ''''// 若是字节, Integer值或Long值...
         lngValue = Value
         RetVal = RegSetValueExLong(handle, ValueName, 0, REG_DWORD, lngValue, Len(lngValue))
      
      Case vbString ''''// 字符串, 扩展环境字符串或多段字符串...
         strValue = Value
         Select Case Flag
            Case IsExpandableString
               RetVal = RegSetValueEx(handle, ValueName, 0, REG_EXPAND_SZ, ByVal strValue, 255)
            Case IsMultiString
               RetVal = RegSetValueEx(handle, ValueName, 0, REG_MULTI_SZ, ByVal strValue, 255)
            Case Else ''''// 正常 REG_SZ 字符串
               RetVal = RegSetValueEx(handle, ValueName, 0, REG_SZ, ByVal strValue, 255)
         End Select
      
      Case vbArray + vbByte ''''// 如果是字节数组...
         binValue = Value
         length = UBound(binValue) - LBound(binValue) + 1
         RetVal = RegSetValueExByte(handle, ValueName, 0, REG_BINARY, binValue(0), length)
      
      Case Else ''''// 如果其它类型
         RegCloseKey handle
         ''''Err.Raise 1001, , "不支持的值类型"
   
   End Select

   ''''// 返回关闭结果
   RegCloseKey handle
   
   ''''// 返回写入成功结果
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
   
   ''''// 默认结果
   GetRegistryValue = IIf(IsMissing(DefaultValue), Empty, DefaultValue)
   
   ''''// 打开键, 不存在则退出
   If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
   
   ''''// 准备 1K  resBinary 用于接收
   length = 1024
   ReDim resBinary(0 To length - 1) As Byte
   
   ''''// 读注册表值
   RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), length)
   
   ''''// 若resBinary 太小则重读
   If RetVal = ERROR_MORE_DATA Then
      ''''// resBinary放大,且重新读取
      ReDim resBinary(0 To length - 1) As Byte
      RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
      length)
   End If
   
   ''''// 返回相应值类型
   Select Case valueType
      Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN
         ''''// REG_DWORD 和 REG_DWORD_LITTLE_ENDIAN 相同
         CopyMemory resLong, resBinary(0), 4
         GetRegistryValue = resLong
      
      Case REG_DWORD_BIG_ENDIAN
         ''''// Big Endian''''s 用在非-Windows环境, 如Unix系统, 本地计算机远程访问
         CopyMemory resLong, resBinary(0), 4
         GetRegistryValue = SwapEndian(resLong)
      
      Case REG_SZ, REG_EXPAND_SZ
         resString = Space$(length - 1)
         CopyMemory ByVal resString, resBinary(0), length - 1
         If valueType = REG_EXPAND_SZ Then
            ''''// 查询对应的环境变量
            GetRegistryValue = ExpandEnvStr(resString)
         Else
            GetRegistryValue = resString
         End If

      Case REG_MULTI_SZ
         ''''// 复制时需指定2个空格符
         resString = Space$(length - 2)
         CopyMemory ByVal resString, resBinary(0), length - 2
         GetRegistryValue = resString

      Case Else '''' 包含 REG_BINARY
         ''''// resBinary 调整
         If length <> UBound(resBinary) + 1 Then
            ReDim Preserve resBinary(0 To length - 1) As Byte
         End If
      GetRegistryValue = resBinary()
   
   End Select
   
   ''''// 关闭
   RegCloseKey handle

End Function


Public Function DeleteRegistryValueOrKey(ByVal hKey As RootKeyEnum, RegKeyName As String, _
   ValueName As String) As Boolean
''''//删除注册表值和键,如果成功返回True

   Dim lRetval As Long      ''''//打开和输出注册表键的返回值
   Dim lRegHWND As Long     ''''//打开注册表键的句柄
   Dim sREGSZData As String ''''//把获取值放入缓冲区
   Dim lSLength As Long     ''''//缓冲区大小.  改变缓冲区大小要在调用之后
   
   ''''//打开键
   lRetval = RegOpenKeyEx(hKey, RegKeyName, 0, KEY_ALL_ACCESS, lRegHWND)
   
   ''''//成功打开
   If lRetval = ERR_NONE Then
      ''''//删除指定值
      lRetval = RegDeleteValue(lRegHWND, ValueName)  ''''//如果已存在则先删除
      
      ''''//如出现错误则删除值并返回False
      If lRetval <> ERR_NONE Then Exit Function
      
      ''''//注意: 如果成功打开仅关闭注册表键
      lRetval = RegCloseKey(lRegHWND)
     
      ''''//如成功关闭则返回 True 或者其它错误
      If lRetval = ERR_NONE Then DeleteRegistryValueOrKey = True
      
   End If

End Function


Public Function ExpandEnvStr(sData As String) As String
''''// 查询环境变量和返回定义值
''''// 如： %PATH% 则返回 "c:\;c:\windows;"

   Dim c As Long, s As String
   
   s = "" ''''// 不支持Windows 95
   
   ''''// get the length
   c = ExpandEnvironmentStrings(sData, s, c)
   
   ''''// 展开字符串
   s = String$(c - 1, 0)
   c = ExpandEnvironmentStrings(sData, s, c)
   
   ''''// 返回环境变量
   ExpandEnvStr = s
   
End Function


Public Function SwapEndian(ByVal dw As Long) As Long
''''// 转换大DWord 到小 DWord
   
   CopyMemory ByVal VarPtr(SwapEndian) + 3, dw, 1
   CopyMemory ByVal VarPtr(SwapEndian) + 2, ByVal VarPtr(dw) + 1, 1
   CopyMemory ByVal VarPtr(SwapEndian) + 1, ByVal VarPtr(dw) + 2, 1
   CopyMemory SwapEndian, ByVal VarPtr(dw) + 3, 1

End Function
''HKLM\SOFTWARE\GOLDROCKFX\RGNUMBER



