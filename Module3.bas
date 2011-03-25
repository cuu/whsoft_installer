Attribute VB_Name = "Module3"
Function clean_ascii(ByVal str As String) As String
    Dim res As String
    
    Dim strcmp As String
    Dim I As Integer
    res = ""
    
    For I = 1 To Len(str)
    
        strcmp = Mid(str, I, 1)
        If Asc(strcmp) >= 32 And Asc(strcmp) <= 126 Then
        res = res & strcmp
        Else
        clean_ascii = res
        Exit Function
        End If
        
    Next I
    clean_ascii = res
    
End Function

Public Function init_files()
    Dim DemoIni As New classIniFile
    
    Vers(0) = "金钥匙智能分析系统.ex4"
    Vers(1) = "金钥匙智能交易单向版.ex4"
    Vers(2) = "金钥匙智能交易趋势版.ex4"
    Vers(3) = "金钥匙智能交易双向版.ex4"
    

    Files(0) = "curl.exe"
    Files(1) = "libeay32.dll"
    Files(2) = "libssl32.dll"
    Files(3) = "goldkey.dll"
    
    Files(5) = "test3.dll"
    Files(6) = "gk4.exe"
    
    If Dir(App.path + "\cfg.ini") = "" Then
        MsgBox "Data not correct cfg.ini exit"
        End
    End If
    DemoIni.INIFileName = App.path & "\cfg.ini"
    ' switch versions
    Files(4) = Trim(DemoIni.GetIniKey("sys", "ext"))
    'MsgBox "files4  " + Files(4)
    Rjbb(0) = CInt(Trim(DemoIni.GetIniKey("sys", "ver")))
    
    'Server_ip(0) = "218.240.38.44"
    Server_ip(0) = Trim(DemoIni.GetIniKey("sys", "srv"))
    If Len(Files(4)) < 6 Or Len(Server_ip(0)) < 6 Then
    MsgBox "Data error"
    End
    End If
   
End Function

