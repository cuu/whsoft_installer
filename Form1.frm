VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "第一步"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      Picture         =   "Form1.frx":058A
      ScaleHeight     =   1455
      ScaleWidth      =   12735
      TabIndex        =   3
      Top             =   0
      Width           =   12735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "卸载以前的安装"
      Height          =   735
      Left            =   1560
      TabIndex        =   2
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   735
      Left            =   6960
      TabIndex        =   1
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "下一步"
      Height          =   735
      Left            =   4320
      TabIndex        =   0
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "安装之前请务必关闭运行中的MT4外汇交易平台软件!"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   2400
      Width           =   7695
   End
   Begin VB.Label Label1 
      Caption         =   "labe1"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function reinit()
    Dim a As String * 256
    Dim n As Integer
    Dim aa As String
    Dim x As Integer
    
    n = GetWindowsDirectory(a, 256)
    If Dir(a + "\system32\comdlg32.ocx") <> "" Then
    Else
    aa = "Regsvr32.exe /s " & App.path & "\comdlg32.ocx"
    x = Shell(aa$, 3)
    End If
    
End Function


Private Sub Command1_Click()
    Form1.Hide
    Form2.Show
    
End Sub

Private Sub Command2_Click()
    If MsgBox("请确认是否退出安装?", vbYesNo) = vbYes Then End
End Sub
Function delete_all()

Dim i As Integer
Dim vs As Variant

 vs = GetRegistryValue(HKEY_CURRENT_USER, "software\GoldRockfx Software\Info", "Fn", 0)

 vs = clean_ascii(CStr(vs))
  'MsgBox Len(vs) & " " & vs
  
 For i = 0 To 2
    If Dir(vs + "\files\" + Files(i)) <> "" Then
    'MsgBox Len(vs) & " " & InStrRev(vs, "\", 1)
    Kill vs + "\files\" + Files(i)
    End If
 Next i
 If Dir(vs + "\libraries\" + Files(3)) <> "" Then
    Kill vs + "\libraries\" + Files(3)
 End If
 If Dir(vs + "\indicators\" + Files(4)) <> "" Then
    Kill vs + "\indicators\" + Files(4)
 End If
 If Dir(vs + "\" + Files(4)) <> "" Then
    Kill vs + "\" + Files(4)
 End If
 'Call DeleteRegistryValueOrKey(HKEY_CURRENT_USER, "software\GoldRockfx Software\Info", "Sn")
 'Call DeleteRegistryValueOrKey(HKEY_CURRENT_USER, "software\GoldRockfx Software\Info", "Fn")
 MsgBox "卸载成功"
 
End Function

Private Sub Command3_Click()
'Dim vs As Variant
'Form6.Show
'Form1.Hide
'vs = GetRegistryValue(HKEY_CURRENT_USER, "software\MetaQuotes Software\MetaTrader 4", "InstallPath", 0)
'vs = clean_ascii(CStr(vs))
'Form6.Text1.Text = vs
Dim vs As Variant

'If MsgBox("请关闭MT外汇交易平台", vbYesNo) = vbYes Then

vs = GetRegistryValue(HKEY_CURRENT_USER, "software\MetaQuotes Software\MetaTrader 4", "InstallPath", 0)
If TypeName(vs) <> "String" Then
    MsgBox "您没有安装 Meta Trader外汇交易平台，请先安装外汇交易平台，然后重新安装本插件"
    End
Else
vs = GetRegistryValue(HKEY_CURRENT_USER, "software\GoldRockfx Software\Info", "Sn", 0)
If TypeName(vs) = "String" Then
    If MsgBox("是否删除本插件?", vbYesNo) = vbYes Then
        Call delete_all
    Else
    End If
End If
End If


End Sub
Function check_gk4()
    Dim ret As Long
    Dim cmd_line As String
    
    ret = GetPID("gk4.exe")
    
    
    If ret <> 0 And ret <> -1 Then
        cmd_line = "/F /pid " + CStr(ret)
        
        ShellExecute 0, "open", "taskkill", cmd_line, 0, 0
        
    End If
End Function

Private Sub Form_Load()
    Dim n As Integer
    Dim str As String
    Dim i As Integer
    
    'Call reinit
    Call init_files
    'MsgBox App.path
    n = InStrRev(Files(4), ".")
    For i = 1 To n - 1
        str = str & Mid(Files(4), i, 1)
    Next i
    
    str = str & " COPYRIGHT 2011 GUOREN"
    
    
   Label1.Caption = str
    Call check_gk4
End Sub
Private Sub Form_unLoad(cancel As Integer)
End

End Sub

