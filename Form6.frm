VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form6"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9780
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command3 
      Caption         =   "取消"
      Height          =   615
      Left            =   5880
      TabIndex        =   5
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "确定删除"
      Height          =   615
      Left            =   7680
      TabIndex        =   3
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "改变路径"
      Height          =   375
      Left            =   8280
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1920
      Width           =   8055
   End
   Begin VB.Label Label3 
      Caption         =   "卸载前请关闭正在使用的MT外汇交易平台软件"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   3000
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "请正确选择卸载路径，卸载"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "以前的安装路径"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function delete_all()

Dim i As Integer
Dim vs As Variant

 vs = GetRegistryValue(HKEY_CURRENT_USER, "software\GoldRockfx Software\Info", "Fn", 0)

 vs = clean_ascii(CStr(vs))
  'MsgBox Len(vs) & " " & vs
  
 For i = 0 To 2
    If Dir(vs + "\files\" + Files(i)) <> "" Then
    'MsgBox Len(vs) & " " & InStrRev(vs, "\", 1)
    'Kill vs + "\files\" + Files(i)
    End If
 Next i
 If Dir(vs + "\libraries\" + Files(3)) <> "" Then
    'Kill vs + "\libraries\" + Files(3)
 End If
 If Dir(vs + "\" + Files(4)) <> "" Then
    Kill vs + "\" + Files(4)
 End If
 'Call DeleteRegistryValueOrKey(HKEY_CURRENT_USER, "software\GoldRockfx Software\Info", "Sn")
 'Call DeleteRegistryValueOrKey(HKEY_CURRENT_USER, "software\GoldRockfx Software\Info", "Fn")
 MsgBox "卸载成功"
 
End Function


Private Sub Command1_Click()
  Dim fDir As String
  
    fDir = get_open_dir
    If Len(fDir) > 1 Then
    
    Text1.Text = fDir
    End If
End Sub

Private Sub Command2_Click()
Dim vs As Variant

'If MsgBox("请关闭MT外汇交易平台", vbYesNo) = vbYes Then

vs = GetRegistryValue(HKEY_CURRENT_USER, "software\MetaQuotes Software\MetaTrader 4", "InstallPath", 0)
If TypeName(vs) <> "String" Then
    MsgBox "您没有安装 Meta Trader外汇交易平台，请先安装外汇交易平台，然后重新安装本插件"
    End
Else
vs = GetRegistryValue(HKEY_CURRENT_USER, "software\GoldRockfx Software\Info", "Sn", 0)
If TypeName(vs) = "String" Then
    'If MsgBox("是否删除本插件?", vbYesNo) = vbYes Then
        Call delete_all
    'Else
    End If
End If
'End If

End Sub

Private Sub Form_unLoad(cancel As Integer)
End

End Sub
