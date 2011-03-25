VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "填写注册信息"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10155
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   10155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   2760
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "下一步"
      Height          =   495
      Left            =   7320
      TabIndex        =   9
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   2160
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   960
      Width           =   7935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label9 
      Caption         =   "具体编号可以询问代理商,必须填写正确"
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label Label8 
      Caption         =   "您的代理商编号"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "MT外汇平台账号"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "所有信息请真实填写，然后点击 下一步 "
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   4440
      Width           =   3735
   End
   Begin VB.Label Label5 
      Caption         =   "手机号码前不用加0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   10
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "移动电话"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "固定电话"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "联系地址"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "用户姓名"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim vs As Variant




If Len(Text1.Text) < 1 Then
MsgBox "请填写完整信息"
    Exit Sub
End If
If Len(Text2.Text) < 1 Then
MsgBox "请填写完整信息"
    Exit Sub
End If
If Len(Text3.Text) < 1 Then
MsgBox "请填写完整信息"
    Exit Sub
End If

If Len(Text4.Text) < 1 Then
MsgBox "请填写完整信息"
    Exit Sub
End If

If Len(Text5.Text) < 1 Then
MsgBox "请填写MT外汇平台账号"
    Exit Sub
End If

If Len(Text5.Text) < 6 Then
MsgBox "请正确填写代理商编号"
    Exit Sub
End If

Form3.Hide
Form4.Label6.Caption = Text1.Text
Form4.Label7.Caption = Text2.Text
Form4.Label8.Caption = Text3.Text
Form4.Label9.Caption = Text4.Text
Form4.Label11.Caption = Text6.Text

'Form4.Text1.Text = GetRegistryValue(HKEY_CURRENT_USER, "software\MetaQuotes Software\MetaTrader 4", "InstallPath", 0)
Form4.Show
End Sub

Private Sub Command2_Click()
Form3.Hide
Form2.Show

End Sub

Private Sub Form_unLoad(cancel As Integer)
End

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label8_Click()

End Sub

  Private Sub text3_KeyPress(KeyAscii As Integer)
          If Not ((Chr(KeyAscii) Like "[0-9]") Or KeyAscii = 8) Then
                  KeyAscii = 0
          End If
  End Sub
    Private Sub text4_KeyPress(KeyAscii As Integer)
          If Not ((Chr(KeyAscii) Like "[0-9]") Or KeyAscii = 8) Then
                  KeyAscii = 0
          End If
  End Sub
    Private Sub text5_KeyPress(KeyAscii As Integer)
          If Not ((Chr(KeyAscii) Like "[0-9]") Or KeyAscii = 8) Then
                  KeyAscii = 0
          End If
  End Sub
 Private Sub text6_KeyPress(KeyAscii As Integer)
          If Not ((Chr(KeyAscii) Like "[0-9]") Or KeyAscii = 8) Then
                  KeyAscii = 0
          End If
  End Sub

Private Sub text3_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
 If Button = vbRightButton Then
 
 Text3.Enabled = False
 Text3.Enabled = True
 End If
 
End Sub
