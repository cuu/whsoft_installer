VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ע��"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command3 
      Caption         =   "���"
      Height          =   495
      Left            =   6240
      TabIndex        =   11
      Top             =   4800
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����ע��"
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   4800
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��һ��"
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "����ע��"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   375
         Left            =   1440
         TabIndex        =   14
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "�����̱��:"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   2280
         Width           =   3735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   480
         Width           =   6855
      End
      Begin VB.Label Label4 
         Caption         =   "�̶��绰��"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "�̶��绰��"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "��ϵ��ַ��"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "�û�������"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Label Label10 
      Caption         =   "����ϸ�˶�ע����Ϣ,ȷ������� [ ����ע�� ] ���ǵõ�����ע����Ϣ֮��ἰʱ������ϵ,����ͨ��Ӧ����"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   9015
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Form4.Hide
Form3.Show

End Sub

Public Function fGetMacAddress() As String
    Dim l As Long
    Dim lngError As Long
    Dim lngSize As Long
    Dim pAdapt As Long
    Dim pAddrStr As Long
    Dim pASTAT As Long
    Dim strTemp As String
    Dim strAddress As String
    Dim strMACAddress As String
    Dim AST As ASTAT
    Dim NCB As NET_CONTROL_BLOCK

    '
    '---------------------------------------------------------------------------
    ' Get the network interface card's MAC address.
    '----------------------------------------------------------------------------
    '
    On Error GoTo ErrorHandler
    fGetMacAddress = ""
    strMACAddress = ""

    '
    ' Try to get MAC address from NetBios. Requires NetBios installed.
    '
    ' Supported on 95, 98, ME, NT, 2K, XP
    '
    ' Results Connected Disconnected
    ' ------- --------- ------------
    '   XP       OK         Fail (Fail after reboot)
    '   NT       OK         OK   (OK after reboot)
    '   98       OK         OK   (OK after reboot)
    '   95       OK         OK   (OK after reboot)
    '
    NCB.ncb_command = NCBRESET
    Call Netbios(NCB)

    NCB.ncb_callname = "*               "
    NCB.ncb_command = NCBASTAT
    NCB.ncb_lana_num = 0
    NCB.ncb_length = Len(AST)

    pASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS Or _
                       HEAP_ZERO_MEMORY, NCB.ncb_length)
    If pASTAT = 0 Then GoTo ErrorHandler

    NCB.ncb_buffer = pASTAT
    Call Netbios(NCB)

    Call CopyMemory(AST, NCB.ncb_buffer, Len(AST))

    strMACAddress = Right$("00" & Hex(AST.adapt.adapter_address(0)), 2) & _
                    Right$("00" & Hex(AST.adapt.adapter_address(1)), 2) & _
                    Right$("00" & Hex(AST.adapt.adapter_address(2)), 2) & _
                    Right$("00" & Hex(AST.adapt.adapter_address(3)), 2) & _
                    Right$("00" & Hex(AST.adapt.adapter_address(4)), 2) & _
                    Right$("00" & Hex(AST.adapt.adapter_address(5)), 2)

    Call HeapFree(GetProcessHeap(), 0, pASTAT)

    fGetMacAddress = strMACAddress
    GoTo NormalExit

ErrorHandler:
    Call MsgBox(Err.Description, vbCritical, "Error")

NormalExit:
    End Function

Function Create_sn() As String

    Dim strMACAddress As String

    strMACAddress = fGetMacAddress()
    
    If strMACAddress <> "" Then
        'Call MsgBox(strMACAddress, vbInformation, "MAC Address")
        Create_sn = strMACAddress
    End If
    
End Function
Public Function reg_info()


End Function
Private Sub Command2_Click()
'curl -d "action=softin&DiskId=yb&D_rjbb=3&D_yhmc=ë��2a&D_lxdz=hunan&D_zh=1020428&D_zhlx=0&D_zhye=0&D_zcfsm=International Gold Rock Ltd&D_serverame=Goldrockfx-Server&D_gddh=05728330004&D_yddh=13341039392" -k http://211.99.249.141//WHSoft/DLL/SoftFind.asp

Dim io_res As String
Dim curl_cmd As String
Dim Server As String
Dim sn As String, yhmc As String, lxdz As String, zh As String, gddh As String, yddh As String, proxy As String


yhmc = Form3.Text1.Text
lxdz = Form3.Text2.Text
gddh = Form3.Text3.Text
yddh = Form3.Text4.Text
zh = Form3.Text5.Text
proxy = Form3.Text6.Text

sn = Create_sn

Server = " https://" + Server_ip(0) + "/DLL/SoftFind.php"

curl_cmd = App.path + "\curl.exe -s -d ""action=softin&DiskId=" + sn + "&D_rjbb=1&D_yhmc=" + yhmc + "&D_lxdz=" + lxdz + "&D_zh=" + zh + "&D_zhlx=0&D_zhye=0&D_zcfsm=International Gold Rock Ltd&D_serverame=Goldrockfx-Server&D_gddh=" + gddh + "&D_yddh=" + yddh + "&D_proxy=" + proxy + """ -k " + Server + ""

Command1.Enabled = False
Command2.Enabled = False

'MsgBox curl_cmd

io_res = RunCommand(curl_cmd)
'io_res = 1

io_res = clean_ascii(io_res)

If io_res = "1" Then

 Call SetRegistryValue(HKEY_CURRENT_USER, "software\GoldRockfx software\Info", "Sn", sn, eString, 0)
 'Call SetRegistryValue(HKEY_CURRENT_USER, "software\GoldRockfx software\Info", "Fn", Form5.Text1.Text & "\experts\", eString, 0)
  
 MsgBox "���ע��ɹ�, ���ǽ���ʱ������ϵ��Ϊ����ͨ��Ӧ�ķ���"

 Else

 MsgBox " ע��ʧ�ܣ�������Ϣ�����Ƿ���ȷ�����Ѿ�ע�������ϵ����"
 Command1.Enabled = True
 Command2.Enabled = True
 
End If


End Sub

Private Sub Command3_Click()
If MsgBox("ȷ���Ƿ��˳�?", vbYesNo) = vbYes Then End
End Sub

Private Sub Command4_Click()

    
End Sub

Private Sub Form_unLoad(cancel As Integer)
End

End Sub

Private Sub Text1_Change()

End Sub

