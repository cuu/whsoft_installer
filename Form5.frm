VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������ ѡ��װ·��"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command3 
      Caption         =   "�ı�·��"
      Height          =   375
      Left            =   7800
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��װ"
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��һ��"
      Height          =   495
      Left            =   4680
      TabIndex        =   2
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1800
      Width           =   7695
   End
   Begin VB.Label Label3 
      Caption         =   "��ע����ȷ����Ŀ¼����experts��files��Ŀ¼���Ա㰲װ˳����ɣ� Ȼ���� ��װ"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   600
      Width           =   8175
   End
   Begin VB.Label Label2 
      Caption         =   "��ȷѡ��װ·�������в�ͬ�汾��MT��㽻��ƽ̨����ѡ���ƽ̨��Ŀ¼�µ� Terminal.exe ����"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   240
      Width           =   8535
   End
   Begin VB.Label Label1 
      Caption         =   "��װ·��"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function copy_all(ByVal path As String) As Integer
    
    Dim pa As Integer
    Dim dfile As String, I As Integer
    Dim str As String
    
    pa = PathFileExists(path)
    If pa = 1 Then
        For I = 0 To 2
            str = path & "\files\" & Files(I)
            If Dir(str) = "" Then
                FileCopy Files(I), path + "\files\" + Files(I)
            Else
                Kill path + "\files\" + Files(I)
                FileCopy Files(I), path + "\files\" + Files(I)
            End If
         Next I
         
         str = path & "\libraries\" & Files(3)
         If Dir(str) = "" Then
            FileCopy Files(3), path + "\libraries\" + Files(3)
         Else
            Kill path + "\libraries\" + Files(3)
            FileCopy Files(3), path + "\libraries\" + Files(3)
         End If
         
         ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         'indicators or \
         'str = path & "\" & Files(4)
         'If Dir(str) = "" Then
         '   FileCopy Files(4), path + "\" + Files(4)
        ' End If
         If Files(4) = Vers(0) Then
            FileCopy Files(4), path + "\indicators\" + Files(4)
        Else
            FileCopy Files(4), path + "\" + Files(4)
        End If
        
        ' goldkey2.dll & gk4.exe
         str = path & "\files\" & Files(5)
         If Dir(str) = "" Then
            FileCopy Files(5), path + "\files\" + Files(5)
         Else
            Kill path + "\files\" + Files(5)
            FileCopy Files(5), path + "\files\" + Files(5)
         End If
         
          str = path & "\files\" & Files(6)
         If Dir(str) = "" Then
            FileCopy Files(6), path + "\files\" + Files(6)
         Else
            Kill path + "\files\" + Files(6)
            FileCopy Files(6), path + "\files\" + Files(6)
         End If
                 
         copy_all = 1
       ' MsgBox "ok"
        
    Else
        MsgBox "��װĿ���ļ��в�����,��ȷ��������Ϣ�Ƿ���ȷ"
    End If
    
End Function
Public Function get_open_dir() As String
    Dim fDir As String
    
    Dim fname As String
    Dim ftmp As String
    
    Dim I As Integer
    Dim n As Integer
    
    fname = OpenFile(Me.hwnd, "��ѡ��Terminal.exe", "Terminal.exe", "Terminal exe|*.exe")
    
    If fname <> "" Then
        n = InStrRev(fname, "\")
        For I = 1 To n
            ftmp = Mid(fname, I, 1)
            fDir = fDir & ftmp
        Next I
    Else
        fDir = fname
    End If
    
    get_open_dir = fDir
    
End Function

Private Sub Command1_Click()
    Form2.Show
    Form5.Hide
    
End Sub

Private Sub Command2_Click()
    Dim n As Integer
        
    Command2.Enabled = False
    
    
    n = copy_all(Text1.Text + "\experts\")
    
  Call DeleteRegistryValueOrKey(HKEY_CURRENT_USER, "software\GoldRockfx Software\Info", "Fn")
  Call SetRegistryValue(HKEY_CURRENT_USER, "software\GoldRockfx software\Info", "Fn", Form5.Text1.Text + "\experts\", vbString, 0)
    'n = 1
    If n = 1 Then
        If MsgBox("��װ�ɹ�,�Ƿ����ڽ���ע��? (�Ѿ�ע����Ļ�Ա����Ҫ�ظ�ע��)?", vbYesNo) = vbYes Then
            Form3.Show
            Form5.Hide
        Else
          
        End
        End If
    Else
        MsgBox ("��װʧ��,��������Ƿ�������ȷ,�����Ƿ����㹻�ռ䰲װ,Ȼ���������б�����")
        End
        
    End If
    Command2.Enabled = True
    
End Sub

Private Sub Command3_Click()
  Dim fDir As String
  
    fDir = get_open_dir
    If Len(fDir) > 1 Then
    
    Text1.Text = fDir
    End If
    
End Sub

Private Sub Form_unLoad(cancel As Integer)
End

End Sub


