VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "RFIDClient"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   6480
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command1 
      Caption         =   "���ի��s"
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox commText 
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Text            =   "1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox StatusText1 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox CardNOText 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.ListBox RxList 
      Height          =   2370
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Width           =   6015
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5520
      Top             =   2280
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4200
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputMode       =   1
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "�]�wCommPort"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "�n�J���A"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "�ϥά���"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "�n�J�d��"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      Caption         =   "�п�J�P���d"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'For SQLite Use
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long


Dim InByte() As Byte
Dim NowSecond As Integer
Dim NowMinute As Integer
Dim NowHour As Integer
Dim NowDay As Integer
Dim NowMonth As Integer
Dim NowYear As Integer
Dim NowTime As Integer

Private Sub cmdExit_Click()
    MSComm1.PortOpen = False
    Timer1.Enabled = False
    End
End Sub

Private Sub cmdTextClear_Click()
    RxList.Clear
    CardNOText.Text = ""
    NameText.Text = ""
    StatusText1.Text = ""
End Sub

Private Sub cmdCheck_Click()
    '�Ұ�RFID Reader
    MSComm1.CommPort = commText.Text
    MSComm1.PortOpen = True
    Timer1.Enabled = True
End Sub

Private Sub Command1_Click()
    Dim uid As String
    Dim cid As String
    Dim status As String
    Dim n As Integer
    Dim dollar As Integer
    Dim newdollar As Integer
    Dim conn As New LiteConnection
    Dim record As New LiteStatement
    

    '�s�u��Ʈw�θ�ƪ�
    Call conn.Open(App.Path & "\CarManager")
    record.ActiveConnection = conn
    
    '�ˬd�O�_�����d���s�b
    cid = CardNOText.Text
    sqlstring = "select * from userlist where cid='" & cid & "'"
    Call record.Prepare(sqlstring)
    n = record.RowCount
    If n > 0 Then   '�p�G�����
        Call record.Step(1)
        dollar = record.ColumnValue("dollar")
        uid = record.ColumnValue("uid")
        
        
        '�ˬd���ϥΪ̬O�n�����٬O�i��
        Call record.Close
        sqlstring = "select * from checkinout where cid='" & cid & "' order by date desc, time desc"
        Call record.Prepare(sqlstring)
        n = record.RowCount
        If n > 0 Then   '�p�G����ơA�h��X��
            Call record.Step(1)
            status = record.ColumnValue("status")
            If status = "�i�J" Then
                status = "���}"
                newdollar = dollar
            Else
                status = "�i�J"
                newdollar = dollar - 40
            End If
        Else    '�S����ƪ��ܡA�ܤ֨����i��b�̭��A�@�w�O�n�i��
            status = "�i�J"
            newdollar = dollar - 40
        End If
        
        
        If newdollar < 0 And status = "�i�J" Then   '�p�G�����O�i�J�����p�U�A���l�B�S�����A���Ӧ^��
            Debug.Print "�l�B�Ⱦl" & dollar & "�A�����\�q�L"
            StatusText1.Text = "�l�B�Ⱦl" & dollar & "�A�����\�q�L"
        Else
            '�}�l��s��Ʈw�����
            Call record.Close
                        
            '�W�[�i�X�O��
            sqlstring = "insert into CheckInOut (date,time,status,cid,uid) values ("
            sqlstring = sqlstring & ("'" & Format(Now, "YYYY/MM/DD") & "',")
            sqlstring = sqlstring & ("'" & Format(Now, "hh:mm:ss") & "',")
            sqlstring = sqlstring & ("'" & status & "',")
            sqlstring = sqlstring & ("'" & cid & "',")
            sqlstring = sqlstring & ("'" & uid & "'")
            sqlstring = sqlstring & ")"
            Call conn.Execute(sqlstring)
            
            If status = "�i�J" Then '�p�G�����O�i�J�����p�U�A�~�ݭn�����ڡA��W�C����O��
                '��s�l�B
                sqlstring = "update userlist set dollar='" & newdollar & "' where cid='" & cid & "'"
                Call conn.Execute(sqlstring)
                
                '�W�[����O��
                sqlstring = "insert into Dollar (date,time,uid,status,dollar) values ("
                sqlstring = sqlstring & ("'" & Format(Now, "YYYY/MM/DD") & "',")
                sqlstring = sqlstring & ("'" & Format(Now, "hh:mm:ss") & "',")
                sqlstring = sqlstring & ("'" & uid & "',")
                sqlstring = sqlstring & ("'����',")
                sqlstring = sqlstring & ("'40'")
                sqlstring = sqlstring & ")"
                Call conn.Execute(sqlstring)
            End If
            
            '�@�����`�����p
            Debug.Print "�w" & status & "�A�l�B�|��:" & newdollar & "��"
            StatusText1.Text = "�w" & status & "�A�l�B�|��:" & newdollar & "��"
        End If
    Else    '�S��ƪ����p
        Debug.Print "���d�����s�b�A�����\�q�L"
        StatusText1.Text = "���d�����s�b�A�����\�q�L"
    End If
    
    
    '������Ʈw
    Call record.Close
    Call conn.Close
    
    
    '��s�e���W�����A
    RxList.AddItem Format(Now, "hh:mm:ss") & StatusText1.Text, 0
End Sub

Private Sub Form_Load()
    Timer1.Enabled = False
    RxList.Clear
    Caption = Now
    CardNOText.Text = ""
    StatusText1.Text = ""
    
        

End Sub

Private Sub Label5_Click()
End Sub

Private Sub Timer1_Timer()
    Dim i%, Buf$
    Dim txtBuf$, comBuf$, NameBuf$

    
    RxFlag = 0
    InByte = MSComm1.Input
    For i = LBound(InByte) To UBound(InByte) - 2
        Buf = Buf + Chr(InByte(i))
        RxFlag = 1
    Next i
    comBuf = Buf
    

    If RxFlag > 0 Then
        CardNOText.Text = comBuf
        
        '�Y�q�L���ܭn�ܸ�Ʈw�����B�z
        Call Command1_Click
    End If

    
    Caption = Now
End Sub

