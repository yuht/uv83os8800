VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UV83 20220625�汾����̼��Զ��忪������"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6150
   StartUpPosition =   3  '����ȱʡ
   Begin MSCommLib.MSComm MSCommSer 
      Left            =   5340
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      InputMode       =   1
   End
   Begin VB.Frame Frame2 
      Caption         =   "�����ַ�"
      Height          =   3225
      Left            =   30
      TabIndex        =   5
      Top             =   990
      Visible         =   0   'False
      Width           =   6075
      Begin VB.CommandButton cmdWirteCommand 
         Caption         =   "д�뿪���ַ�"
         Height          =   735
         Left            =   960
         TabIndex        =   13
         Top             =   2370
         Width           =   4065
      End
      Begin VB.Frame Frame3 
         Caption         =   "�ַ������14���ַ���"
         Height          =   1245
         Left            =   90
         TabIndex        =   11
         Top             =   1020
         Width           =   5895
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   36
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   945
            Left            =   120
            MaxLength       =   14
            TabIndex        =   12
            Text            =   "BI1SLC"
            Top             =   240
            Width           =   5655
         End
      End
      Begin VB.ComboBox ComboCol 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   577
         Width           =   1665
      End
      Begin VB.ComboBox ComboRow 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   217
         Width           =   1665
      End
      Begin VB.Label Label4 
         Caption         =   "һ���ַ�ռ��8��λ��"
         Height          =   255
         Left            =   3150
         TabIndex        =   10
         Top             =   600
         Width           =   1875
      End
      Begin VB.Label Label3 
         Caption         =   "�ַ�λ�� �У�"
         Height          =   255
         Left            =   90
         TabIndex        =   9
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "�ַ�λ�� �У�"
         Height          =   255
         Left            =   90
         TabIndex        =   7
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ѡ�񴮿�"
      Height          =   915
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6075
      Begin VB.CommandButton Command3 
         Caption         =   "�ر�"
         Height          =   375
         Left            =   4080
         TabIndex        =   14
         Top             =   300
         Width           =   645
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   337
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "��"
         Height          =   375
         Left            =   3300
         TabIndex        =   2
         Top             =   300
         Width           =   645
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ˢ�´���"
         Height          =   375
         Left            =   4860
         TabIndex        =   1
         Top             =   300
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "ѡ�񴮿ڣ�"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdWirteCommand_Click()
    Dim Sumk As Integer
    Dim SerData(19) As Byte
    Dim i As Integer
    SerData(0) = &H57
    SerData(1) = &HA
    SerData(2) = &HF0
    SerData(3) = ComboRow.ListIndex * 16 + Len(Text1.Text) '��8λ����λ�ã���8λ���ַ�����
    SerData(4) = ComboCol.ListIndex '��λ�ã�0-255��
    For i = 5 To 18
        SerData(i) = &HFF
    Next
    
    For i = 1 To Len(Text1.Text)
        SerData(4 + i) = Right("0" & Asc(Mid(Text1.Text, i, 1)), 2)
    Next
    
    For i = 0 To 18
        'On Error Resume Next
        Sumk = Sumk + SerData(i)
        'Debug.Print "serdata(" & Right("0" & i, 2) & ")=" & Right("0" & Hex(SerData(i)), 2)
    Next
    'Debug.Print "hex Sumk=" & Hex(Sumk)
    SerData(19) = Sumk Mod 256
    'Debug.Print "serdata(" & Right("0" & i, 2) & ")=" & Right("0" & Hex(SerData(19)), 2)
    
'    MSCommSer.Settings = "9600,n,8,1"
'    MSCommSer.OutBufferCount = 0
    
    'MSCommSer.OutBufferSize = 20
    MSCommSer.Output = SerData
    
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    If GetComboComNO < 0 Then
        Debug.Print "exit open port"
        Exit Sub
    End If
    
    MSCommSer.CommPort = GetComboComNO
    

    MSCommSer.PortOpen = True
    If Err Then
        Exit Sub
    End If
    Frame2.Visible = True

    
    
    
End Sub

Private Sub Command2_Click()
    Combo1.Clear
    GetExistPort
    Frame2.Visible = False
End Sub

Private Sub Command3_Click()
    If MSCommSer.PortOpen = True Then
        MSCommSer.PortOpen = False
    End If
    Frame2.Visible = False
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    ComboRow.Clear
    ComboCol.Clear
    Frame2.Visible = False
    For i = 0 To 7
        ComboRow.AddItem i
        ComboCol.AddItem i
    Next
    For i = 8 To 255
        ComboCol.AddItem i
    Next
    ComboRow.ListIndex = 3
    ComboCol.ListIndex = 40
    
    GetExistPort
End Sub

Function GetComboComNO() As Integer
    
    Dim i As String 'Integer
    'Debug.Print "Combo1.ListCount:" & Combo1.ListCount
    If Combo1.ListCount = 0 Then
        GetComboComNO = -1
        Exit Function
    End If
    i = Combo1.List(Combo1.ListIndex)
    'Debug.Print i
    i = Right(i, Len(i) - 3)
    'Debug.Print i
    GetComboComNO = i
    
End Function


Function GetExistPort()
    On Error Resume Next
    Dim i As Integer
    For i = 1 To 256
        Err.Clear
        If MSCommSer.PortOpen = True Then MSCommSer.PortOpen = False
        If Err Then Debug.Print Err.Description
        MSCommSer.CommPort = i
            MSCommSer.PortOpen = True
            If MSCommSer.PortOpen = False Then
                'Debug.Print "err: com" & i
            Else
                Combo1.AddItem "COM" & i
                MSCommSer.PortOpen = False
            End If
    Next
    If Combo1.ListCount Then
        Combo1.ListIndex = 0
    End If
    
End Function

Private Sub MSCommSer_OnComm()

    Dim karr() As Byte
    Dim i As Integer
    
    Select Case MSCommSer.CommEvent
        Case comEvReceive   '���ܵ�Rthreshold���ַ������¼�������������ֱ����Input���Դӽ��ܻ������ж�ȡ��ɾ���ַ���

            karr = MSCommSer.Input
            For i = 0 To UBound(karr)
                Debug.Print "karr(" & i & ")=" & karr(i)
                If karr(i) = 6 Then
                    MsgBox "д��ɹ�!", vbInformation + vbMsgBoxSetForeground, "д�뿪���ַ�"
                    Exit For
                End If
            Next
        Case comEvSend   '���ͻ���������������Sthreshold����˵�������Ѿ�������һЩ���ݣ����������Output���Լ����������ݡ�
            Debug.Print "output buffer size :"; MSCommSer.OutBufferSize
           
            
        Case comEvCTS   'Clear To Send'�ź���״̬�����仯��
        Case comEvDSR   'Data Set Ready'�ź���״̬��1�䵽0��
        Case comEvCD   'Carrier Detect'�ź���״̬�����仯��
        Case comEvRing   '��⵽�����ź�?
        Case comEvEOF   '���ܵ��ļ�������
        Case Else
    End Select

End Sub

Private Sub Text1_Change()
    Dim Txt As String
    Dim Txt2 As String
    Dim i As Integer
    Txt = Text1.Text
    Txt = Left(Txt, IIf(Len(Txt) > 14, 14, Len(Txt)))
    For i = 1 To Len(Txt)
        Debug.Print Asc(Mid(Txt, i, 1))
        If Asc(Mid(Txt, i, 1)) > 0 Then
            Txt2 = Txt2 & Mid(Txt, i, 1)
        End If
    Next
    Text1.Text = Txt2
    
End Sub