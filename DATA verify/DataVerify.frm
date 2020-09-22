VERSION 5.00
Begin VB.Form DataVerify 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Data Verify"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10590
   Icon            =   "DataVerify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   2280
      TabIndex        =   6
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton Command1 
         Caption         =   "Change"
         Height          =   375
         Left            =   4080
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox CmbBaudRate 
         Height          =   315
         ItemData        =   "DataVerify.frx":08CA
         Left            =   2760
         List            =   "DataVerify.frx":08DA
         TabIndex        =   8
         Text            =   "9600"
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox CmbCommPort 
         Height          =   315
         Left            =   960
         TabIndex        =   7
         Text            =   "1"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comm Port:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "BaudRate:"
         Height          =   195
         Left            =   1920
         TabIndex        =   10
         Top             =   360
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   10335
      Begin VB.TextBox TxtSend 
         Height          =   855
         Left            =   2520
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   3000
         Width           =   4815
      End
      Begin VB.TextBox TxtDisp 
         Height          =   2415
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   480
         Width           =   5000
      End
      Begin VB.TextBox TxtSendLog 
         Height          =   2415
         Left            =   5160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   480
         Width           =   5000
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"DataVerify.frx":08FA
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   4200
         Width           =   10140
      End
      Begin VB.Label LblReceived 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Received"
         Height          =   195
         Left            =   1080
         TabIndex        =   5
         Top             =   240
         Width           =   690
      End
      Begin VB.Label LblSend 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sent"
         Height          =   195
         Left            =   6720
         TabIndex        =   4
         Top             =   240
         Width           =   330
      End
   End
End
Attribute VB_Name = "DataVerify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents comm As MSComm
Attribute comm.VB_VarHelpID = -1
Private fhndl As Integer
Private bytesReceived As Long
Dim temp As String
Dim DummyBuff As String



Private Const DatabaseExt = "\ZBAComm.ini"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function ReadINI(Section As String, KeyName As String, Filename As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName$, "", sRet, Len(sRet), Filename))
End Function
Public Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function














Private Sub CmdSend_Click()

End Sub

Private Sub Command1_Click()
retval = WriteCommSet(CmbCommPort.Text, CmbBaudRate.Text, "1", "n", "8", "0", "2048", "2048")
START

End Sub








'---OnComm is an event-------
Private Sub comm_OnComm()
   ' Dim temp As String
    temp = comm.Input
    'Put #fhndl, , temp
    TxtDisp.Text = TxtDisp & temp
    bytesReceived = bytesReceived + Len(temp)
    LblReceived.Caption = "Received: " & bytesReceived & " bytes"
End Sub

Public Function DispBytes() As String
DispBytes = bytesReceived
End Function

Public Function DispData() As String
DispData = temp
End Function





Public Function START() As String
Set comm = New MSComm
            
            Dim temp As String
            Dim i As Integer
            For i = 4 To 1 Step -1
                temp = temp & "," & ReadINI("ZBAComm", "par" & i, App.Path & DatabaseExt)
            Next i

If Len(temp) < 6 Then
MsgBox "No Comm Settings Found", 16, "Error"
START = "NoComm"
Exit Function
End If



     

    
            '-------READING COMM PARAMETERS
            'comm.Settings = "115200,n,8,1"
            comm.Settings = Right(temp, Len(temp) - 1)
            comm.Handshaking = ReadINI("ZBAComm", "par0", App.Path & DatabaseExt)
            comm.CommPort = CLng(ReadINI("ZBAComm", "par5", App.Path & DatabaseExt))
            comm.RThreshold = 1
            comm.InBufferSize = ReadINI("ZBAComm", "InBuff", App.Path & DatabaseExt)
            comm.OutBufferSize = ReadINI("ZBAComm", "OutBuff", App.Path & DatabaseExt)
            comm.PortOpen = True
            '-----------------------------------








        comm.InBufferCount = 0
        
        
'        fhndl = FreeFile
'        Open App.Path + "\Image.dat" For Output As fhndl
'        Close fhndl
'        Open App.Path + "\Image.dat" For Binary As fhndl
        bytesReceived = 0






End Function












Public Function WriteCommSet(PortNo As String, BaudRate As String, StopBits As String, Parity As String, Databits As String, FlowControl As String, InBuff As String, OutBuff As String)
Dim a As String

WriteINI "ZBAComm", "par5", PortNo, App.Path & DatabaseExt
WriteINI "ZBAComm", "par4", BaudRate, App.Path & DatabaseExt
WriteINI "ZBAComm", "par1", StopBits, App.Path & DatabaseExt
WriteINI "ZBAComm", "par3", Parity, App.Path & DatabaseExt
WriteINI "ZBAComm", "par2", Databits, App.Path & DatabaseExt
WriteINI "ZBAComm", "par0", FlowControl, App.Path & DatabaseExt
WriteINI "ZBAComm", "InBuff", InBuff, App.Path & DatabaseExt
WriteINI "ZBAComm", "OutBuff", OutBuff, App.Path & DatabaseExt


End Function




Public Function CheckAvailablePort(CommPort As ComboBox)
Dim a As String
Dim i As Integer
For i = 1 To 30
If isavailable(i) = True Then CommPort.AddItem (i)
Next i
End Function


Private Function isavailable(port As Integer) As Boolean
Dim temp As String
temp = port
On Error Resume Next
comm.CommPort = temp
comm.PortOpen = True
If comm.PortOpen = True Then isavailable = True
comm.PortOpen = False
End Function




Private Sub Form_Load()
CmbCommPort.Text = ReadINI("ZBAComm", "par5", App.Path & DatabaseExt)
CmbBaudRate.Text = ReadINI("ZBAComm", "par4", App.Path & DatabaseExt)

If START = "NoComm" Then
Call CheckAvailablePort(CmbCommPort)
End If

End Sub

Private Sub TxtSend_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
TxtSend.Text = temp
End If



If KeyCode = 27 Then
TxtSendLog.Text = ""
TxtSend.Text = ""
TxtDisp.Text = ""
End If


If KeyCode = 38 Then
TxtSend.Text = DummyBuff
End If

If KeyCode = 46 Then
TxtSend.Text = ""
End If

If KeyCode = 13 Then
comm.Output = TxtSend.Text
TxtSendLog.Text = TxtSendLog.Text & TxtSend.Text
LblSend.Caption = "Sent: " & Len(TxtSendLog.Text) & " bytes"
DummyBuff = TxtSend.Text
TxtSend.Text = ""
End If


End Sub
