VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZBAComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents comm As MSComm
Attribute comm.VB_VarHelpID = -1
Private fhndl As Integer
Private bytesReceived As Long
Dim temp As String


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





'---OnComm is an event-------
Private Sub comm_OnComm()
   ' Dim temp As String
    temp = comm.Input
    Put #fhndl, , temp
    bytesReceived = bytesReceived + Len(temp)
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
        'comm.Output = Chr(&H86) + Chr(&H0) + Chr(&H86)
        fhndl = FreeFile
        Open App.Path + "\Image.dat" For Output As fhndl
        Close fhndl
        Open App.Path + "\Image.dat" For Binary As fhndl
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




Public Function CheckAvailablePort() As String
Dim a As String
Dim i As Integer
For i = 1 To 30
If isavailable(i) = True Then CheckAvailablePort = CheckAvailablePort & ", " & (i)
Next i
CheckAvailablePort = Right(CheckAvailablePort, Len(CheckAvailablePort) - 1)
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









