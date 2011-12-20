VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarshallxInet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function InternetGetConnectedState Lib "wininet.dll" (lpdwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetDial Lib "wininet.dll" (ByVal hwndParent As Integer, ByVal lpszConnectoid As String, ByVal dwFlags As Integer, ByRef lpdwConnection As Integer, ByVal dwReserved As Integer) As Integer
Private Declare Function InternetHangUp Lib "wininet.dll" (ByVal dwConnection As Long, ByVal dwReserved As Long) As Long

Dim thisInet As Inet
Dim thisSelfConnect As Boolean
Dim thisCancelDownload As Boolean
Dim thisUserAgent As String

Const icNone = 0               'No state to report.
Const icHostResolvingHost = 1  'The control is looking up the IP address of the specified host computer.
Const icHostResolved = 2       'The control successfully found the IP address of the specified host computer.
Const icConnecting = 3         'The control is connecting to the host computer.
Const icConnected = 4          'The control successfully connected to the host computer.
Const icRequesting = 5         'The control is sending a request to the host computer.
Const icRequestSent = 6        'The control successfully sent the request.
Const icReceivingResponse = 7  'The control is receiving a response from the host computer.
Const icResponseReceived = 8   'The control successfully received a response from the host computer.
Const icDisconnecting = 9      'The control is disconnecting from the host computer.
Const icDisconnected = 10      'The control successfully disconnected from the host computer.
Const icError = 11             'An error occurred in communicating with the host computer.
Const icResponseCompleted = 12 'The request has completed and all data has been received.
Const icString = 0             'Default. Retrieves data as string.
Const icByteArray = 1          'Retrieves data as a byte array.

Private Sub Class_Initialize()
    Set thisInet = frmMain.Inet1
    thisSelfConnect = False
End Sub

Public Property Get UserAgent() As String
    UserAgent = thisUserAgent
End Property

Public Property Let UserAgent(ByVal sUserAgent As String)
    thisUserAgent = sUserAgent
End Property

Public Sub CancelDownload()
    thisCancelDownload = True
End Sub

Public Function DownloadCancelled() As Boolean
    DownloadCancelled = thisCancelDownload
End Function

Public Function Connected() As Boolean
    Dim Flags As Long
    Connected = InternetGetConnectedState(Flags, 0&)
End Function

Public Function Connect(ByVal hwndParent As Integer)
    Call InternetDial(hwndParent, "DefaultDialUp", 1, 0, 0)
    thisSelfConnect = True
End Function

Public Sub Disconnect(Optional ByVal OnlyIfSelfConnected As Boolean = True)
    If thisSelfConnect Or Not OnlyIfSelfConnected Then
        Call InternetHangUp(0, 0)
        thisSelfConnect = False
    End If
End Sub

Private Function StartDownload(ByVal sURL As String, Optional ByVal sExecuteCommand As String = "GET") As Boolean
    Dim sBuffer As String
    Dim dDate As Date
    Dim iTime As Integer
    Dim bWaitForCancel As Boolean
    Dim iError As Long
    Dim sHeader As String
    On Error GoTo LocalErr
    StartDownload = False
    thisCancelDownload = False
    thisInet.URL = sURL
    dDate = Now()
    bWaitForCancel = False
    If Len(thisUserAgent) <> 0 Then sHeader = "User-Agent: " & thisUserAgent & vbCrLf
    sHeader = sHeader & "Range: bytes0-" & vbCrLf
    Call thisInet.Execute(, sExecuteCommand, , sHeader)
    Do While thisInet.StillExecuting
        DoEvents
        If Not bWaitForCancel Then
            If thisCancelDownload Then
                Call thisInet.Cancel
                bWaitForCancel = True
            End If
            iTime = DateDiff("s", dDate, Now())
            If iTime >= 5 Then
                Call thisInet.Cancel
                bWaitForCancel = True
            End If
        End If
    Loop
    If iTime > 5 Then
        Call frmMain.WriteLogEntry("Internet timeout! Waiting for a response from " & sURL & " for more than 5 seconds.")
    Else
        StartDownload = True
    End If
LocalErr:
    iError = Err.Number
    Select Case iError
    Case 0
        'do nothing
    Case 35761, 35756
        StartDownload = False
        Call frmMain.WriteLogEntry("Internet timeout! Waiting for a response from " & sURL & " for more than 5 seconds.")
        Call thisInet.Cancel
    Case Else
       On Error GoTo 0
       StartDownload = False
       Err.Raise iError
    End Select
    Do While thisInet.StillExecuting
        DoEvents
    Loop
End Function

Public Function GetRemoteFileSize(ByVal sURL As String) As Long
    Dim sValue As String
    sValue = GetHeader("Content-Length", sURL)
    If Len(sValue) <> 0 And Len(sValue) = Len(StripNonNumbers(sValue)) Then
        GetRemoteFileSize = Val(sValue)
    Else
        GetRemoteFileSize = 0
    End If
End Function

Public Function GetHeader(ByVal sHeader As String, Optional ByVal sURL As String = "") As String
    'sURL can be blank if we are getting a further header, having already gotten the first (and therefore already have the data)
    Dim sBuffer As String
    Dim iError As Long
    On Error GoTo LocalErr
    GetHeader = ""
    If Len(sURL) <> 0 Then
        If StartDownload(sURL, "HEAD") Then
            If Not thisCancelDownload Then
                sBuffer = thisInet.GetHeader(sHeader)
                DoEvents
                GetHeader = sBuffer
            End If
        End If
    End If
LocalErr:
    iError = Err.Number
    Select Case iError
    Case 0
        'do nothing
    Case 35761, 35756
        Call frmMain.WriteLogEntry("Internet timeout! Waiting for a response from " & thisInet.URL & " for more than 5 seconds.")
        Call thisInet.Cancel
    Case Else
       On Error GoTo 0
       Err.Raise iError
    End Select
    Do While thisInet.StillExecuting
        DoEvents
    Loop
End Function

Public Function CopyURLToFile(ByVal sURL As String, ByVal sDestFile As String, Optional ByRef frmCallback As Form, Optional ByVal bIgnoreSize As Boolean = False) As Boolean
    Dim hFile As Integer
    Dim sBuffer As String
    Dim iBytesExpected As Long
    Dim iBytesRead As Long
    Dim byteBuffer() As Byte
    Dim iError As Long
    On Error GoTo LocalErr
    thisCancelDownload = False
    CopyURLToFile = False
    hFile = 0
    iBytesRead = 0
    iBytesExpected = 0
    If StartDownload(sURL, "GET") Then
        If Not thisCancelDownload Then
            sBuffer = thisInet.GetHeader("Content-Length")
            DoEvents
            If Len(sBuffer) <> 0 Then iBytesExpected = Val(sBuffer)
            If iBytesExpected <> 0 Or bIgnoreSize Then
                If FileExists(sDestFile) Then
                    Call Kill(sDestFile)
                Else
                    If Not DirExists(GetFilePath(sDestFile)) Then Call MakePath(GetFilePath(sDestFile))
                End If
                hFile = FreeFile()
                Open sDestFile For Binary Access Write As #hFile
                iBytesRead = 0
                Do
                    byteBuffer() = thisInet.GetChunk(1024, icByteArray)
                    iBytesRead = iBytesRead + UBound(byteBuffer(), 1) + 1
                    Put #hFile, , byteBuffer()
                    If Not (frmCallback Is Nothing) Then Call frmCallback.DownloadProgress(iBytesRead, iBytesExpected)
                    If thisCancelDownload Then
                        Call thisInet.Cancel
                        Exit Do
                    End If
                Loop While UBound(byteBuffer(), 1) <> -1
                If iBytesExpected <> 0 Then
                    If iBytesRead = iBytesExpected Then CopyURLToFile = True
                ElseIf bIgnoreSize Then
                    If iBytesRead <> 0 Then CopyURLToFile = True
                End If
            End If
        End If
    End If
    'flow into error handler
LocalErr:
    iError = Err.Number
    Select Case iError
    Case 0
        'do nothing
    Case 35761, 35756
        Call frmMain.WriteLogEntry("Internet timeout! Waiting for a response from " & sURL & " for more than 5 seconds.")
        Call thisInet.Cancel
    Case Else
       On Error GoTo 0
       Err.Raise iError
    End Select
    If hFile <> 0 Then Close #hFile
    If Not CopyURLToFile Then
        If FileExists(sDestFile) Then Call Kill(sDestFile)
    End If
    Do While thisInet.StillExecuting
        DoEvents
    Loop
End Function

Public Function CopyURLToString(ByVal sURL As String, ByRef sRetVal As String, Optional ByRef frmCallback As Form, Optional ByVal bIgnoreSize As Boolean = False) As Boolean
    Dim sBuffer As String
    Dim iBytesExpected As Long
    Dim iBytesRead As Long
    Dim iError As Long
    On Error GoTo LocalErr
    thisCancelDownload = False
    CopyURLToString = False
    sRetVal = ""
    iBytesRead = 0
    iBytesExpected = 0
    If StartDownload(sURL, "GET") Then
        If Not thisCancelDownload Then
            sBuffer = thisInet.GetHeader("Content-Length")
            DoEvents
            If Len(sBuffer) <> 0 Then iBytesExpected = Val(sBuffer)
            If iBytesExpected <> 0 Or bIgnoreSize Then
                Do
                    sBuffer = thisInet.GetChunk(1024, icString)
                    iBytesRead = iBytesRead + Len(sBuffer)
                    'If Len(sBuffer) <> 0 Then dDate = Now() 'not sure why this is here. maybe we were planning to do our own check for 5 seconds of no data?
                    sRetVal = sRetVal & sBuffer
                    If thisCancelDownload Then
                        Call thisInet.Cancel
                        Exit Do
                    End If
                    If Not (frmCallback Is Nothing) Then Call frmCallback.DownloadProgress(iBytesRead, iBytesExpected)
                    If thisCancelDownload Then
                        Call thisInet.Cancel
                        Exit Do
                    End If
                Loop While Len(sBuffer) <> 0
                If iBytesExpected <> 0 Then
                    If iBytesRead = iBytesExpected Then CopyURLToString = True
                ElseIf bIgnoreSize Then
                    If iBytesRead <> 0 Then CopyURLToString = True
                End If
            End If
            Do While thisInet.StillExecuting 'do we really need this here? It's at the end as well and we don't do this here in CopyURLToFile...
                DoEvents
            Loop
        End If
    End If
    'flow into error handler
LocalErr:
    iError = Err.Number
    Select Case iError
    Case 0
        'do nothing
    Case 35761, 35756
        Call frmMain.WriteLogEntry("Internet timeout! Waiting for a response from " & sURL & " for more than 5 seconds.")
        Call thisInet.Cancel
    Case Else
       On Error GoTo 0
       Err.Raise iError
    End Select
    Do While thisInet.StillExecuting
        DoEvents
    Loop
End Function
