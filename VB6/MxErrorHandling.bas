Attribute VB_Name = "MxErrorHandling"
Option Explicit

Dim CallStack() As String
Dim CallStackSize As Integer

Public Sub GlobalErr(Optional ByVal sMsg As String = "")
    Dim sEntry As String
    If Len(sMsg) = 0 Then sMsg = Err.Description
    sEntry = sMsg & vbCrLf & _
    "Error: " & Err.Number & vbCrLf & _
    "Description: " & Err.Description & vbCrLf & _
    "Return stack:"
    Do While CallStackSize > 0
        sEntry = sEntry & vbCrLf & "    " & CallStackPop
    Loop
    Call frmMain.WriteLogEntry(sEntry, LogIE)
End Sub

Public Sub Panic(Optional ByVal sMsg As String = "")
    Call CallStackPush("MxErrorHandling.Panic(" & CStr(sMsg) & ")")
    If CL_noexcept Then
        Call Err.Raise(51) '51="Internal error"
    Else
        Call GlobalErr(sMsg)
    End If
    Call CallStackPop
End Sub

Public Sub CallStackPush(sProc As String)
    CallStackSize = CallStackSize + 1
    ReDim Preserve CallStack(CallStackSize)
    CallStack(CallStackSize) = sProc
End Sub

Public Function CallStackPop() As String
    CallStackPop = CallStack(CallStackSize)
    CallStackSize = CallStackSize - 1
    ReDim Preserve CallStack(CallStackSize)
End Function
