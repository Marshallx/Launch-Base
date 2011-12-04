VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarshallxCSFClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type CSFLABEL
    ValueID() As String * 4
    Substrings As Long
    LabelNameLength As Long
    LabelName As String
    ValueLength() As Long
    Value() As String
    ExtraValueLength() As Long
    ExtraValue() As String
End Type
Private Type CSFFILEHEADER
     HeaderID As String * 4 'CSF header identifier, must be " FSC".
     Version As Long 'CSF Version (YR uses 3, unknown what the difference is)
     Labels As Long 'Total number of labels.
     ExtraValues As Long 'Total number of extra values.
     Zippo As Long 'Unused by the game.
     Language As Long 'Language of the string table. See below for a list.
     '0 = US (English)*
     '1 = UK (English)
     '2 = German*
     '3 = French*
     '4 = Spanish
     '5 = Italian
     '6 = Japanese
     '7 = Jabberwockie
     '8 = Korean*
     '9 = Chinese*
     '>9 = Unknown
     '*RA2/YR has been released in this language.
End Type
Private thisCSFPath As String 'path last loaded from / default to load from next time
Private thisCSFHeader As CSFFILEHEADER
Private thisLabels() As CSFLABEL
Private thisLabelPos As Integer
'Public Type CSFLABELHEADER
'     LabelID As String * 4 'Label header identifier, must be " LBL" otherwise read the next 4 bytes.
'     Substrings As Long 'Total number of label values for this label, usually just 1.
'     LabelNameLength As Long 'Size of the label name that follows.
'     LabelName As String * LabelNameLength 'A non-zero-terminated string that is as long as the DWORD at 0x8 says. If it is longer, the rest will be cut off.
'End Type
'Public Type CSFLABELVALUE
'     ValueID As String * 4 '" RTS" (no extra value) or "WRTS" (there is an extra value)
'     ValueLength As Long 'This holds the length of the Unicode string (the value) that follows.
'     Value As String * ValueLength 'This holds the encoded value of the label. Note that this is ValueLength*2 bytes long, because the value is a Unicode string, i.e. every character is a word instead of a byte. To decode the value to a Unicode string, not every byte of the value data (or substract it from 0xFF).
'     ExtraValueLength As Long 'This holds the length of the extra value string that follow. This only applies if the identifier is "WRTS" and not " RTS".
'     ExtraValue As String * ExtraValueLength 'Like the label name, a non-zero-terminated string that is as long as ExtraValueLength says. If it is longer, the rest will be cut off.
'End Type

Public Property Get LabelCount() As Long
    LabelCount = thisCSFHeader.Labels
End Property

Public Property Get LabelName(ByVal iLabel As Long) As String
    LabelName = thisLabels(iLabel).LabelName
End Property

Public Property Get Value(ByVal iLabel As Long) As String
    'only gives access to first value
    Value = thisLabels(iLabel).Value(1)
End Property

Public Property Get ExtraValue(ByVal iLabel As Long) As String
    'only gives access to first value
    If thisLabels(iLabel).ValueID(1) = "WRTS" Then
        ExtraValue = thisLabels(iLabel).ExtraValue(1)
    Else
        ExtraValue = ""
    End If
End Property

Public Sub Initialise(Optional ByVal iLanguage As Long = 0)
    thisCSFPath = ""
    thisCSFHeader.HeaderID = " FSC"
    thisCSFHeader.Version = 3
    thisCSFHeader.Labels = 0
    thisCSFHeader.ExtraValues = 0
    thisCSFHeader.Zippo = 0
    thisCSFHeader.Language = iLanguage
    ReDim thisLabels(500)
    thisLabelPos = 0
End Sub

Public Sub SaveCSF(Optional ByVal sSavePath As String = "")
    Dim hFile As Integer
    Dim iLabel As Long
    Dim iValue As Long
    If sSavePath <> "" Then thisCSFPath = sSavePath
    If FileExists(thisCSFPath) Then Call Kill(thisCSFPath)
    hFile = FreeFile()
    Open thisCSFPath For Binary As #hFile
    Put #hFile, , thisCSFHeader
    For iLabel = 1 To thisCSFHeader.Labels
        Put #hFile, , " LBL"
        Put #hFile, , thisLabels(iLabel).Substrings
        Put #hFile, , thisLabels(iLabel).LabelNameLength
        Put #hFile, , thisLabels(iLabel).LabelName
        iValue = 1
        Do While iValue <= thisLabels(iLabel).Substrings
            Put #hFile, , thisLabels(iLabel).ValueID(iValue)
            Put #hFile, , thisLabels(iLabel).ValueLength(iValue)
            If thisLabels(iLabel).ValueLength(iValue) > 0 Then Put #hFile, , thisLabels(iLabel).Value(iValue)
            If thisLabels(iLabel).ValueID(iValue) = "WRTS" Then
                Put #hFile, , thisLabels(iLabel).ExtraValueLength(iValue)
                If thisLabels(iLabel).ExtraValueLength(iValue) > 0 Then Put #hFile, , thisLabels(iLabel).ExtraValue(iValue)
            End If
            iValue = iValue + 1
        Loop
    Next iLabel
    Close #hFile
End Sub

Private Function DecodeValue(ByVal sValue As String, Optional ByVal bToUnicode As Boolean = False) As String
'This works great, but what if it was supposed to be in unicode (e.g. Chinese?)
    Dim iLoop As Long
    Dim sText$, sConverted() As Byte
    If bToUnicode Then
        'string we want to encode is ANSI, so need to convert to unicode first
        sValue = VBA.StrConv(sValue, vbUnicode)
    End If
    sText = sValue
    sConverted = VBA.StrConv(sText, vbFromUnicode)
    iLoop = Len(sValue) - 1
    Do While iLoop <> -1
        sConverted(iLoop) = Not sConverted(iLoop)
        iLoop = iLoop - 1
    Loop
    DecodeValue = VBA.StrConv(sConverted, vbUnicode)
End Function

Public Function GetLang(Optional ByVal sLoadPath As String = "") As Long
    Dim hFile As Integer
    If Len(sLoadPath) <> 0 Then
        If thisCSFPath <> sLoadPath Then
            Call Initialise
            thisCSFPath = sLoadPath
            hFile = FreeFile()
            Open thisCSFPath For Binary As #hFile
            Get #hFile, , thisCSFHeader
            Close #hFile
            thisCSFPath = ""
        End If
    End If
    If thisCSFHeader.HeaderID = " FSC" Then
        If thisCSFHeader.Version = 3 Then
            GetLang = thisCSFHeader.Language
        Else
            GetLang = -1
        End If
    Else
        GetLang = -1
    End If
End Function

Public Function LoadCSF(Optional ByVal sLoadPath As String = "") As Boolean
    Dim hFile As Integer
    Dim iLabel As Long
    Dim iValue As Long
    Dim sBuffer As String
    LoadCSF = False
    If Len(sLoadPath) <> 0 Then thisCSFPath = sLoadPath
    hFile = FreeFile()
    Open thisCSFPath For Binary As #hFile
    Get #hFile, , thisCSFHeader
    If thisCSFHeader.HeaderID = " FSC" Then
        If thisCSFHeader.Version = 3 Then
            ReDim thisLabels(thisCSFHeader.Labels + 500)
            iLabel = 1
            Do While iLabel <= thisCSFHeader.Labels
                sBuffer = String$(4, " ")
                Do While Not EOF(hFile)
                    Get #hFile, , sBuffer
                    If sBuffer = " LBL" Then Exit Do
                Loop
                Get #hFile, , thisLabels(iLabel).Substrings
                Get #hFile, , thisLabels(iLabel).LabelNameLength
                sBuffer = String$(thisLabels(iLabel).LabelNameLength, " ")
                iValue = thisLabels(iLabel).Substrings
                ReDim Preserve thisLabels(iLabel).ValueID(iValue)
                ReDim Preserve thisLabels(iLabel).ValueLength(iValue)
                ReDim Preserve thisLabels(iLabel).Value(iValue)
                ReDim Preserve thisLabels(iLabel).ExtraValueLength(iValue)
                ReDim Preserve thisLabels(iLabel).ExtraValue(iValue)
                Get #hFile, , sBuffer
                thisLabels(iLabel).LabelName = sBuffer
                iValue = 1
                Do While iValue <= thisLabels(iLabel).Substrings
                    Get #hFile, , thisLabels(iLabel).ValueID(iValue)
                    Select Case thisLabels(iLabel).ValueID(iValue)
                    Case " RTS", "WRTS"
                        'valid
                    Case Else
                        Exit Function
                    End Select
                    Get #hFile, , thisLabels(iLabel).ValueLength(iValue)
                    If thisLabels(iLabel).ValueLength(iValue) > 0 Then
                        sBuffer = String$(thisLabels(iLabel).ValueLength(iValue) * 2, " ")
                        Get #hFile, , sBuffer
                        thisLabels(iLabel).Value(iValue) = sBuffer
                    Else
                        thisLabels(iLabel).Value(iValue) = ""
                    End If
                    If thisLabels(iLabel).ValueID(iValue) = "WRTS" Then
                        Get #hFile, , thisLabels(iLabel).ExtraValueLength(iValue)
                        If thisLabels(iLabel).ExtraValueLength(iValue) > 0 Then
                            sBuffer = String$(thisLabels(iLabel).ExtraValueLength(iValue), " ")
                            Get #hFile, , sBuffer
                            thisLabels(iLabel).ExtraValue(iValue) = sBuffer
                        Else
                            thisLabels(iLabel).ExtraValue(iValue) = ""
                        End If
                    Else
                        thisLabels(iLabel).ExtraValueLength(iValue) = 0
                        thisLabels(iLabel).ExtraValue(iValue) = ""
                    End If
                    iValue = iValue + 1
                Loop
                iLabel = iLabel + 1
            Loop
            LoadCSF = True
        'Else
            'unsupported CSF version
        End If
    'Else
        'not a CSF file
    End If
    Close #hFile
End Function

Public Sub AddLabel(ByVal sLabelName As String, ByVal sValue As String, Optional ByVal sExtraValue As String = "", Optional ByVal bANSI As Boolean = True, Optional ByVal bReplaceExisting As Boolean = True)
    Dim iLabelNameLength
    Dim iLabel As Long
    Dim iValue As Long
    sLabelName = LCase$(sLabelName)
    iLabelNameLength = Len(sLabelName)
    iLabel = 1
    Do While iLabel <= thisCSFHeader.Labels
        If thisLabels(iLabel).LabelNameLength = iLabelNameLength Then
            If LCase$(thisLabels(iLabel).LabelName) = sLabelName Then Exit Do
        End If
        iLabel = iLabel + 1
    Loop
    If LCase$(thisLabels(iLabel).LabelName) <> sLabelName Or bReplaceExisting Then
        If iLabel > thisCSFHeader.Labels Then
            thisCSFHeader.Labels = iLabel
            thisCSFHeader.ExtraValues = iLabel 'ra2md.csf has this set the same as label count, even though not every label has an extra value
            If iLabel > UBound(thisLabels()) Then ReDim Preserve thisLabels(iLabel + 500)
            thisLabels(iLabel).LabelName = sLabelName
            thisLabels(iLabel).LabelNameLength = Len(sLabelName)
            thisLabels(iLabel).Substrings = 1
        End If
        ReDim thisLabels(iLabel).ValueID(1)
        ReDim thisLabels(iLabel).ValueLength(1)
        ReDim thisLabels(iLabel).Value(1)
        ReDim thisLabels(iLabel).ExtraValueLength(1)
        ReDim thisLabels(iLabel).ExtraValue(1)
        If Len(sExtraValue) = 0 Then
            thisLabels(iLabel).ValueID(1) = " RTS"
            thisLabels(iLabel).ExtraValueLength(1) = 0
            thisLabels(iLabel).ExtraValue(1) = ""
        Else
            thisLabels(iLabel).ValueID(1) = "WRTS"
            thisLabels(iLabel).ExtraValueLength(1) = Len(sExtraValue)
            thisLabels(iLabel).ExtraValue(1) = sExtraValue
        End If
        If bANSI Then
            thisLabels(iLabel).ValueLength(1) = Len(sValue)
        Else
            thisLabels(iLabel).ValueLength(1) = Len(sValue) / 2
        End If
        thisLabels(iLabel).Value(1) = DecodeValue(sValue, bANSI)
    End If
End Sub

Public Sub UpdateWith(ByVal sFilePath As String, Optional ByVal bReplaceExisting As Boolean = True)
    Dim secondCSF As MarshallxCSFClass
    Dim iLabel As Long
    Dim hFile As Integer
    Dim sLabelValue As String
    Dim bUnicode As Boolean
    Dim sLabelName As String
    Dim sBytes() As Byte
    Dim sWord(1) As Byte
    Dim iPos As Long
    Dim iLen As Long
    Dim bValueStage As Boolean
    Select Case FileType(sFilePath)
    Case "CSF"
        Set secondCSF = New MarshallxCSFClass
        Call secondCSF.LoadCSF(sFilePath)
        iLabel = 1
        Do While iLabel <= secondCSF.LabelCount
            Call AddLabel(secondCSF.LabelName(iLabel), secondCSF.Value(iLabel), secondCSF.ExtraValue(iLabel), False, bReplaceExisting)
            iLabel = iLabel + 1
        Loop
        Set secondCSF = Nothing
    Case "TXT"
        bUnicode = False
        hFile = FreeFile()
        Open sFilePath For Binary As #hFile
        iLen = LOF(hFile)
        If iLen >= 3 Then 'must have at least 3 characters to contain one label in ansi, and at least 2 bytes to identify a unicode file
            iLen = iLen - 1 '(ubound of sbytes)
            ReDim sBytes(iLen)
            Get #hFile, , sBytes
            If sBytes(0) = 255 And sBytes(1) = 254 Then 'HEX(FFFE) - the UTF-16 LE BOM
                bUnicode = True
                iPos = 2 'next byte to read
            Else
                bUnicode = False
                iPos = 0 'next byte to read
            End If
            Close #hFile
            sLabelValue = ""
            sLabelName = ""
            bValueStage = False
            Do
                If bUnicode Then
                    If iPos + 1 > iLen Then
                        'no more characters. see if we have a complete label
                        If bValueStage Then
                            If Len(sLabelName) <> 0 Then
                                Call AddLabel(sLabelName, sLabelValue, "", False, bReplaceExisting)
                                'else label name is blank (invalid)
                            End If
                        End If
                        Exit Do 'no more characters
                    End If
                    sWord(0) = sBytes(iPos)
                    sWord(1) = sBytes(iPos + 1)
                    iPos = iPos + 2
                Else
                    If iPos > iLen Then
                        'no more characters. see if we have a complete label
                        If bValueStage Then
                            If Len(sLabelName) <> 0 Then
                                Call AddLabel(sLabelName, sLabelValue, "", False, bReplaceExisting)
                                'else label name is blank (invalid)
                            End If
                        End If
                        Exit Do 'no more characters
                    End If
                    sWord(0) = sBytes(iPos)
                    sWord(1) = 0
                    iPos = iPos + 1
                End If
                If sWord(1) = 0 Then
                    If Not bValueStage And sWord(0) = 61 Then '=
                        bValueStage = True
                    Else
                        Select Case sWord(0)
                        Case 13, 10 'CR, LF
                            If bValueStage Then
                                If Len(sLabelName) <> 0 Then
                                    Call AddLabel(sLabelName, sLabelValue, "", False, bReplaceExisting)
                                    'else label name is blank (invalid)
                                End If
                                'reset for next label
                                bValueStage = False
                                sLabelName = ""
                                sLabelValue = ""
                            End If
                        Case Else
                            If sWord(0) > 31 Or sWord(0) = 9 Then 'not a control char, unless its a tab
                                If Not bValueStage Then
                                    'label name can't have spaces or tabs
                                    If sWord(0) <> 32 And sWord(0) <> 9 Then '32=space, 9=tab
                                        sLabelName = sLabelName & Chr(sWord(0))
                                    End If
                                Else
                                    If sWord(0) = 36 Then '$
                                        'need to check if following characters are \n and replace with new line
                                        If bUnicode Then
                                            If (iPos + 2) < iLen Then
                                                If sBytes(iPos) = 92 And sBytes(iPos + 1) = 0 And sBytes(iPos + 2) = 110 & sBytes(iPos + 3) = 0 Then  '\n
                                                    sLabelValue = sLabelValue & Chr(13) & Chr(0)
                                                    sWord(0) = 10
                                                    iPos = iPos + 4
                                                End If
                                            End If
                                        Else
                                            If iPos < iLen Then
                                                If sBytes(iPos) = 92 And sBytes(iPos + 1) = 110 Then '\n
                                                    sLabelValue = sLabelValue & Chr(13) & Chr(0)
                                                    sWord(0) = 10
                                                    iPos = iPos + 2
                                                End If
                                            End If
                                        End If
                                    End If
                                    sLabelValue = sLabelValue & Chr(sWord(0)) & Chr(sWord(1))
                                End If
                            End If
                        End Select
                    End If
                Else
                    If bValueStage Then sLabelValue = sLabelValue & Chr(sWord(0)) & Chr(sWord(1))
                    'else not valid - unicode chars not allowed in label name
                End If
            Loop
        Else
            Close #hFile
        End If
    End Select
End Sub
