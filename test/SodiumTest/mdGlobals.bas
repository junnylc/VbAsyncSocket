Attribute VB_Name = "mdGlobals"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long

Public Function DesignDumpArray(baData() As Byte, Optional ByVal lPos As Long, Optional ByVal lSize As Long = -1) As String
    If lSize < 0 Then
        lSize = UBound(baData) + 1 - lPos
    End If
    If lSize > 0 Then
        DesignDumpArray = DesignDumpMemory(VarPtr(baData(lPos)), lSize)
    End If
End Function

Public Function DesignDumpMemory(ByVal lPtr As Long, ByVal lSize As Long) As String
    Dim lIdx            As Long
    Dim sHex            As String
    Dim sChar           As String
    Dim lValue          As Long
    Dim aResult()       As String
    
    ReDim aResult(0 To (lSize + 15) \ 16) As String
    For lIdx = 0 To ((lSize + 15) \ 16) * 16
        If lIdx < lSize Then
            If IsBadReadPtr(lPtr, 1) = 0 Then
                Call CopyMemory(lValue, ByVal lPtr, 1)
                sHex = sHex & Right$("0" & Hex$(lValue), 2) & " "
                If lValue >= 32 Then
                    sChar = sChar & Chr$(lValue)
                Else
                    sChar = sChar & "."
                End If
            Else
                sHex = sHex & "?? "
                sChar = sChar & "."
            End If
        Else
            sHex = sHex & "   "
        End If
        If ((lIdx + 1) Mod 4) = 0 Then
            sHex = sHex & " "
        End If
        If ((lIdx + 1) Mod 16) = 0 Then
            aResult(lIdx \ 16) = Right$("000" & Hex$(lIdx - 15), 4) & " - " & sHex & sChar
            sHex = vbNullString
            sChar = vbNullString
        End If
        lPtr = (lPtr Xor &H80000000) + 1 Xor &H80000000
    Next
    DesignDumpMemory = Join(aResult, vbCrLf)
End Function

Public Sub WriteBinaryFile(sFile As String, baBuffer() As Byte)
    Dim nFile           As Integer
    
    nFile = FreeFile
    Open sFile For Binary Access Write Shared As nFile
    If UBound(baBuffer) >= 0 Then
        Put nFile, , baBuffer
    End If
    Close nFile
End Sub

Public Function ToHex(baText() As Byte, Optional Delimiter As String = "-") As String
    Dim aText()         As String
    Dim lIdx            As Long
    
    If LenB(CStr(baText)) <> 0 Then
        ReDim aText(0 To UBound(baText)) As String
        For lIdx = 0 To UBound(baText)
            aText(lIdx) = Right$("0" & Hex$(baText(lIdx)), 2)
        Next
        ToHex = Join(aText, Delimiter)
    End If
End Function

Public Function FromHex(sText As String) As Byte()
    Dim baRetVal()      As Byte
    Dim lIdx            As Long
    
    On Error GoTo QH
    '--- check for hexdump delimiter
    If sText Like "*[!0-9A-Fa-f]*" Then
        ReDim baRetVal(0 To Len(sText) \ 3) As Byte
        For lIdx = 1 To Len(sText) Step 3
            baRetVal(lIdx \ 3) = "&H" & Mid$(sText, lIdx, 2)
        Next
    ElseIf LenB(sText) <> 0 Then
        ReDim baRetVal(0 To Len(sText) \ 2 - 1) As Byte
        For lIdx = 1 To Len(sText) Step 2
            baRetVal(lIdx \ 2) = "&H" & Mid$(sText, lIdx, 2)
        Next
    Else
        baRetVal = vbNullString
    End If
    FromHex = baRetVal
QH:
End Function


