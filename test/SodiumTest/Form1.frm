VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5568
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9948
   LinkTopic       =   "Form1"
   ScaleHeight     =   5568
   ScaleWidth      =   9948
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResult 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   10.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4800
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   588
      Width           =   9756
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Default         =   -1  'True
      Height          =   348
      Left            =   3864
      TabIndex        =   2
      Top             =   168
      Width           =   1356
   End
   Begin VB.TextBox txtUrl 
      Height          =   348
      Left            =   1260
      TabIndex        =   0
      Text            =   "tls13.1d.pw"
      Top             =   168
      Width           =   2532
   End
   Begin VB.Label Label1 
      Caption         =   "Server:"
      Height          =   348
      Left            =   168
      TabIndex        =   1
      Top             =   168
      Width           =   936
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
'--- libsodium
Private Declare Function sodium_init Lib "libsodium" () As Long

Private Type UcsParsedUrl
    Protocol        As String
    Host            As String
    Port            As Long
    Path            As String
    User            As String
    Pass            As String
End Type

'=========================================================================
' Events
'=========================================================================

Private Sub Form_Load()
    If GetModuleHandle("libsodium.dll") = 0 Then
        Call LoadLibrary(App.Path & "\libsodium.dll")
        Call sodium_init
    End If
    If txtResult.Font.Name = "Arial" Then
        txtResult.Font.Name = "Courier New"
    End If
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        txtResult.Move 0, txtResult.Top, ScaleWidth, ScaleHeight - txtResult.Top
    End If
End Sub

Private Sub Command1_Click()
    Dim uRemote         As UcsParsedUrl
    Dim uCtx            As UcsClientContextType
    Dim sResult         As String
    Dim sError          As String
    
    On Error GoTo EH
    ' tls13.1d.pw, localhost:44330
    If Not ParseUrl(txtUrl.Text, uRemote, DefProtocol:="https") Then
        MsgBox "Wrong URL", vbCritical
        GoTo QH
    End If
    uCtx = TlsInitClient(uRemote.Host)
    sResult = HttpsRequest(uCtx, uRemote.Host & ":" & uRemote.Port, uRemote.Path, sError)
    If LenB(sError) <> 0 Then
        MsgBox sError, vbCritical
        GoTo QH
    End If
    txtResult.Text = sResult
QH:
    Exit Sub
EH:
    MsgBox Err.Description & " [" & Err.Source & "]", vbCritical
End Sub

'=========================================================================
' Methods
'=========================================================================

Private Function HttpsRequest(uCtx As UcsClientContextType, sServer As String, sPath As String, sError As String) As String
    Dim oSocket         As cAsyncSocket
    Dim vSplit          As Variant
    Dim baRecv()        As Byte
    Dim sRequest        As String
    Dim baSend()        As Byte
    Dim lSize           As Long
    Dim bComplete       As Boolean
    Dim baDecr()        As Byte
    
    Set oSocket = New cAsyncSocket
    vSplit = Split(sServer & ":443", ":")
    If Not oSocket.SyncConnect(CStr(vSplit(0)), Val(vSplit(1))) Then
        sError = oSocket.GetErrorDescription(oSocket.LastError)
        GoTo QH
    End If
    Do
        If Not oSocket.ReceiveArray(baRecv) Then
            sError = oSocket.GetErrorDescription(oSocket.LastError)
            GoTo QH
        End If
        If pvArraySize(baRecv) <> 0 Then
            txtResult.Text = "pvArraySize(baRecv)=" & pvArraySize(baRecv) & vbCrLf & DesignDumpMemory(VarPtr(baRecv(0)), pvArraySize(baRecv))
        End If
        lSize = 0
        If Not TlsHandshake(uCtx, baRecv, -1, baSend, lSize, bComplete) Then
            sError = TlsGetLastError(uCtx)
            GoTo QH
        End If
        If lSize > 0 Then
            If Not oSocket.SyncSend(VarPtr(baSend(0)), lSize) Then
                sError = oSocket.GetErrorDescription(oSocket.LastError)
                GoTo QH
            End If
        End If
    Loop While Not bComplete
    sRequest = "GET " & sPath & " HTTP/1.0" & vbCrLf & _
               "Host: " & vSplit(0) & vbCrLf & vbCrLf
    lSize = 0
    If Not TlsSend(uCtx, StrConv(sRequest, vbFromUnicode), -1, baSend, lSize) Then
        sError = TlsGetLastError(uCtx)
        GoTo QH
    End If
    If lSize > 0 Then
        If Not oSocket.SyncSend(VarPtr(baSend(0)), lSize) Then
            sError = oSocket.GetErrorDescription(oSocket.LastError)
            GoTo QH
        End If
    End If
    Do
        If Not oSocket.ReceiveArray(baRecv) Then
            sError = oSocket.GetErrorDescription(oSocket.LastError)
            GoTo QH
        End If
        If pvArraySize(baRecv) <> 0 Then
            txtResult.Text = "pvArraySize(baRecv)=" & pvArraySize(baRecv) & vbCrLf & DesignDumpMemory(VarPtr(baRecv(0)), pvArraySize(baRecv))
        End If
        lSize = 0
        If Not TlsReceive(uCtx, baRecv, -1, baDecr, lSize) Then
            sError = TlsGetLastError(uCtx)
            GoTo QH
        End If
    Loop While lSize = 0
    HttpsRequest = Replace(Replace(StrConv(baDecr, vbUnicode), vbCr, vbNullString), vbLf, vbCrLf)
QH:
End Function

Private Function ParseUrl(sUrl As String, uParsed As UcsParsedUrl, Optional DefProtocol As String) As Boolean
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = "^(?:(.*)://)?(?:(?:([^:]*):)?([^@]*)@)?([A-Za-z0-9\-\.]+)(:[0-9]+)?(.*)$"
        With .Execute(sUrl)
            If .Count > 0 Then
                With .Item(0).SubMatches
                    uParsed.Protocol = .Item(0)
                    uParsed.User = .Item(1)
                    If LenB(uParsed.User) = 0 Then
                        uParsed.User = .Item(2)
                    Else
                        uParsed.Pass = .Item(2)
                    End If
                    uParsed.Host = .Item(3)
                    uParsed.Port = Val(Mid$(.Item(4), 2))
                    If uParsed.Port = 0 Then
                        Select Case LCase$(IIf(LenB(uParsed.Protocol) = 0, DefProtocol, uParsed.Protocol))
                        Case "https"
                            uParsed.Port = 443
                        Case "socks5"
                            uParsed.Port = 1080
                        Case Else
                            uParsed.Port = 80
                        End Select
                    End If
                    uParsed.Path = .Item(5)
                    If LenB(uParsed.Path) = 0 Then
                        uParsed.Path = "/"
                    End If
                End With
                ParseUrl = True
            End If
        End With
    End With
End Function

Private Function DesignDumpMemory(ByVal lPtr As Long, ByVal lSize As Long) As String
    Dim lIdx            As Long
    Dim sHex            As String
    Dim sChar           As String
    Dim lValue          As Long

    For lIdx = 0 To ((lSize + 15) \ 16) * 16
        If lIdx < lSize Then
            If IsBadReadPtr(UnsignedAdd(lPtr, lIdx), 1) = 0 Then
                Call CopyMemory(lValue, ByVal UnsignedAdd(lPtr, lIdx), 1)
                sHex = sHex & Right$("00" & Hex$(lValue), 2) & " "
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
        If ((lIdx + 1) Mod 16) = 0 Then
            DesignDumpMemory = DesignDumpMemory & Right$("0000" & Hex$(lIdx - 15), 4) & ": " & sHex & " " & sChar & vbCrLf
            sHex = vbNullString
            sChar = vbNullString
        End If
    Next
End Function

Private Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function

Private Function pvArraySize(baArray() As Byte, Optional RetVal As Long) As Long
    Dim lPtr            As Long
    
    '--- peek long at ArrPtr(baArray)
    Call CopyMemory(lPtr, ByVal ArrPtr(baArray), 4)
    If lPtr <> 0 Then
        RetVal = UBound(baArray) + 1
    Else
        RetVal = 0
    End If
    pvArraySize = RetVal
End Function
