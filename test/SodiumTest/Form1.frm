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
      Caption         =   "Download"
      Default         =   -1  'True
      Height          =   348
      Left            =   8400
      TabIndex        =   2
      Top             =   168
      Width           =   1356
   End
   Begin VB.TextBox txtUrl 
      Height          =   348
      Left            =   1260
      TabIndex        =   0
      Text            =   "cert-test.sandbox.google.com"
      Top             =   168
      Width           =   7068
   End
   Begin VB.Label Label1 
      Caption         =   "Address:"
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
Private Const MODULE_NAME As String = "Form1"

'=========================================================================
' API
'=========================================================================

'--- Windows Messages
Private Const WM_SETREDRAW              As Long = &HB
Private Const EM_SETSEL                 As Long = &HB1
Private Const EM_REPLACESEL             As Long = &HC2
Private Const WM_VSCROLL                As Long = &H115
'--- for WM_VSCROLL
Private Const SB_BOTTOM                 As Long = 7

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_oSocket               As cAsyncSocket
Private m_sServerName           As String
Private m_uCtx                  As UcsTlsContext
Private WithEvents m_oServerSocket As cAsyncSocket
Attribute m_oServerSocket.VB_VarHelpID = -1
Private m_cRequestHandlers      As Collection
Private m_cCertificates         As Collection
Private m_baPrivateKey()        As Byte
Private m_hRootStore            As Long

Private Type UcsParsedUrl
    Protocol        As String
    Host            As String
    Port            As Long
    Path            As String
    QueryString     As String
    Anchor          As String
    User            As String
    Pass            As String
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    #If ImplUseDebugLog Then
        DebugLog MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description & " &H" & Hex$(Err.Number), vbLogEventTypeError
    #Else
        Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    #End If
End Sub

'=========================================================================
' Events
'=========================================================================

Private Sub Form_Load()
    Const PEM_FILES     As String = "eccert.pem|ecprivkey.pem|fullchain2.pem"
    Const PFX_FILE      As String = "eccert.pfx"
    Const PFX_PASSWORD  As String = ""
    Dim sAddr           As String
    Dim lPort           As Long
    
    On Error GoTo EH
    If txtResult.Font.Name = "Arial" Then
        txtResult.Font.Name = "Courier New"
    End If
    sAddr = GetSetting(App.Title, "Form1", "txtUrl", txtUrl.Text)
    If LenB(sAddr) <> 0 Then
        txtUrl.Text = sAddr
    End If
    m_hRootStore = PkiPemImportRootCaCertStore(App.Path & "\ca-bundle.pem")
    ChDir App.Path
'    If Not PkiPemImportCertificates(Split(PEM_FILES, "|"), m_cCertificates, m_baPrivateKey) Then
'    If Not PkiPkcs12ImportCertificates(PFX_FILE, PFX_PASSWORD, m_cCertificates, m_baPrivateKey) Then
    If Not PkiGenerSelfSignedCertificate(m_cCertificates, m_baPrivateKey) Then
        MsgBox "Error starting TLS server on localhost:10443" & vbCrLf & vbCrLf & "No private key found!", vbExclamation
        GoTo QH
    End If
    Set m_oServerSocket = New cAsyncSocket
    m_oServerSocket.Create SocketPort:=10443, SocketAddress:="localhost"
    If m_oServerSocket.Listen() Then
        Set m_cRequestHandlers = New Collection
        m_oServerSocket.GetSockName sAddr, lPort
        Debug.Print "Listening on " & sAddr & ":" & lPort
    End If
QH:
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical
    Resume QH
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        txtResult.Move 0, txtResult.Top, ScaleWidth, ScaleHeight - txtResult.Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Form1", "txtUrl", txtUrl.Text
    Set m_cRequestHandlers = Nothing
End Sub

Private Sub Command1_Click()
    Dim uRemote         As UcsParsedUrl
    Dim sResult         As String
    Dim sError          As String
    Dim bKeepDebug      As Boolean
    
    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    bKeepDebug = IsKeyPressed(vbKeyControl)
    ' tls13.1d.pw, localhost:44330, tls.ctf.network, www.mikestoolbox.org, swifttls.org, tls13.pinterjann.is, cert-test.sandbox.google.com
    If Not ParseUrl(Trim$(txtUrl.Text), uRemote, DefProtocol:="https") Then
        txtResult.Text = "Error: Invalid URL"
        GoTo QH
    End If
    txtResult.Text = vbNullString
    sResult = HttpsRequest(m_uCtx, uRemote, sError)
    If LenB(sError) <> 0 Then
        pvAppendLogText txtResult, "Error: " & sError
        GoTo QH
    End If
    If LenB(sResult) <> 0 Then
        If Not bKeepDebug Then
            txtResult.Text = vbNullString
        End If
        pvAppendLogText txtResult, sResult
        txtResult.SelStart = 0
    End If
QH:
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description & " [" & Err.Source & "]", vbCritical
    Set m_oSocket = Nothing
End Sub

'=========================================================================
' Methods
'=========================================================================

Private Function HttpsRequest(uCtx As UcsTlsContext, uRemote As UcsParsedUrl, sError As String) As String
    Const INDENT        As Long = 4
    Dim baRecv()        As Byte
    Dim sRequest        As String
    Dim baSend()        As Byte
    Dim lSize           As Long
    Dim baDecr()        As Byte
    Dim dblTimer        As Double
    
    pvAppendLogText txtResult, "Connecting to " & uRemote.Host & vbCrLf
    If m_sServerName <> uRemote.Host & ":" & uRemote.Port Or m_oSocket Is Nothing Then
        Set m_oSocket = New cAsyncSocket
        If Not m_oSocket.SyncConnect(uRemote.Host, uRemote.Port) Then
            sError = m_oSocket.GetErrorDescription(m_oSocket.LastError)
            GoTo QH
        End If
        m_sServerName = uRemote.Host & ":" & uRemote.Port
        '--- send TLS handshake
        If Not TlsInitClient(uCtx, RemoteHostName:=uRemote.Host) Then ' , LocalFeatures:=ucsTlsSupportTls12)
            sError = TlsGetLastError(uCtx)
            GoTo QH
        End If
        GoTo InLoop
        Do
            If Not m_oSocket.SyncReceiveArray(baRecv, Timeout:=1000) Then
                sError = m_oSocket.GetErrorDescription(m_oSocket.LastError)
                GoTo QH
            End If
            If pvArraySize(baRecv) <> 0 Then
                pvAppendLogText txtResult, String$(INDENT - 1, ">") & " Recv " & pvArraySize(baRecv) & _
                    RTrim$(Replace(vbCrLf & DesignDumpMemory(VarPtr(baRecv(0)), pvArraySize(baRecv)), vbCrLf, vbCrLf & Space$(INDENT)))
            End If
InLoop:
            lSize = 0
            If Not mdTlsSupport.TlsHandshake(uCtx, baRecv, -1, baSend, lSize) Then
                sError = TlsGetLastError(uCtx)
                GoTo QH
            End If
            If lSize > 0 Then
                pvAppendLogText txtResult, String$(INDENT - 1, "<") & " Send " & lSize & _
                    RTrim$(Replace(vbCrLf & DesignDumpMemory(VarPtr(baSend(0)), lSize), vbCrLf, vbCrLf & Space$(INDENT)))
                If Not m_oSocket.SyncSend(VarPtr(baSend(0)), lSize) Then
                    sError = m_oSocket.GetErrorDescription(m_oSocket.LastError)
                    GoTo QH
                End If
            End If
            If TlsIsClosed(uCtx) Then
                sError = "Unexpected TLS session close"
                GoTo QH
            End If
        Loop While Not TlsIsReady(uCtx)
        If Not PkiCertChainValidate(uRemote.Host, m_uCtx.RemoteCertificates, m_hRootStore, sError) Then
            If MsgBox(sError & vbCrLf & vbCrLf & "Do you want to continue?", vbQuestion Or vbYesNo) = vbNo Then
                GoTo QH
            End If
            sError = vbNullString
        End If
    End If
    '--- send TLS application data and wait for reply
    sRequest = "GET " & uRemote.Path & uRemote.QueryString & " HTTP/1.1" & vbCrLf & _
               "Connection: keep-alive" & vbCrLf & _
               "Host: " & uRemote.Host & vbCrLf & vbCrLf
    lSize = 0
    If Not mdTlsSupport.TlsSend(uCtx, StrConv(sRequest, vbFromUnicode), -1, baSend, lSize) Then
        sError = TlsGetLastError(uCtx)
        GoTo QH
    End If
    If lSize > 0 Then
        pvAppendLogText txtResult, String$(INDENT - 1, "<") & " Send " & lSize & _
            RTrim$(Replace(vbCrLf & DesignDumpMemory(VarPtr(baSend(0)), lSize), vbCrLf, vbCrLf & Space$(INDENT)))
        If Not m_oSocket.SyncSend(VarPtr(baSend(0)), lSize) Then
            sError = m_oSocket.GetErrorDescription(m_oSocket.LastError)
            GoTo QH
        End If
    End If
    lSize = 0
    dblTimer = Timer
    Do
        If Not m_oSocket.ReceiveArray(baRecv) Then
            sError = m_oSocket.GetErrorDescription(m_oSocket.LastError)
            GoTo QH
        End If
        If pvArraySize(baRecv) <> 0 Then
            pvAppendLogText txtResult, String$(INDENT - 1, ">") & " Recv " & pvArraySize(baRecv) & _
                RTrim$(Replace(vbCrLf & DesignDumpMemory(VarPtr(baRecv(0)), pvArraySize(baRecv)), vbCrLf, vbCrLf & Space$(INDENT)))
            dblTimer = Timer
        ElseIf lSize > 0 And Timer > dblTimer + 0.2 Then
            Exit Do
        ElseIf Timer > dblTimer + 2 Then
            sError = "Timeout waiting for response for " & Format$(Timer - dblTimer, "0.000") & " seconds"
            Exit Do
        End If
        If Not mdTlsSupport.TlsReceive(uCtx, baRecv, -1, baDecr, lSize) Then
            sError = TlsGetLastError(uCtx)
            GoTo QH
        End If
        If TlsIsClosed(uCtx) Then
            Set m_oSocket = Nothing
            Exit Do
        End If
    Loop
    HttpsRequest = Replace(Replace(StrConv(baDecr, vbUnicode), vbCr, vbNullString), vbLf, vbCrLf)
    lSize = InStr(1, HttpsRequest, vbCrLf & vbCrLf)
    If Not m_oSocket Is Nothing Then
        If lSize > 0 Then
            If InStr(1, Left$(HttpsRequest, lSize), "Connection: close", vbTextCompare) = 0 Then
                '--- keep TLS session
                GoTo QH
            End If
        End If
        lSize = 0
        If Not mdTlsSupport.TlsShutdown(uCtx, baSend, lSize) Then
            sError = TlsGetLastError(uCtx)
            GoTo QH
        End If
        If lSize > 0 Then
            pvAppendLogText txtResult, String$(INDENT - 1, "<") & " Send " & lSize & _
                RTrim$(Replace(vbCrLf & DesignDumpMemory(VarPtr(baSend(0)), lSize), vbCrLf, vbCrLf & Space$(INDENT)))
            If Not m_oSocket.SyncSend(VarPtr(baSend(0)), lSize) Then
                sError = m_oSocket.GetErrorDescription(m_oSocket.LastError)
                GoTo QH
            End If
        End If
        Set m_oSocket = Nothing
    End If
QH:
'    HttpsRequest = vbNullString
    If LenB(sError) <> 0 Then
        Set m_oSocket = Nothing
    End If
End Function

Private Function ParseUrl(sUrl As String, uParsed As UcsParsedUrl, Optional DefProtocol As String) As Boolean
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = "^(?:(.*)://)?(?:(?:([^:]*):)?([^@]*)@)?([A-Za-z0-9\-\.]+)(:[0-9]+)?(/[^?#]*)?(\?[^#]*)?(#.*)?$"
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
                    uParsed.QueryString = .Item(6)
                    uParsed.Anchor = .Item(7)
                End With
                ParseUrl = True
            End If
        End With
    End With
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

Private Sub pvAppendLogText(txtLog As TextBox, sValue As String)
    Call SendMessage(txtLog.hWnd, WM_SETREDRAW, 0, ByVal 0)
    Call SendMessage(txtLog.hWnd, EM_SETSEL, 0, ByVal -1)
    Call SendMessage(txtLog.hWnd, EM_SETSEL, -1, ByVal -1)
    Call SendMessage(txtLog.hWnd, EM_REPLACESEL, 1, ByVal sValue)
    Call SendMessage(txtLog.hWnd, EM_SETSEL, 0, ByVal -1)
    Call SendMessage(txtLog.hWnd, EM_SETSEL, -1, ByVal -1)
    Call SendMessage(txtLog.hWnd, WM_SETREDRAW, 1, ByVal 0)
    Call SendMessage(txtLog.hWnd, WM_VSCROLL, SB_BOTTOM, ByVal 0)
End Sub

Public Function IsKeyPressed(ByVal lVirtKey As KeyCodeConstants) As Boolean
    IsKeyPressed = ((GetAsyncKeyState(lVirtKey) And &H8000) = &H8000)
End Function

Private Sub m_oServerSocket_OnAccept()
    Const FUNC_NAME     As String = "m_oServerSocket_OnAccept"
    Dim oSocket         As cAsyncSocket
    Dim oHandler        As cRequestHandler
    Dim sKey            As String
    
    On Error GoTo EH
    If Not m_oServerSocket.Accept(oSocket) Then
        GoTo QH
    End If
    Set oHandler = New cRequestHandler
    sKey = "#" & ObjPtr(oHandler)
    If Not oHandler.Init(oSocket, sKey, Me, m_cCertificates, m_baPrivateKey) Then
        GoTo QH
    End If
    m_cRequestHandlers.Add oHandler, "#" & ObjPtr(oHandler)
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Friend Sub frRemoveHandler(sKey As String)
    RemoveCollection m_cRequestHandlers, sKey
End Sub
