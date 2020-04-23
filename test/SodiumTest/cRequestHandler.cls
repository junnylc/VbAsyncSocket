VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRequestHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MODULE_NAME As String = "cRequestHandler"

Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long

Private WithEvents m_oSocket    As cAsyncSocket
Attribute m_oSocket.VB_VarHelpID = -1
Private m_uCtx                  As UcsTlsContext
Private m_sKey                  As String
Private m_lParentWeakRef        As Long

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
' Properties
'=========================================================================

Property Get Parent() As Form1
    Call vbaObjSetAddref(Parent, m_lParentWeakRef)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(oSocket As cAsyncSocket, sKey As String, oParent As Form1, cCertificates As Collection, baCertKey() As Byte) As Boolean
    Dim sAddr           As String
    Dim lPort           As Long
    
    Set m_oSocket = oSocket
    m_sKey = sKey
    m_lParentWeakRef = ObjPtr(oParent)
    oSocket.GetPeerName sAddr, lPort
    If Not mdTlsSupport.TlsInitServer(m_uCtx, RemoteHostName:=sAddr & ":" & lPort, Certificates:=cCertificates, CertKey:=baCertKey) Then
        GoTo QH
    End If
    '--- success
    Init = True
QH:
End Function

Private Function HandleRequest(sText As String) As String
    Dim vSplit          As Variant
    Dim sRetVal         As String
    
    vSplit = Split(sText, vbCrLf)
    vSplit = Split(vSplit(0), " ")
    Debug.Print "HandleRequestm Path=" & vSplit(1)
    sRetVal = "<html><body>" & Date & "</body></html>"
    HandleRequest = "HTTP/1.0 200 Ok" & vbCrLf & _
        "Content-Type: text/html; charset=UTF-8" & vbCrLf & _
        "Content-Length: " & Len(sRetVal) & vbCrLf & _
        "Connection: Close" & vbCrLf & vbCrLf & _
        sRetVal
End Function

Private Sub m_oSocket_OnClose()
    Parent.frRemoveHandler m_sKey
End Sub

Private Sub m_oSocket_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    Parent.frRemoveHandler m_sKey
End Sub

Private Sub m_oSocket_OnReceive()
    Const FUNC_NAME     As String = "m_oSocket_OnReceive"
    Dim baRecv()        As Byte
    Dim baBuffer()      As Byte
    Dim lSize           As Long
    Dim sError          As String
    Dim sText           As String
    
    On Error GoTo EH
    If Not m_oSocket.ReceiveArray(baRecv) Then
        GoTo QH
    End If
    If Not TlsIsReady(m_uCtx) Then
        If Not mdTlsSupport.TlsHandshake(m_uCtx, baRecv, -1, baBuffer, lSize) Then
            sError = TlsGetLastError(m_uCtx)
        ElseIf lSize > 0 Then
            Call m_oSocket.Send(VarPtr(baBuffer(0)), lSize)
            If m_oSocket.LastError <> 0 Then
                sError = m_oSocket.GetErrorDescription(m_oSocket.LastError)
                GoTo QH
            End If
        End If
    Else
        If Not mdTlsSupport.TlsReceive(m_uCtx, baRecv, -1, baBuffer, lSize) Then
            sError = TlsGetLastError(m_uCtx)
        ElseIf lSize > 0 Then
            If UBound(baBuffer) <> lSize - 1 Then
                ReDim Preserve baBuffer(0 To lSize - 1) As Byte
            End If
            sText = HandleRequest(FromUtf8Array(baBuffer))
        End If
        lSize = 0
        If Not mdTlsSupport.TlsSend(m_uCtx, m_oSocket.ToTextArray(sText, ucsScpUtf8), -1, baBuffer, lSize) Then
            If LenB(sError) = 0 Then
                sError = TlsGetLastError(m_uCtx)
            End If
        ElseIf lSize > 0 Then
            Call m_oSocket.Send(VarPtr(baBuffer(0)), lSize)
            If m_oSocket.LastError <> 0 Then
                sError = m_oSocket.GetErrorDescription(m_oSocket.LastError)
                GoTo QH
            End If
        End If
    End If
QH:
    If LenB(sError) <> 0 Then
        Debug.Print "Error in request handler: " & sError
    End If
    If TlsIsClosed(m_uCtx) Then
        m_oSocket.Close_
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub
