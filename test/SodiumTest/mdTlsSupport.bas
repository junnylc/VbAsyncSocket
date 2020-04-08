Attribute VB_Name = "mdTlsSupport"
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Const TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC       As Long = 20
Private Const TLS_CONTENT_TYPE_ALERT                    As Long = 21
Private Const TLS_CONTENT_TYPE_HANDSHAKE                As Long = 22
Private Const TLS_CONTENT_TYPE_APPDATA                  As Long = 23
Private Const TLS_RECORD_VERSION                        As Long = &H303
Private Const TLS_CLIENT_LEGACY_VERSION                 As Long = &H303
Private Const TLS_HANDSHAKE_TYPE_CLIENT_HELLO           As Long = 1
Private Const TLS_HANDSHAKE_TYPE_SERVER_HELLO           As Long = 2
Private Const TLS_HANDSHAKE_TYPE_NEW_SESSION_TICKET     As Long = 4
'Private Const TLS_HANDSHAKE_TYPE_END_OF_EARLY_DATA      As Long = 5
Private Const TLS_HANDSHAKE_TYPE_ENCRYPTED_EXTENSIONS   As Long = 8
Private Const TLS_HANDSHAKE_TYPE_CERTIFICATE            As Long = 11
'Private Const TLS_HANDSHAKE_TYPE_CERTIFICATE_REQUEST    As Long = 13
Private Const TLS_HANDSHAKE_TYPE_CERTIFICATE_VERIFY     As Long = 15
Private Const TLS_HANDSHAKE_TYPE_FINISHED               As Long = 20
Private Const TLS_HANDSHAKE_TYPE_KEY_UPDATE             As Long = 24
Private Const TLS_HANDSHAKE_TYPE_COMPRESSED_CERTIFICATE As Long = 25
'Private Const TLS_HANDSHAKE_TYPE_MESSAGE_HASH           As Long = 254
Private Const TLS_EXTENSION_TYPE_SERVER_NAME            As Long = 0
'Private Const TLS_EXTENSION_TYPE_STATUS_REQUEST         As Long = 5
Private Const TLS_EXTENSION_TYPE_SUPPORTED_GROUPS       As Long = 10
Private Const TLS_EXTENSION_TYPE_SIGNATURE_ALGORITHMS   As Long = 13
'Private Const TLS_EXTENSION_TYPE_ALPN                   As Long = 16
'Private Const TLS_EXTENSION_TYPE_COMPRESS_CERTIFICATE   As Long = 27
'Private Const TLS_EXTENSION_TYPE_PRE_SHARED_KEY         As Long = 41
'Private Const TLS_EXTENSION_TYPE_EARLY_DATA             As Long = 42
Private Const TLS_EXTENSION_TYPE_SUPPORTED_VERSIONS     As Long = 43
'Private Const TLS_EXTENSION_TYPE_COOKIE                 As Long = 44
Private Const TLS_EXTENSION_TYPE_PSK_KEY_EXCHANGE_MODES As Long = 45
Private Const TLS_EXTENSION_TYPE_KEY_SHARE              As Long = 51
'Private Const TLS_CIPHER_SUITE_AES_128_GCM_SHA256       As Long = &H1301
Private Const TLS_CIPHER_SUITE_AES_256_GCM_SHA384       As Long = &H1302
Private Const TLS_CIPHER_SUITE_CHACHA20_POLY1305_SHA256 As Long = &H1303
'Private Const TLS_GROUP_SECP256R1                       As Long = 23
'Private Const TLS_GROUP_SECP384R1                       As Long = 24
'Private Const TLS_GROUP_SECP521R1                       As Long = 25
Private Const TLS_GROUP_X25519                          As Long = 29
'Private Const TLS_GROUP_X448                            As Long = 30
Private Const TLS_SIGNATURE_RSA_PKCS1_SHA1              As Long = &H201
Private Const TLS_SIGNATURE_RSA_PKCS1_SHA256            As Long = &H401
Private Const TLS_SIGNATURE_ECDSA_SECP256R1_SHA256      As Long = &H403
'Private Const TLS_SIGNATURE_ECDSA_SECP384R1_SHA384      As Long = &H503
'Private Const TLS_SIGNATURE_ECDSA_SECP521R1_SHA512      As Long = &H603
Private Const TLS_SIGNATURE_RSA_PSS_RSAE_SHA256         As Long = &H804
'Private Const TLS_SIGNATURE_RSA_PSS_RSAE_SHA384         As Long = &H805
'Private Const TLS_SIGNATURE_RSA_PSS_RSAE_SHA512         As Long = &H806
Private Const TLS_PSK_KE_MODE_PSK_DHE                   As Long = 1
Private Const TLS_PROTOCOL_VERSION_TLS13_FINAL          As Long = &H304
Private Const TLS_CHACHA20_KEY_SIZE                     As Long = 32
Private Const TLS_CHACHA20POLY1305_IV_SIZE              As Long = 12
Private Const TLS_CHACHA20POLY1305_TAG_SIZE             As Long = 16
Private Const TLS_AES256_KEY_SIZE                       As Long = 32
Private Const TLS_AESGCM_IV_SIZE                        As Long = 12
Private Const TLS_AESGCM_TAG_SIZE                       As Long = 16
Private Const TLS_COMPRESS_NULL                         As Long = 0
Private Const TLS_SERVER_NAME_TYPE_HOSTNAME             As Long = 0
Private Const TLS_ALERT_LEVEL_WARNING                   As Long = 1
Private Const TLS_ALERT_LEVEL_FATAL                     As Long = 2
Private Const TLS_SHA256_DIGEST_SIZE                    As Long = 32
Private Const TLS_SHA384_DIGEST_SIZE                    As Long = 48
Private Const TLS_X25519_KEY_SIZE                       As Long = 32
Private Const TLS_MAX_PLAINTEXT_RECORD_SIZE             As Long = 16384
Private Const TLS_MAX_ENCRYPTED_RECORD_SIZE             As Long = (16384 + 256)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (Source As Any, Destination As Any) As Long
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
'--- libsodium
Private Declare Function randombytes_buf Lib "libsodium" (lpOut As Any, ByVal lSize As Long) As Long
Private Declare Function crypto_scalarmult_curve25519 Lib "libsodium" (lpOut As Any, lpConstN As Any, lpConstP As Any) As Long
Private Declare Function crypto_scalarmult_curve25519_base Lib "libsodium" (lpOut As Any, lpConstN As Any) As Long
Private Declare Function crypto_hash_sha256 Lib "libsodium" (lpOut As Any, lpConstIn As Any, ByVal lSize As Long, Optional ByVal lHighSize As Long) As Long
Private Declare Function crypto_auth_hmacsha256 Lib "libsodium" (lpOut As Any, lpConstIn As Any, ByVal lSize As Long, ByVal lHighSize As Long, lpConstKey As Any) As Long
Private Declare Function crypto_aead_chacha20poly1305_ietf_decrypt Lib "libsodium" (lpOut As Any, lOutSize As Any, ByVal nSec As Long, lConstIn As Any, ByVal lInSize As Long, ByVal lHighInSize As Long, lpConstAd As Any, ByVal lAdSize As Long, ByVal lHighAdSize As Long, lpConstNonce As Any, lpConstKey As Any) As Long
Private Declare Function crypto_aead_chacha20poly1305_ietf_encrypt Lib "libsodium" (lpOut As Any, lOutSize As Any, lConstIn As Any, ByVal lInSize As Long, ByVal lHighInSize As Long, lpConstAd As Any, ByVal lAdSize As Long, ByVal lHighAdSize As Long, ByVal nSec As Long, lpConstNonce As Any, lpConstKey As Any) As Long
'Private Declare Function crypto_aead_aes256gcm_decrypt Lib "libsodium" (lpOut As Any, lOutSize As Any, ByVal nSec As Long, lConstIn As Any, ByVal lInSize As Long, ByVal lHighInSize As Long, lpConstAd As Any, ByVal lAdSize As Long, ByVal lHighAdSize As Long, lpConstNonce As Any, lpConstKey As Any) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Public Enum UcsClientStatesEnum
    ucsStateHandshakeStart
    ucsStateExpectServerHello
    ucsStateExpectEncryptedExtensions
    ucsStatePostHandshake
End Enum

Public Enum UcsCryptoAlgorithmsEnum
    '--- key exchange
    ucsAlgoKeyX25519
    '--- authenticated encryption w/ additional data
    ucsAlgoAeadAes128
    ucsAlgoAeadChacha20Poly1305
    ucsAlgoAeadAes256
    '--- digest
    ucsAlgoDigestSha256
    ucsAlgoDigestSha384
End Enum

Public Type UcsClientContextType
    ServerName                  As String
    DebugBox                    As TextBox
    
    LegacySessionID()           As Byte '--- not used
    ClientRandom()              As Byte
    ClientPrivate()             As Byte
    ClientPublic()              As Byte
    ServerRandom()              As Byte '--- not used
    ServerPublic()              As Byte
    ServerSupportedVersion      As Long '--- not used
    ServerCertificate()         As Byte
    
    HandshakeMessages()         As Byte
    HandshakeSecret()           As Byte
    MasterSecret()              As Byte
    ServerTrafficSecret()       As Byte
    ServerTrafficKey()          As Byte
    ServerTrafficIV()           As Byte
    ServerTrafficSeqNo          As Long
    ClientTrafficSecret()       As Byte
    ClientTrafficKey()          As Byte
    ClientTrafficIV()           As Byte
    ClientTrafficSeqNo          As Long
    PrevServerTrafficKey()      As Byte
    PrevServerTrafficIV()       As Byte
    PrevServerTrafficSeqNo      As Long
    
    State                       As UcsClientStatesEnum
    
    KxAlgo                      As UcsCryptoAlgorithmsEnum
    SecretSize                  As Long
    CipherSuite                 As Long
    AeadAlgo                    As UcsCryptoAlgorithmsEnum
    KeySize                     As Long
    IvSize                      As Long
    TagSize                     As Long
    DigestAlgo                  As UcsCryptoAlgorithmsEnum
    DigestSize                  As Long

    RecvBuffer()                As Byte
    RecvPos                     As Long
    DecrBuffer()                As Byte
    DecrPos                     As Long
    SendBuffer()                As Byte
    SendPos                     As Long
End Type

'=========================================================================
' Methods
'=========================================================================

Public Function TlsInitClient(sServerName As String, oDebugBox As TextBox) As UcsClientContextType
    Dim uRetVal         As UcsClientContextType
    
    With uRetVal
        .ServerName = sServerName
        Set .DebugBox = oDebugBox
'        .LegacySessionID = pvCryptoRandomBytes(32)
        '--- setup key exchange ephemeral priv/pub keys
        .KxAlgo = ucsAlgoKeyX25519
        .SecretSize = TLS_X25519_KEY_SIZE
        .ClientRandom = pvCryptoRandomBytes(.SecretSize)
        If .KxAlgo = ucsAlgoKeyX25519 Then
            .ClientPrivate = pvCryptoRandomBytes(TLS_X25519_KEY_SIZE)
            '--- fix some issues w/ specific privkeys
            .ClientPrivate(0) = .ClientPrivate(0) And 248
            .ClientPrivate(UBound(.ClientPrivate)) = (.ClientPrivate(UBound(.ClientPrivate)) And 127) Or 64
            ReDim .ClientPublic(0 To TLS_X25519_KEY_SIZE - 1) As Byte
            Call crypto_scalarmult_curve25519_base(.ClientPublic(0), .ClientPrivate(0))
        Else
            Err.Raise vbObjectError, "Unsupported key-exchange type"
        End If
    End With
    TlsInitClient = uRetVal
End Function

Public Function TlsFetchHttp(uCtx As UcsClientContextType, sPath As String, sError As String) As String
    Dim oSocket         As cAsyncSocket
    Dim vSplit          As Variant
    Dim baRecv()        As Byte
    Dim sRequest        As String
    
    Set oSocket = New cAsyncSocket
    vSplit = Split(uCtx.ServerName & ":443", ":")
    If Not oSocket.SyncConnect(CStr(vSplit(0)), Val(vSplit(1))) Then
        sError = oSocket.GetErrorDescription(oSocket.LastError)
        GoTo QH
    End If
    uCtx.SendPos = pvSendClientHello(uCtx, uCtx.SendBuffer, uCtx.SendPos)
    If uCtx.SendPos > 0 Then
        pvWriteBuffer uCtx.HandshakeMessages, 0, VarPtr(uCtx.SendBuffer(5)), uCtx.SendPos - 5
        If Not oSocket.SyncSend(VarPtr(uCtx.SendBuffer(0)), uCtx.SendPos) Then
            sError = oSocket.GetErrorDescription(oSocket.LastError)
            GoTo QH
        End If
        uCtx.SendPos = 0
    End If
    uCtx.State = ucsStateExpectServerHello
    Do
        If Not oSocket.SyncReceiveArray(baRecv, Timeout:=-1) Then
            sError = oSocket.GetErrorDescription(oSocket.LastError)
            GoTo QH
        End If
        uCtx.DebugBox.Text = "pvArraySize(baRecv)=" & pvArraySize(baRecv) & vbCrLf & DesignDumpMemory(VarPtr(baRecv(0)), pvArraySize(baRecv))
        If Not pvHandleInput(uCtx, baRecv, sError) Then
            GoTo QH
        End If
        If uCtx.SendPos <> 0 Then
            If Not oSocket.SyncSend(VarPtr(uCtx.SendBuffer(0)), uCtx.SendPos) Then
                sError = oSocket.GetErrorDescription(oSocket.LastError)
                GoTo QH
            End If
            uCtx.SendPos = 0
        End If
    Loop While uCtx.RecvPos <> 0
    uCtx.SendPos = pvSendClientHandshakeFinished(uCtx, uCtx.SendBuffer, uCtx.SendPos, sError)
    If LenB(sError) <> 0 Then
        GoTo QH
    End If
    If uCtx.SendPos > 0 Then
        If Not oSocket.SyncSend(VarPtr(uCtx.SendBuffer(0)), uCtx.SendPos) Then
            sError = oSocket.GetErrorDescription(oSocket.LastError)
            GoTo QH
        End If
        uCtx.SendPos = 0
    End If
    If Not pvDeriveApplicationSecrets(uCtx, sError) Then
        GoTo QH
    End If
    uCtx.State = ucsStatePostHandshake
    sRequest = "GET " & sPath & " HTTP/1.0" & vbCrLf & _
               "Host: " & vSplit(0) & vbCrLf & vbCrLf
    uCtx.SendPos = pvSendClientTrafficData(uCtx, uCtx.SendBuffer, uCtx.SendPos, StrConv(sRequest, vbFromUnicode), sError)
    If LenB(sError) <> 0 Then
        GoTo QH
    End If
    If uCtx.SendPos > 0 Then
        If Not oSocket.SyncSend(VarPtr(uCtx.SendBuffer(0)), uCtx.SendPos) Then
            sError = oSocket.GetErrorDescription(oSocket.LastError)
            GoTo QH
        End If
        uCtx.SendPos = 0
    End If
    Do
        If Not oSocket.SyncReceiveArray(baRecv, Timeout:=-1) Then
            sError = oSocket.GetErrorDescription(oSocket.LastError)
            GoTo QH
        End If
        If Not pvHandleInput(uCtx, baRecv, sError) Then
            GoTo QH
        End If
        If uCtx.SendPos <> 0 Then
            If Not oSocket.SyncSend(VarPtr(uCtx.SendBuffer(0)), uCtx.SendPos) Then
                sError = oSocket.GetErrorDescription(oSocket.LastError)
                GoTo QH
            End If
            uCtx.SendPos = 0
        End If
    Loop While uCtx.RecvPos <> 0
    If uCtx.DecrPos > 0 Then
        TlsFetchHttp = Replace(Replace(StrConv(uCtx.DecrBuffer, vbUnicode), vbCr, vbNullString), vbLf, vbCrLf)
    End If
QH:
End Function

Private Function pvSendClientHello(uCtx As UcsClientContextType, baOutput() As Byte, ByVal lPos As Long) As Long
    Dim cBlocks         As Collection
    
    '--- Record Header
    lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
    lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
    lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks, Size:=2)
        '--- Handshake Header
        lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_CLIENT_HELLO)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks, Size:=3)
            lPos = pvWriteLong(baOutput, lPos, TLS_CLIENT_LEGACY_VERSION, Size:=2)
            lPos = pvWriteArray(baOutput, lPos, uCtx.ClientRandom)
            '--- Legacy Session ID
            lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks)
                lPos = pvWriteArray(baOutput, lPos, uCtx.LegacySessionID)
            lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
            '--- Cipher Suites
            lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks, Size:=2)
'                lPos = pvWriteLong(baOutput, lPos, TLS_CIPHER_SUITE_AES_256_GCM_SHA384, Size:=2)
                lPos = pvWriteLong(baOutput, lPos, TLS_CIPHER_SUITE_CHACHA20_POLY1305_SHA256, Size:=2)
            lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
            '--- Legacy Compression Methods
            lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks)
                lPos = pvWriteLong(baOutput, lPos, TLS_COMPRESS_NULL)
            lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
            '--- Extensions
            lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks, Size:=2)
                If LenB(uCtx.ServerName) <> 0 Then
                    '--- Extension - Server Name
                    lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_SERVER_NAME, Size:=2)
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks, Size:=2)
                            lPos = pvWriteLong(baOutput, lPos, TLS_SERVER_NAME_TYPE_HOSTNAME) '--- FQDN
                            lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks, Size:=2)
                                lPos = pvWriteString(baOutput, lPos, uCtx.ServerName)
                            lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
                        lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
                    lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
                End If
                '--- Extension - Supported Groups
                lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_SUPPORTED_GROUPS, Size:=2)
                lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks, Size:=2)
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks, Size:=2)
                        lPos = pvWriteLong(baOutput, lPos, TLS_GROUP_X25519, Size:=2)
                    lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
                lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
                '--- Extension - Signature Algorithms
                lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_SIGNATURE_ALGORITHMS, Size:=2)
                lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks, Size:=2)
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks, Size:=2)
                        lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, Size:=2)
                        lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_ECDSA_SECP256R1_SHA256, Size:=2)
                        lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PKCS1_SHA256, Size:=2)
                        lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PKCS1_SHA1, Size:=2)
                    lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
                lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
                '--- Extension - Key Share
                lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_KEY_SHARE, Size:=2)
                lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks, Size:=2)
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks, Size:=2)
                        lPos = pvWriteLong(baOutput, lPos, TLS_GROUP_X25519, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks, Size:=2)
                            lPos = pvWriteArray(baOutput, lPos, uCtx.ClientPublic)
                        lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
                    lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
                lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
                '--- Extension - PSK Key Exchange Modes
                lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_PSK_KEY_EXCHANGE_MODES, Size:=2)
                lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks, Size:=2)
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks)
                        lPos = pvWriteLong(baOutput, lPos, TLS_PSK_KE_MODE_PSK_DHE)
                    lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
                lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
                '--- Extension - Supported Versions
                lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_SUPPORTED_VERSIONS, Size:=2)
                lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks, Size:=2)
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks)
                        lPos = pvWriteLong(baOutput, lPos, TLS_PROTOCOL_VERSION_TLS13_FINAL, Size:=2)
                    lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
                lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
            lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
        lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
    lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
    Debug.Assert cBlocks.Count = 0
    pvSendClientHello = lPos
End Function

Private Function pvSendClientHandshakeFinished(uCtx As UcsClientContextType, baOutput() As Byte, ByVal lPos As Long, sError As String) As Long
    Dim cBlocks         As Collection
    Dim lRecordPos      As Long
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim baVerifyData()  As Byte
    Dim baClientIV()    As Byte
    Dim baHandshakeHash() As Byte
    
    '--- Legacy Change Cipher Spec
    baVerifyData = FromHex("140303000101")
    lPos = pvWriteArray(baOutput, lPos, baVerifyData)
    '--- Record Header
    lRecordPos = lPos
    lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_APPDATA)
    lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
    lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks, Size:=2)
        lMessagePos = lPos
        '--- Handshake Finish
        lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_FINISHED)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks, Size:=3)
            baHandshakeHash = pvCryptoHash(uCtx.DigestAlgo, uCtx.HandshakeMessages, 0)
            baVerifyData = pvHkdfExpand(uCtx.DigestAlgo, uCtx.ClientTrafficSecret, "finished", EmptyByteArray, uCtx.DigestSize)
            baVerifyData = pvHkdfExtract(uCtx.DigestAlgo, baVerifyData, baHandshakeHash)
            lPos = pvWriteArray(baOutput, lPos, baVerifyData)
        lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
        lMessageSize = lPos - lMessagePos
        lPos = pvWriteReserved(baOutput, lPos, uCtx.TagSize)
    lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
    Debug.Assert cBlocks.Count = 0
    baClientIV = pvArrayXor(uCtx.ClientTrafficIV, uCtx.ClientTrafficSeqNo)
    If pvCryptoEncrypt(uCtx.AeadAlgo, baClientIV, uCtx.ClientTrafficKey, baOutput, lRecordPos, 5, baOutput, lMessagePos, lMessageSize) Then
        uCtx.ClientTrafficSeqNo = uCtx.ClientTrafficSeqNo + 1
    Else
        sError = "Encryption failed"
        GoTo QH
    End If
    pvSendClientHandshakeFinished = lPos
QH:
End Function

Private Function pvSendClientTrafficData(uCtx As UcsClientContextType, baOutput() As Byte, ByVal lPos As Long, baData() As Byte, sError As String) As Long
    Dim cBlocks         As Collection
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim baClientIV()    As Byte
    
    '--- Record Header
    lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_APPDATA)
    lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
    lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks, Size:=2)
        lMessagePos = lPos
        lPos = pvWriteArray(baOutput, lPos, baData)
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_APPDATA)
        lMessageSize = lPos - lMessagePos
        lPos = pvWriteReserved(baOutput, lPos, uCtx.TagSize)
    lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
    Debug.Assert cBlocks.Count = 0
    baClientIV = pvArrayXor(uCtx.ClientTrafficIV, uCtx.ClientTrafficSeqNo)
    If pvCryptoEncrypt(uCtx.AeadAlgo, baClientIV, uCtx.ClientTrafficKey, baOutput, 0, 5, baOutput, lMessagePos, lMessageSize) Then
        uCtx.ClientTrafficSeqNo = uCtx.ClientTrafficSeqNo + 1
    Else
        sError = "Encryption failed"
        GoTo QH
    End If
    pvSendClientTrafficData = lPos
QH:
End Function

Private Function pvHandleInput(uCtx As UcsClientContextType, baInput() As Byte, sError As String) As Boolean
    Dim lPos            As Long
    Dim lSize           As Long
    
    If uCtx.RecvPos <> 0 Then
        lPos = lPos
    End If
    uCtx.RecvPos = pvWriteArray(uCtx.RecvBuffer, uCtx.RecvPos, baInput)
    lPos = pvHandleRecord(uCtx, uCtx.RecvBuffer, uCtx.RecvPos, sError)
    If LenB(sError) <> 0 Then
        GoTo QH
    End If
    lSize = uCtx.RecvPos - lPos
    If lPos > 0 And lSize > 0 Then
        Call CopyMemory(uCtx.RecvBuffer(0), uCtx.RecvBuffer(lPos), lSize)
    End If
    uCtx.RecvPos = IIf(lSize > 0, lSize, 0)
    '--- success
    pvHandleInput = True
QH:
End Function

Private Function pvHandleRecord(uCtx As UcsClientContextType, baInput() As Byte, ByVal lSize As Long, sError As String) As Long
    Dim lRecordPos      As Long
    Dim lRecordSize     As Long
    Dim lPos            As Long
    Dim cBlocks         As Collection
    Dim lRecordType     As Long
    Dim lLegacyProtocol As Long
    Dim lHandshakeType  As Long
    Dim baServerIV()    As Byte
    Dim lEnd            As Long
    Dim baVerifyData()  As Byte
    Dim baMessage()     As Byte
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim lVerifyPos      As Long
    Dim baHandshakeHash() As Byte
    Dim lRequestUpdate  As Long
        
    Do While lPos + 6 <= lSize
        lRecordPos = lPos
        lPos = pvReadLong(baInput, lPos, lRecordType)
        lPos = pvReadLong(baInput, lPos, lLegacyProtocol, Size:=2)
        lPos = pvReadBeginOfBlock(baInput, lPos, cBlocks, Size:=2, BlockSize:=lRecordSize)
        If lRecordSize > IIf(lRecordType = TLS_CONTENT_TYPE_APPDATA, TLS_MAX_ENCRYPTED_RECORD_SIZE, TLS_MAX_PLAINTEXT_RECORD_SIZE) Then
            sError = "Record size too big"
            GoTo QH
        End If
        If lPos + lRecordSize > lSize Then
            lPos = lRecordPos
            Exit Do
        End If
            Select Case lRecordType
            Case TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC
                lPos = lPos + lRecordSize
            Case TLS_CONTENT_TYPE_HANDSHAKE
                lMessagePos = lPos
                lPos = pvReadLong(baInput, lPos, lHandshakeType)
                lPos = pvReadBeginOfBlock(baInput, lPos, cBlocks, Size:=3, BlockSize:=lMessageSize)
                    Select Case uCtx.State
                    Case ucsStateExpectServerHello
                        Select Case lHandshakeType
                        Case TLS_HANDSHAKE_TYPE_SERVER_HELLO
                            lPos = pvReadArray(baInput, lPos, baMessage, lMessageSize)
                            If Not pvHandleServerHello(uCtx, baMessage, sError) Then
                                GoTo QH
                            End If
                            pvWriteBuffer uCtx.HandshakeMessages, pvArraySize(uCtx.HandshakeMessages), VarPtr(baInput(lMessagePos)), lMessageSize + 4
                            If Not pvDeriveHandshakeSecrets(uCtx, sError) Then
                                GoTo QH
                            End If
                            uCtx.State = ucsStateExpectEncryptedExtensions
                        Case Else
                            sError = "Unexpected message type for ucsStateExpectServerHello (lHandshakeType=" & lHandshakeType & ")"
                            GoTo QH
                        End Select
                    End Select
                lPos = pvReadEndOfBlock(baInput, lPos, cBlocks)
            Case TLS_CONTENT_TYPE_APPDATA
                baServerIV = pvArrayXor(uCtx.ServerTrafficIV, uCtx.ServerTrafficSeqNo)
                If pvCryptoDecrypt(uCtx.AeadAlgo, baServerIV, uCtx.ServerTrafficKey, baInput, lRecordPos, 5, baInput, lPos, lRecordSize) Then
                    uCtx.ServerTrafficSeqNo = uCtx.ServerTrafficSeqNo + 1
                ElseIf pvArraySize(uCtx.PrevServerTrafficIV) <> 0 Then
                    baServerIV = pvArrayXor(uCtx.PrevServerTrafficIV, uCtx.PrevServerTrafficSeqNo)
                    If pvCryptoDecrypt(uCtx.AeadAlgo, baServerIV, uCtx.PrevServerTrafficKey, baInput, lRecordPos, 5, baInput, lPos, lRecordSize) Then
                        uCtx.PrevServerTrafficSeqNo = uCtx.PrevServerTrafficSeqNo + 1
                    Else
                        sError = "pvCryptoDecrypt w/ PrevServerTrafficIV failed"
                        GoTo QH
                    End If
                Else
                    sError = "pvCryptoDecrypt w/ ServerTrafficIV failed"
                    GoTo QH
                End If
                lEnd = lPos + lRecordSize - uCtx.TagSize - 1
                '--- trim zero padding at the end of decrypted record
                Do While baInput(lEnd) = 0
                    lEnd = lEnd - 1
                Loop
                lRecordType = baInput(lEnd)
                Select Case lRecordType
                Case TLS_CONTENT_TYPE_HANDSHAKE
                    Select Case uCtx.State
                    Case ucsStateExpectEncryptedExtensions
                        Do While lPos < lEnd
                            lMessagePos = lPos
                            lPos = pvReadLong(baInput, lPos, lHandshakeType)
                            lPos = pvReadBeginOfBlock(baInput, lPos, cBlocks, Size:=3, BlockSize:=lMessageSize)
                            If lMessageSize > 16 * 1024 Then
                                sError = "Unexpected message size (lMessageSize=" & lMessageSize & ")"
                                GoTo QH
                            End If
                            lPos = pvReadArray(baInput, lPos, baMessage, lMessageSize)
                            Select Case lHandshakeType
                            Case TLS_HANDSHAKE_TYPE_CERTIFICATE
                                uCtx.ServerCertificate = baMessage
                            Case TLS_HANDSHAKE_TYPE_CERTIFICATE_VERIFY
                                baHandshakeHash = pvCryptoHash(uCtx.DigestAlgo, uCtx.HandshakeMessages, 0)
                                lVerifyPos = pvWriteString(baVerifyData, 0, Space$(64) & "TLS 1.3, server CertificateVerify" & Chr$(0))
                                lVerifyPos = pvWriteArray(baVerifyData, lVerifyPos, baHandshakeHash)
                                '--- ToDo: verify uCtx.ServerCertificate signature
                                '--- ShellExecute("openssl x509 -pubkey -noout -in server.crt > server.pub")
                            Case TLS_HANDSHAKE_TYPE_FINISHED
                                baHandshakeHash = pvCryptoHash(uCtx.DigestAlgo, uCtx.HandshakeMessages, 0)
                                baVerifyData = pvHkdfExpand(uCtx.DigestAlgo, uCtx.ServerTrafficSecret, "finished", EmptyByteArray, uCtx.DigestSize)
                                baVerifyData = pvHkdfExtract(uCtx.DigestAlgo, baVerifyData, baHandshakeHash)
                                Debug.Assert StrConv(baVerifyData, vbUnicode) = StrConv(baMessage, vbUnicode)
                                If StrConv(baVerifyData, vbUnicode) <> StrConv(baMessage, vbUnicode) Then
                                    sError = "Server Handshake verification failed"
                                    GoTo QH
                                End If
                            Case TLS_HANDSHAKE_TYPE_ENCRYPTED_EXTENSIONS, TLS_HANDSHAKE_TYPE_COMPRESSED_CERTIFICATE
                                '--- do nothing
                            Case Else
                                sError = "Unexpected message type for ucsStateExpectEncryptedExtensions (lHandshakeType=" & lHandshakeType & ")"
                                GoTo QH
                            End Select
                            pvWriteBuffer uCtx.HandshakeMessages, pvArraySize(uCtx.HandshakeMessages), VarPtr(baInput(lMessagePos)), lMessageSize + 4
                            lPos = pvReadEndOfBlock(baInput, lPos, cBlocks)
                        Loop
                        '--- note: skip padding too
                        lPos = lRecordPos + lRecordSize + 5
                    Case ucsStatePostHandshake
                        Do While lPos < lEnd
                            lMessagePos = lPos
                            lPos = pvReadLong(baInput, lPos, lHandshakeType)
                            lPos = pvReadBeginOfBlock(baInput, lPos, cBlocks, Size:=3, BlockSize:=lMessageSize)
                            lPos = pvReadArray(baInput, lPos, baMessage, lMessageSize)
                            Select Case lHandshakeType
                            Case TLS_HANDSHAKE_TYPE_NEW_SESSION_TICKET
                                '--- do nothing
                            Case TLS_HANDSHAKE_TYPE_KEY_UPDATE
                                Debug.Print "TLS_HANDSHAKE_TYPE_KEY_UPDATE"
                                If pvArraySize(baMessage) = 1 Then
                                    lRequestUpdate = baMessage(0)
                                Else
                                    lRequestUpdate = -1
                                End If
                                Select Case lRequestUpdate
                                Case 0, 1
                                    If Not pvDeriveKeyUpdate(uCtx, lRequestUpdate <> 0, sError) Then
                                        GoTo QH
                                    End If
                                    If lRequestUpdate = 1 Then
                                        '--- ack by TLS_HANDSHAKE_TYPE_KEY_UPDATE w/ update_not_requested(0)
                                        If pvSendClientTrafficData(uCtx, baMessage, 0, FromHex("1800000100"), sError) = 0 Then
                                            GoTo QH
                                        End If
                                        uCtx.SendPos = pvWriteArray(uCtx.SendBuffer, uCtx.SendPos, baMessage)
                                    End If
                                Case Else
                                    sError = "Unexpected value in TLS_HANDSHAKE_TYPE_KEY_UPDATE (lRequestUpdate=" & lRequestUpdate & ")"
                                    GoTo QH
                                End Select
                            Case Else
                                sError = "Unexpected message type for ucsStatePostHandshake (lHandshakeType=" & lHandshakeType & ")"
                                GoTo QH
                            End Select
                            lPos = pvReadEndOfBlock(baInput, lPos, cBlocks)
                        Loop
                        '--- note: skip padding too
                        lPos = lRecordPos + lRecordSize + 5
                    Case Else
                        sError = "Invalid state for TLS_CONTENT_TYPE_HANDSHAKE (" & uCtx.State & ")"
                        GoTo QH
                    End Select
                Case TLS_CONTENT_TYPE_ALERT
                    If Not uCtx.DebugBox Is Nothing Then
                        uCtx.DebugBox.Text = uCtx.DebugBox.Text & vbCrLf & _
                            Switch(baInput(lPos) = TLS_ALERT_LEVEL_WARNING, "Warning alert: ", _
                                   baInput(lPos) = TLS_ALERT_LEVEL_FATAL, "Fatal alert: ", _
                                   True, "Unknown alert: ") & baInput(lPos + 1)
                    End If
                    lPos = lPos + lRecordSize
                Case TLS_CONTENT_TYPE_APPDATA
                    Select Case uCtx.State
                    Case ucsStatePostHandshake
                        uCtx.DecrPos = pvWriteBuffer(uCtx.DecrBuffer, uCtx.DecrPos, VarPtr(baInput(lPos)), lEnd - lPos)
                    Case Else
                        sError = "Invalid state for TLS_CONTENT_TYPE_APPDATA (" & uCtx.State & ")"
                        GoTo QH
                    End Select
                    lPos = lPos + lRecordSize
                Case Else
                    lPos = lPos + lRecordSize
                End Select
            Case TLS_CONTENT_TYPE_ALERT
                If Not uCtx.DebugBox Is Nothing Then
                    uCtx.DebugBox.Text = uCtx.DebugBox.Text & vbCrLf & _
                        Switch(baInput(lPos) = 1, "Warning alert: ", baInput(lPos) = 2, "Fatal alert: ", True, "Unknown alert: ") & baInput(lPos + 1)
                End If
                If baInput(lPos) = 2 Then
                    GoTo QH
                End If
                lPos = lPos + lRecordSize
            Case Else
                sError = "Unexpected record type (" & lRecordType & ")"
                GoTo QH
            End Select
        lPos = pvReadEndOfBlock(baInput, lPos, cBlocks)
    Loop
    '--- success
    pvHandleRecord = lPos
QH:
End Function

Private Function pvHandleServerHello(uCtx As UcsClientContextType, baInput() As Byte, sError As String) As Boolean
    Dim lPos            As Long
    Dim cBlocks         As Collection
    Dim lLegacyVersion  As Long
    Dim lBlockSize      As Long
    Dim lLegacyCompression As Long
    Dim lExtType        As Long
    Dim lExchangeGroup  As Long
    
    lPos = pvReadLong(baInput, lPos, lLegacyVersion, Size:=2)
    lPos = pvReadArray(baInput, lPos, uCtx.ServerRandom, uCtx.SecretSize)
    lPos = pvReadBeginOfBlock(baInput, lPos, cBlocks, BlockSize:=lBlockSize)
        lPos = pvReadArray(baInput, lPos, uCtx.LegacySessionID, lBlockSize)
    lPos = pvReadEndOfBlock(baInput, lPos, cBlocks)
    lPos = pvReadLong(baInput, lPos, uCtx.CipherSuite, Size:=2)
    Select Case uCtx.CipherSuite
    Case TLS_CIPHER_SUITE_CHACHA20_POLY1305_SHA256
        uCtx.AeadAlgo = ucsAlgoAeadChacha20Poly1305
        uCtx.KeySize = TLS_CHACHA20_KEY_SIZE
        uCtx.IvSize = TLS_CHACHA20POLY1305_IV_SIZE
        uCtx.TagSize = TLS_CHACHA20POLY1305_TAG_SIZE
        uCtx.DigestAlgo = ucsAlgoDigestSha256
        uCtx.DigestSize = TLS_SHA256_DIGEST_SIZE
    Case TLS_CIPHER_SUITE_AES_256_GCM_SHA384
        uCtx.AeadAlgo = ucsAlgoAeadChacha20Poly1305
        uCtx.KeySize = TLS_AES256_KEY_SIZE
        uCtx.IvSize = TLS_AESGCM_IV_SIZE
        uCtx.TagSize = TLS_AESGCM_TAG_SIZE
        uCtx.DigestAlgo = ucsAlgoDigestSha384
        uCtx.DigestSize = TLS_SHA384_DIGEST_SIZE
    Case Else
        sError = "Unsupported cipher suite (0x" & Hex$(uCtx.CipherSuite) & ")"
        GoTo QH
    End Select
    lPos = pvReadLong(baInput, lPos, lLegacyCompression)
    Debug.Assert lLegacyCompression = 0
    lPos = pvReadBeginOfBlock(baInput, lPos, cBlocks, Size:=2)
        Do While lPos < cBlocks.Item(cBlocks.Count)
            lPos = pvReadLong(baInput, lPos, lExtType, Size:=2)
            lPos = pvReadBeginOfBlock(baInput, lPos, cBlocks, Size:=2, BlockSize:=lBlockSize)
                Select Case lExtType
                Case TLS_EXTENSION_TYPE_KEY_SHARE
                    lPos = pvReadLong(baInput, lPos, lExchangeGroup, Size:=2)
                    Debug.Assert lExchangeGroup = TLS_GROUP_X25519
                    lPos = pvReadBeginOfBlock(baInput, lPos, cBlocks, Size:=2, BlockSize:=lBlockSize)
                        Debug.Assert lBlockSize = uCtx.KeySize
                        If lBlockSize <> uCtx.KeySize Then
                            sError = "Invalid server key size"
                            GoTo QH
                        End If
                        lPos = pvReadArray(baInput, lPos, uCtx.ServerPublic, uCtx.KeySize)
                    lPos = pvReadEndOfBlock(baInput, lPos, cBlocks)
                Case TLS_EXTENSION_TYPE_SUPPORTED_VERSIONS
                    If lBlockSize >= 2 Then
                        Call pvReadLong(baInput, lPos, uCtx.ServerSupportedVersion, Size:=2)
                    End If
                    lPos = lPos + lBlockSize
                Case Else
                    lPos = lPos + lBlockSize
                End Select
            lPos = pvReadEndOfBlock(baInput, lPos, cBlocks)
        Loop
    lPos = pvReadEndOfBlock(baInput, lPos, cBlocks)
    '--- success
    pvHandleServerHello = True
QH:
End Function

'= HMAC-based key derivation functions ===================================

Private Function pvDeriveHandshakeSecrets(uCtx As UcsClientContextType, sError As String) As Boolean
    Dim baHandshakeHash() As Byte
    Dim baEarlySecret() As Byte
    Dim baEmptyHash()   As Byte
    Dim baDerivedSecret() As Byte
    Dim baSharedSecret() As Byte
    
    With uCtx
        If pvArraySize(.HandshakeMessages) = 0 Then
            sError = "Missing handshake records"
            GoTo QH
        End If
        baHandshakeHash = pvCryptoHash(.DigestAlgo, .HandshakeMessages, 0)
        '--- for ucsAlgoDigestSha256 always 33AD0A1C607EC03B09E6CD9893680CE210ADF300AA1F2660E1B22E10F170F92A
        baEarlySecret = pvHkdfExtract(.DigestAlgo, EmptyByteArray(.DigestSize), EmptyByteArray(.KeySize))
        '--- for ucsAlgoDigestSha256 always  E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855
        baEmptyHash = pvCryptoHash(.DigestAlgo, EmptyByteArray, 0)
        '--- for ucsAlgoDigestSha256 always 6F2615A108C702C5678F54FC9DBAB69716C076189C48250CEBEAC3576C3611BA
        baDerivedSecret = pvHkdfExpand(.DigestAlgo, baEarlySecret, "derived", baEmptyHash, .DigestSize)
        baSharedSecret = pvCryptoDeriveSecret(.KxAlgo, .ClientPrivate, .ServerPublic)
        .HandshakeSecret = pvHkdfExtract(.DigestAlgo, baDerivedSecret, baSharedSecret)
        
        .ServerTrafficSecret = pvHkdfExpand(.DigestAlgo, .HandshakeSecret, "s hs traffic", baHandshakeHash, .DigestSize)
        .ServerTrafficKey = pvHkdfExpand(.DigestAlgo, .ServerTrafficSecret, "key", EmptyByteArray, .KeySize)
        .ServerTrafficIV = pvHkdfExpand(.DigestAlgo, .ServerTrafficSecret, "iv", EmptyByteArray, .IvSize)
        .ServerTrafficSeqNo = 0
        .ClientTrafficSecret = pvHkdfExpand(.DigestAlgo, .HandshakeSecret, "c hs traffic", baHandshakeHash, .DigestSize)
        .ClientTrafficKey = pvHkdfExpand(.DigestAlgo, .ClientTrafficSecret, "key", EmptyByteArray, .KeySize)
        .ClientTrafficIV = pvHkdfExpand(.DigestAlgo, .ClientTrafficSecret, "iv", EmptyByteArray, .IvSize)
        .ClientTrafficSeqNo = 0
    End With
    '--- success
    pvDeriveHandshakeSecrets = True
QH:
End Function

Private Function pvDeriveApplicationSecrets(uCtx As UcsClientContextType, sError As String) As Boolean
    Dim baHandshakeHash() As Byte
    Dim baEmptyHash()   As Byte
    Dim baDerivedSecret() As Byte
    
    With uCtx
        If pvArraySize(.HandshakeMessages) = 0 Then
            sError = "Missing handshake records"
            GoTo QH
        End If
        baHandshakeHash = pvCryptoHash(.DigestAlgo, .HandshakeMessages, 0)
        '--- for ucsAlgoDigestSha256 always E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855
        baEmptyHash = pvCryptoHash(.DigestAlgo, EmptyByteArray, 0)
        '--- for ucsAlgoDigestSha256 always 6F2615A108C702C5678F54FC9DBAB69716C076189C48250CEBEAC3576C3611BA
        baDerivedSecret = pvHkdfExpand(.DigestAlgo, .HandshakeSecret, "derived", baEmptyHash, .DigestSize)
        .MasterSecret = pvHkdfExtract(.DigestAlgo, baDerivedSecret, EmptyByteArray(.KeySize))
        
        .ServerTrafficSecret = pvHkdfExpand(.DigestAlgo, .MasterSecret, "s ap traffic", baHandshakeHash, .DigestSize)
        .ServerTrafficKey = pvHkdfExpand(.DigestAlgo, .ServerTrafficSecret, "key", EmptyByteArray, .KeySize)
        .ServerTrafficIV = pvHkdfExpand(.DigestAlgo, .ServerTrafficSecret, "iv", EmptyByteArray, .IvSize)
        .ServerTrafficSeqNo = 0
        .ClientTrafficSecret = pvHkdfExpand(.DigestAlgo, .MasterSecret, "c ap traffic", baHandshakeHash, .DigestSize)
        .ClientTrafficKey = pvHkdfExpand(.DigestAlgo, .ClientTrafficSecret, "key", EmptyByteArray, .KeySize)
        .ClientTrafficIV = pvHkdfExpand(.DigestAlgo, .ClientTrafficSecret, "iv", EmptyByteArray, .IvSize)
        .ClientTrafficSeqNo = 0
    End With
    '--- success
    pvDeriveApplicationSecrets = True
QH:
End Function

Private Function pvDeriveKeyUpdate(uCtx As UcsClientContextType, ByVal bUpdateClient As Boolean, sError As String) As Boolean
    With uCtx
        If pvArraySize(.ServerTrafficSecret) = 0 Then
            sError = "Missing previous server secret"
            GoTo QH
        End If
        .PrevServerTrafficKey = .ServerTrafficSecret
        .PrevServerTrafficIV = .ServerTrafficIV
        .PrevServerTrafficSeqNo = .ServerTrafficSeqNo
        .ServerTrafficSecret = pvHkdfExpand(.DigestAlgo, .ServerTrafficSecret, "traffic upd", EmptyByteArray, .DigestSize)
        .ServerTrafficKey = pvHkdfExpand(.DigestAlgo, .ServerTrafficSecret, "key", EmptyByteArray, .KeySize)
        .ServerTrafficIV = pvHkdfExpand(.DigestAlgo, .ServerTrafficSecret, "iv", EmptyByteArray, .IvSize)
        .ServerTrafficSeqNo = 0
        If bUpdateClient Then
            If pvArraySize(.ClientTrafficSecret) = 0 Then
                sError = "Missing previous client secret"
                GoTo QH
            End If
            .ClientTrafficSecret = pvHkdfExpand(.DigestAlgo, .ClientTrafficSecret, "traffic upd", EmptyByteArray, .DigestSize)
            .ClientTrafficKey = pvHkdfExpand(.DigestAlgo, .ClientTrafficSecret, "key", EmptyByteArray, .KeySize)
            .ClientTrafficIV = pvHkdfExpand(.DigestAlgo, .ClientTrafficSecret, "iv", EmptyByteArray, .IvSize)
            .ClientTrafficSeqNo = 0
        End If
    End With
    '--- success
    pvDeriveKeyUpdate = True
QH:
End Function

Private Function pvHkdfExtract(ByVal eHash As UcsCryptoAlgorithmsEnum, baSalt() As Byte, baInput() As Byte) As Byte()
    pvHkdfExtract = pvCryptoHmac(eHash, baSalt, baInput, 0)
End Function

Private Function pvHkdfExpand(ByVal eHash As UcsCryptoAlgorithmsEnum, baSalt() As Byte, ByVal sLabel As String, baContext() As Byte, ByVal lSize As Long) As Byte()
    Dim baRetVal()      As Byte
    Dim lRetValPos      As Long
    Dim baInfo()        As Byte
    Dim lInfoPos        As Long
    Dim baInput()       As Byte
    Dim lInputPos       As Long
    Dim lIdx            As Long
    Dim baLast()        As Byte
    
    If LenB(sLabel) <> 0 Then
        sLabel = "tls13 " & sLabel
        pvWriteReserved baInfo, 0, 3 + Len(sLabel) + 1 + pvArraySize(baContext)
        lInfoPos = pvWriteLong(baInfo, lInfoPos, lSize, Size:=2)
        lInfoPos = pvWriteLong(baInfo, lInfoPos, Len(sLabel))
        lInfoPos = pvWriteString(baInfo, lInfoPos, sLabel)
        lInfoPos = pvWriteLong(baInfo, lInfoPos, pvArraySize(baContext))
        lInfoPos = pvWriteArray(baInfo, lInfoPos, baContext)
    Else
        baInfo = baContext
    End If
    lIdx = 1
    Do While lRetValPos < lSize
        lInputPos = pvWriteArray(baInput, 0, baLast)
        lInputPos = pvWriteArray(baInput, lInputPos, baInfo)
        lInputPos = pvWriteLong(baInput, lInputPos, lIdx)
        baLast = pvCryptoHmac(eHash, baSalt, baInput, 0, Size:=lInputPos)
        lRetValPos = pvWriteArray(baRetVal, lRetValPos, baLast)
        lIdx = lIdx + 1
    Loop
    If UBound(baRetVal) <> lSize - 1 Then
        ReDim Preserve baRetVal(0 To lSize - 1) As Byte
    End If
    pvHkdfExpand = baRetVal
    Debug.Print "sLabel=" & sLabel & ", pvHkdfExpand=0x" & ToHex(baRetVal)
End Function

'= crypto wrappers =======================================================

Private Function pvCryptoDecrypt(eAead As UcsCryptoAlgorithmsEnum, baServerIV() As Byte, baServerKey() As Byte, baAd() As Byte, ByVal lAdPos As Long, ByVal lAdSize As Long, baInput() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Dim baBuffer()      As Byte
    Dim lResult         As Long
    Dim aSize(0 To 1)   As Long
    
    If eAead = ucsAlgoAeadChacha20Poly1305 Then
        Debug.Assert pvArraySize(baServerIV) = TLS_CHACHA20POLY1305_IV_SIZE
        Debug.Assert pvArraySize(baServerKey) = TLS_CHACHA20_KEY_SIZE
        Debug.Assert pvArraySize(baInput) >= lPos + lSize - TLS_CHACHA20POLY1305_TAG_SIZE
        ReDim baBuffer(0 To lSize - TLS_CHACHA20POLY1305_TAG_SIZE - 1 + 1000) As Byte
        lResult = crypto_aead_chacha20poly1305_ietf_decrypt(baBuffer(0), aSize(0), 0, baInput(lPos), lSize, 0, baAd(lAdPos), lAdSize, 0, baServerIV(0), baServerKey(0))
        If lResult <> 0 Then
            GoTo QH
        End If
        Call CopyMemory(baInput(lPos), baBuffer(0), lSize - TLS_CHACHA20POLY1305_TAG_SIZE)
    Else
        Err.Raise vbObjectError, "Unsupported aead type"
    End If
    '--- success
    pvCryptoDecrypt = True
QH:
End Function

Private Function pvCryptoEncrypt(eAead As UcsCryptoAlgorithmsEnum, baClientIV() As Byte, baClientKey() As Byte, baAd() As Byte, ByVal lAdPos As Long, ByVal lAdSize As Long, baInput() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Dim baBuffer()      As Byte
    Dim lResult         As Long
    Dim aSize(0 To 1)   As Long
    
    If eAead = ucsAlgoAeadChacha20Poly1305 Then
        Debug.Assert pvArraySize(baClientIV) = TLS_CHACHA20POLY1305_IV_SIZE
        Debug.Assert pvArraySize(baClientKey) = TLS_CHACHA20_KEY_SIZE
        Debug.Assert pvArraySize(baInput) >= lPos + lSize + TLS_CHACHA20POLY1305_TAG_SIZE
        ReDim baBuffer(0 To lSize + TLS_CHACHA20POLY1305_TAG_SIZE - 1 + 1000) As Byte
        lResult = crypto_aead_chacha20poly1305_ietf_encrypt(baBuffer(0), aSize(0), baInput(lPos), lSize, 0, baAd(lAdPos), lAdSize, 0, 0, baClientIV(0), baClientKey(0))
        Debug.Assert lResult = 0
        If lResult <> 0 Then
            GoTo QH
        End If
        Call CopyMemory(baInput(lPos), baBuffer(0), lSize + TLS_CHACHA20POLY1305_TAG_SIZE)
    Else
        Err.Raise vbObjectError, "Unsupported aead type"
    End If
    '--- success
    pvCryptoEncrypt = True
QH:
End Function

Private Function pvCryptoHash(eHash As UcsCryptoAlgorithmsEnum, baInput() As Byte, ByVal lPos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim baRetVal()      As Byte
    Dim lPtr            As Long
    
    Select Case eHash
    Case ucsAlgoDigestSha256
        If Size < 0 Then
            Size = pvArraySize(baInput) - lPos
        Else
            Debug.Assert pvArraySize(baInput) >= lPos + Size
        End If
        If Size > 0 Then
            lPtr = VarPtr(baInput(lPos))
        End If
        ReDim baRetVal(0 To TLS_SHA256_DIGEST_SIZE - 1) As Byte
        Call crypto_hash_sha256(baRetVal(0), ByVal lPtr, Size)
    Case Else
        Err.Raise vbObjectError, "Unsupported hash type"
    End Select
    pvCryptoHash = baRetVal
End Function

Private Function pvCryptoHmac(ByVal eHash As UcsCryptoAlgorithmsEnum, baSalt() As Byte, baInput() As Byte, ByVal lPos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim baRetVal()      As Byte
    Dim lPtr            As Long
    
    Select Case eHash
    Case ucsAlgoDigestSha256
        Debug.Assert pvArraySize(baSalt) = TLS_SHA256_DIGEST_SIZE
        If Size < 0 Then
            Size = pvArraySize(baInput) - lPos
        Else
            Debug.Assert pvArraySize(baInput) >= lPos + Size
        End If
        If Size > 0 Then
            lPtr = VarPtr(baInput(lPos))
        End If
        ReDim baRetVal(0 To TLS_SHA256_DIGEST_SIZE - 1) As Byte
        Call crypto_auth_hmacsha256(baRetVal(0), ByVal lPtr, Size, 0, baSalt(0))
    Case Else
        Err.Raise vbObjectError, "Unsupported hash type"
    End Select
    pvCryptoHmac = baRetVal
End Function

Private Function pvCryptoDeriveSecret(ByVal eKeyX As UcsCryptoAlgorithmsEnum, baPriv() As Byte, baPub() As Byte) As Byte()
    Dim baRetVal()      As Byte
    
    If eKeyX = ucsAlgoKeyX25519 Then
        Debug.Assert pvArraySize(baPriv) = TLS_X25519_KEY_SIZE
        Debug.Assert pvArraySize(baPub) = TLS_X25519_KEY_SIZE
        ReDim baRetVal(0 To TLS_X25519_KEY_SIZE - 1) As Byte
        Call crypto_scalarmult_curve25519(baRetVal(0), baPriv(0), baPub(0))
    Else
        Err.Raise vbObjectError, "Unsupported key-exchange type"
    End If
    pvCryptoDeriveSecret = baRetVal
End Function

Private Function pvCryptoRandomBytes(ByVal lSize As Long) As Byte()
    Dim baRetVal()      As Byte
    
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
        Call randombytes_buf(baRetVal(0), lSize)
    End If
    pvCryptoRandomBytes = baRetVal
End Function

'= buffer management =====================================================

Private Function pvWriteBeginOfBlock(baBuffer() As Byte, ByVal lPos As Long, cBlocks As Collection, Optional ByVal Size As Long = 1) As Long
    If cBlocks Is Nothing Then
        Set cBlocks = New Collection
    End If
    cBlocks.Add lPos
    pvWriteBeginOfBlock = pvWriteReserved(baBuffer, lPos, Size)
    '--- note: keep Size in baBuffer
    baBuffer(lPos) = (Size And &HFF)
End Function

Private Function pvWriteEndOfBlock(baBuffer() As Byte, ByVal lPos As Long, cBlocks As Collection) As Long
    Dim lStart          As Long
    
    lStart = cBlocks.Item(cBlocks.Count)
    cBlocks.Remove cBlocks.Count
    pvWriteLong baBuffer, lStart, lPos - lStart - baBuffer(lStart), Size:=baBuffer(lStart)
    pvWriteEndOfBlock = lPos
End Function

Private Function pvWriteString(baBuffer() As Byte, ByVal lPos As Long, sValue As String) As Long
    pvWriteString = pvWriteArray(baBuffer, lPos, StrConv(sValue, vbFromUnicode))
End Function

Private Function pvWriteArray(baBuffer() As Byte, ByVal lPos As Long, baSrc() As Byte) As Long
    Dim lSize       As Long
    
    If pvArraySize(baSrc, RetVal:=lSize) > 0 Then
        lPos = pvWriteBuffer(baBuffer, lPos, VarPtr(baSrc(0)), lSize)
    End If
    pvWriteArray = lPos
End Function

Private Function pvWriteLong(baBuffer() As Byte, ByVal lPos As Long, ByVal lValue As Long, Optional ByVal Size As Long = 1) As Long
    Static baTemp(0 To 3) As Byte

    If Size <= 1 Then
        pvWriteLong = pvWriteBuffer(baBuffer, lPos, VarPtr(lValue), Size)
    Else
        baTemp(Size - 1) = (lValue And &HFF): lValue = lValue \ &H100
        baTemp(Size - 2) = (lValue And &HFF): lValue = lValue \ &H100
        If Size >= 3 Then
            baTemp(Size - 3) = (lValue And &HFF): lValue = lValue \ &H100
            If Size >= 4 Then
                baTemp(Size - 4) = (lValue And &HFF)
            End If
        End If
        pvWriteLong = pvWriteBuffer(baBuffer, lPos, VarPtr(baTemp(0)), Size)
    End If
End Function

Private Function pvWriteReserved(baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Long
    pvWriteReserved = pvWriteBuffer(baBuffer, lPos, 0, lSize)
End Function

Private Function pvWriteBuffer(baBuffer() As Byte, ByVal lPos As Long, ByVal lPtr As Long, ByVal lSize As Long) As Long
    If Peek(ArrPtr(baBuffer)) = 0 Then
        ReDim baBuffer(0 To lPos + lSize - 1) As Byte
    ElseIf UBound(baBuffer) < lPos + lSize - 1 Then
        ReDim Preserve baBuffer(0 To lPos + lSize - 1) As Byte
    End If
    If lSize > 0 And lPtr <> 0 Then
        Debug.Assert IsBadReadPtr(lPtr, lSize) = 0
        Call CopyMemory(baBuffer(lPos), ByVal lPtr, lSize)
    End If
    pvWriteBuffer = lPos + lSize
End Function

Private Function pvReadBeginOfBlock(baBuffer() As Byte, ByVal lPos As Long, cBlocks As Collection, Optional ByVal Size As Long = 1, Optional BlockSize As Long) As Long
    If cBlocks Is Nothing Then
        Set cBlocks = New Collection
    End If
    pvReadBeginOfBlock = pvReadLong(baBuffer, lPos, BlockSize, Size)
    cBlocks.Add pvReadBeginOfBlock + BlockSize
End Function

Private Function pvReadEndOfBlock(baBuffer() As Byte, ByVal lPos As Long, cBlocks As Collection) As Long
    Dim lEnd          As Long
    
    #If baBuffer Then '--- touch args
    #End If
    lEnd = cBlocks.Item(cBlocks.Count)
    cBlocks.Remove cBlocks.Count
    Debug.Assert lPos = lEnd
    pvReadEndOfBlock = lEnd
End Function

Private Function pvReadLong(baBuffer() As Byte, ByVal lPos As Long, lValue As Long, Optional ByVal Size As Long = 1) As Long
    Static baTemp(0 To 3) As Byte
    
    If lPos + Size <= pvArraySize(baBuffer) Then
        If Size <= 1 Then
            lValue = baBuffer(lPos)
        Else
            baTemp(Size - 1) = baBuffer(lPos + 0)
            baTemp(Size - 2) = baBuffer(lPos + 1)
            If Size >= 3 Then baTemp(Size - 3) = baBuffer(lPos + 2)
            If Size >= 4 Then baTemp(Size - 4) = baBuffer(lPos + 3)
            Call CopyMemory(lValue, baTemp(0), Size)
        End If
    Else
        lValue = 0
    End If
    pvReadLong = lPos + Size
End Function

Private Function pvReadArray(baBuffer() As Byte, ByVal lPos As Long, baDest() As Byte, ByVal lSize As Long) As Long
    If lSize < 0 Then
        lSize = pvArraySize(baBuffer) - lPos
    End If
    If lSize > 0 Then
        ReDim baDest(0 To lSize - 1) As Byte
        If lPos + lSize <= pvArraySize(baBuffer) Then
            Call CopyMemory(baDest(0), baBuffer(lPos), lSize)
        ElseIf lPos < pvArraySize(baBuffer) Then
            Call CopyMemory(baDest(0), baBuffer(lPos), pvArraySize(baBuffer) - lPos)
        End If
    Else
        Erase baDest
    End If
    pvReadArray = lPos + lSize
End Function

'Private Function pvReadRecord(baBuffer() As Byte, ByVal lPos As Long) As Byte()
'    Dim baRetVal()      As Byte
'    Dim lSize           As Long
'
'    lSize = baBuffer(lPos + 3) * &H100& + baBuffer(lPos + 4)
'    pvReadArray baBuffer, lPos + 5, baRetVal, lSize
'    pvReadRecord = baRetVal
'End Function

'= arrays helpers ========================================================

Private Function pvArraySize(baArray() As Byte, Optional RetVal As Long) As Long
    If Peek(ArrPtr(baArray)) <> 0 Then
        RetVal = UBound(baArray) + 1
    Else
        RetVal = 0
    End If
    pvArraySize = RetVal
End Function

Private Function pvArrayXor(baArray() As Byte, ByVal lSeqNo As Long) As Byte()
    Dim baRetVal()      As Byte
    Dim lIdx            As Long
    
    baRetVal = baArray
    lIdx = pvArraySize(baRetVal)
    Do While lSeqNo <> 0 And lIdx > 0
        lIdx = lIdx - 1
        baRetVal(lIdx) = baRetVal(lIdx) Xor (lSeqNo And &HFF)
        lSeqNo = lSeqNo \ &H100
    Loop
    pvArrayXor = baRetVal
End Function

'= global helpers ========================================================

Public Function Peek(ByVal lPtr As Long) As Long
    Call GetMem4(ByVal lPtr, Peek)
End Function

Public Function ToHex(baText() As Byte) As String
    ToHex = ToHexDump(StrConv(baText, vbUnicode))
End Function

Public Function FromHex(sText As String) As Byte()
    FromHex = StrConv(FromHexDump(sText), vbFromUnicode)
End Function

Public Function ToHexDump(sText As String) As String
    Dim lIdx            As Long
    
    For lIdx = 1 To Len(sText)
        ToHexDump = ToHexDump & Right$("0" & Hex$(Asc(Mid$(sText, lIdx, 1))), 2)
    Next
End Function

Public Function FromHexDump(sText As String) As String
    Dim lIdx            As Long
    Dim sRetVal         As String
    
    On Error GoTo EH
    '--- note: sys StrPtr(FromHexDump) = 0 signalizira error
    sRetVal = ""
    For lIdx = 1 To Len(sText) Step 2
        sRetVal = sRetVal & Chr$(CLng("&H" & Mid$(sText, lIdx, 2)))
    Next
    FromHexDump = sRetVal
    Exit Function
EH:
End Function

Public Function DesignDumpMemory(ByVal lPtr As Long, ByVal lSize As Long) As String
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

Public Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function

Private Function EmptyByteArray(Optional ByVal Size As Long) As Byte()
    Dim baRetVal()      As Byte
    
    If Size > 0 Then
        ReDim baRetVal(0 To Size - 1) As Byte
    End If
    EmptyByteArray = baRetVal
End Function
