Attribute VB_Name = "mdTlsSupport"
'=========================================================================
'
' Based on RFC 8446 at https://tools.ietf.org/html/rfc8446
'   and illustrated traffic-dump at https://tls13.ulfheim.net/
'
' More TLS 1.3 implementations at https://github.com/h2o/picotls
'   and https://github.com/openssl/openssl
'
' Additional links with TLS 1.3 resources
'   https://github.com/tlswg/tls13-spec/wiki/Implementations
'   https://sans-io.readthedocs.io/how-to-sans-io.html
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Const TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC       As Long = 20
Private Const TLS_CONTENT_TYPE_ALERT                    As Long = 21
Private Const TLS_CONTENT_TYPE_HANDSHAKE                As Long = 22
Private Const TLS_CONTENT_TYPE_APPDATA                  As Long = 23
Private Const TLS_HANDSHAKE_TYPE_CLIENT_HELLO           As Long = 1
Private Const TLS_HANDSHAKE_TYPE_SERVER_HELLO           As Long = 2
Private Const TLS_HANDSHAKE_TYPE_NEW_SESSION_TICKET     As Long = 4
'Private Const TLS_HANDSHAKE_TYPE_END_OF_EARLY_DATA      As Long = 5
'Private Const TLS_HANDSHAKE_TYPE_ENCRYPTED_EXTENSIONS   As Long = 8
Private Const TLS_HANDSHAKE_TYPE_CERTIFICATE            As Long = 11
Private Const TLS_HANDSHAKE_TYPE_SERVER_KEY_EXCHANGE    As Long = 12
'Private Const TLS_HANDSHAKE_TYPE_CERTIFICATE_REQUEST    As Long = 13
Private Const TLS_HANDSHAKE_TYPE_SERVER_HELLO_DONE      As Long = 14
Private Const TLS_HANDSHAKE_TYPE_CERTIFICATE_VERIFY     As Long = 15
Private Const TLS_HANDSHAKE_TYPE_CLIENT_KEY_EXCHANGE    As Long = 16
Private Const TLS_HANDSHAKE_TYPE_FINISHED               As Long = 20
Private Const TLS_HANDSHAKE_TYPE_KEY_UPDATE             As Long = 24
'Private Const TLS_HANDSHAKE_TYPE_COMPRESSED_CERTIFICATE As Long = 25
Private Const TLS_HANDSHAKE_TYPE_MESSAGE_HASH           As Long = 254
Private Const TLS_EXTENSION_TYPE_SERVER_NAME            As Long = 0
'Private Const TLS_EXTENSION_TYPE_STATUS_REQUEST         As Long = 5
Private Const TLS_EXTENSION_TYPE_SUPPORTED_GROUPS       As Long = 10
Private Const TLS_EXTENSION_TYPE_SIGNATURE_ALGORITHMS   As Long = 13
'Private Const TLS_EXTENSION_TYPE_ALPN                   As Long = 16
'Private Const TLS_EXTENSION_TYPE_COMPRESS_CERTIFICATE   As Long = 27
'Private Const TLS_EXTENSION_TYPE_PRE_SHARED_KEY         As Long = 41
'Private Const TLS_EXTENSION_TYPE_EARLY_DATA             As Long = 42
Private Const TLS_EXTENSION_TYPE_SUPPORTED_VERSIONS     As Long = 43
Private Const TLS_EXTENSION_TYPE_COOKIE                 As Long = 44
'Private Const TLS_EXTENSION_TYPE_PSK_KEY_EXCHANGE_MODES As Long = 45
Private Const TLS_EXTENSION_TYPE_KEY_SHARE              As Long = 51
Private Const TLS_CIPHER_SUITE_AES_128_GCM_SHA256       As Long = &H1301
Private Const TLS_CIPHER_SUITE_AES_256_GCM_SHA384       As Long = &H1302
Private Const TLS_CIPHER_SUITE_CHACHA20_POLY1305_SHA256 As Long = &H1303
Private Const TLS_CIPHER_SUITE_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256 As Long = &HC02B&
Private Const TLS_CIPHER_SUITE_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384 As Long = &HC02C&
Private Const TLS_CIPHER_SUITE_ECDHE_RSA_WITH_AES_128_GCM_SHA256 As Long = &HC02F&
Private Const TLS_CIPHER_SUITE_ECDHE_RSA_WITH_AES_256_GCM_SHA384 As Long = &HC030&
Private Const TLS_CIPHER_SUITE_ECDHE_RSA_WITH_CHACHA20_POLY1305_SHA256 As Long = &HCCA8&
Private Const TLS_CIPHER_SUITE_ECDHE_ECDSA_WITH_CHACHA20_POLY1305_SHA256 As Long = &HCCA9&
Private Const TLS_CIPHER_SUITE_RSA_WITH_AES_128_GCM_SHA256 As Long = &H9C
Private Const TLS_CIPHER_SUITE_RSA_WITH_AES_256_GCM_SHA384 As Long = &H9D
Private Const TLS_GROUP_SECP256R1                       As Long = 23
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
'Private Const TLS_PSK_KE_MODE_PSK_DHE                   As Long = 1
Private Const TLS_PROTOCOL_VERSION_TLS12                As Long = &H303
Private Const TLS_PROTOCOL_VERSION_TLS13_FINAL          As Long = &H304
Private Const TLS_CHACHA20_KEY_SIZE                     As Long = 32
Private Const TLS_CHACHA20POLY1305_IV_SIZE              As Long = 12
Private Const TLS_CHACHA20POLY1305_TAG_SIZE             As Long = 16
Private Const TLS_AES128_KEY_SIZE                       As Long = 16
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
Private Const TLS_RECORD_VERSION                        As Long = TLS_PROTOCOL_VERSION_TLS12 '--- always legacy version
Private Const TLS_CLIENT_LEGACY_VERSION                 As Long = &H303
Private Const TLS_HELLO_RANDOM_SIZE                     As Long = 32

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_VL_ALERTS             As String = "0|Close notify|10|Unexpected message|20|Bad record mac|40|Handshake failure|42|Bad certificate|44|Certificate revoked|45|Certificate expired|46|Certificate unknown|47|Illegal parameter|48|Unknown CA|50|Decode error|51|Decrypt error|70|Protocol version|80|Internal error|90|User canceled|109|Missing extension|112|Unrecognized name|116|Certificate required|120|No application protocol"
Private Const STR_HELLO_RETRY_RANDOM    As String = "cf21ad74e59a6111be1d8c021e65b891c2a211167abb8c5e079e09e2c8a8339c"
Private Const STR_UNKNOWN_ALERT         As String = "Unknown (%1)"
Private Const STR_FORMAT_ALERT          As String = """%1"" alert"
'--- numeric
Private Const LNG_AAD_SIZE              As Long = 5     '--- size of additional authenticated data for TLS 1.3
Private Const LNG_LEGACY_AAD_SIZE       As Long = 13    '--- for TLS 1.2
'--- errors
Private Const ERR_CONNECTION_CLOSED     As String = "Connection closed"
Private Const ERR_GEN_KEYPAIR_FAILED    As String = "Failed generating key pair (%1)"
Private Const ERR_UNSUPPORTED_EX_GROUP  As String = "Unsupported exchange group (%1)"
Private Const ERR_UNSUPPORTED_CIPHER_SUITE As String = "Unsupported cipher suite (%1)"
Private Const ERR_ENCRYPTION_FAILED     As String = "Encryption failed"
Private Const ERR_RECORD_TOO_BIG        As String = "Record size too big"
Private Const ERR_DECRYPTION_FAILED     As String = "Decryption failed"
Private Const ERR_FATAL_ALERT           As String = "Fatal alert"
Private Const ERR_UNEXPECTED_RECORD_TYPE As String = "Unexpected record type (%1)"
Private Const ERR_UNEXPECTED_MSG_TYPE   As String = "Unexpected message type for %1 (%2)"
Private Const ERR_UNEXPECTED_PROTOCOL   As String = "Unexpected protocol for %1 (%2)"
Private Const ERR_SERVER_HANDSHAKE_FAILED As String = "Server Handshake verification failed"
Private Const ERR_INVALID_STATE_HANDSHAKE As String = "Invalid state for handshake content (%1)"
Private Const ERR_INVALID_SIZE_KEY_SHARE As String = "Invalid data size for key share"
Private Const ERR_INVALID_SERVER_KEY    As String = "Invalid server key size"
Private Const ERR_INVALID_SIZE_SERVER_KEY As String = "Invalid data size for server key"
Private Const ERR_INVALID_SIZE_VERSIONS As String = "Invalid data size for supported versions"
Private Const ERR_COOKIE_NOT_ALLOWED    As String = "Cookie not allowed outside HelloRetryRequest"
Private Const ERR_NO_HANDSHAKE_MESSAGES As String = "Missing handshake messages"
Private Const ERR_NO_PREV_SERVER_SECRET As String = "Missing previous server secret"
Private Const ERR_NO_PREV_CLIENT_SECRET As String = "Missing previous client secret"
Private Const ERR_NO_SERVER_RANDOM      As String = "Missing server random"
Private Const ERR_NO_SERVER_CERTIFICATE As String = "Missing server certificate"

Public Enum UcsTlsClientFeaturesEnum '--- bitmask
    ucsTlsSupportTls12 = 2 ^ 0
    ucsTlsSupportTls13 = 2 ^ 1
    ucsTlsSupportAll = -1
End Enum

Public Enum UcsTlsStatesEnum
    ucsTlsStateClosed
    ucsTlsStateHandshakeStart
    ucsTlsStateExpectServerHello
    ucsTlsStateExpectExtensions
    ucsTlsStateExpectServerFinish '--- not used in TLS 1.3
    ucsTlsStatePostHandshake
End Enum

Public Enum UcsTlsCryptoAlgorithmsEnum
    '--- key exchange
    ucsTlsAlgoKeyX25519 = 1
    ucsTlsAlgoKeySecp256r1 = 2
    ucsTlsAlgoKeyCertificate = 3
    '--- authenticated encryption w/ additional data
    ucsTlsAlgoAeadChacha20Poly1305 = 11
    ucsTlsAlgoAeadAes128 = 12
    ucsTlsAlgoAeadAes256 = 13
    '--- digest
    ucsTlsAlgoDigestSha256 = 21
    ucsTlsAlgoDigestSha384 = 22
End Enum

Public Enum UcsTlsAlertDescriptionsEnum
    uscTlsAlertCloseNotify = 0
    uscTlsAlertUnexpectedMessage = 10
    uscTlsAlertBadRecordMac = 20
    uscTlsAlertHandshakeFailure = 40
    uscTlsAlertBadCertificate = 42
    uscTlsAlertCertificateRevoked = 44
    uscTlsAlertCertificateExpired = 45
    uscTlsAlertCertificateUnknown = 46
    uscTlsAlertIllegalParameter = 47
    uscTlsAlertUnknownCa = 48
    uscTlsAlertDecodeError = 50
    uscTlsAlertDecryptError = 51
    uscTlsAlertProtocolVersion = 70
    uscTlsAlertInternalError = 80
    uscTlsAlertUserCanceled = 90
    uscTlsAlertMissingExtension = 109
    uscTlsAlertUnrecognizedName = 112
    uscTlsAlertCertificateRequired = 116
    uscTlsAlertNoApplicationProtocol = 120
End Enum

Public Type UcsTlsContext
    '--- config
    TargetHost          As String
    ClientFeatures      As UcsTlsClientFeaturesEnum
    ResumeSessionID()   As Byte
    ClientRandom()      As Byte
    RandomSize          As Long
    '--- state
    State               As UcsTlsStatesEnum
    LastError           As String
    LastAlertCode       As UcsTlsAlertDescriptionsEnum
    BlocksStack         As Collection
    '--- hello retry request
    HelloRetryCipherSuite As Long
    HelloRetryExchangeGroup As Long
    HelloRetryCookie()  As Byte
    '--- server handshake
    ServerProtocol      As Long
    ServerRandom()      As Byte
    ServerPublic()      As Byte
    ServerCertReqContext() As Byte
    ServerCertificates  As Collection
    ServerSessionID()   As Byte
    '--- crypto settings
    ExchangeGroup       As Long
    ExchangeAlgo        As UcsTlsCryptoAlgorithmsEnum
    ClientPrivate()     As Byte
    ClientPublic()      As Byte
    ClientEncrPrivate() As Byte
    CipherSuite         As Long
    AeadAlgo            As UcsTlsCryptoAlgorithmsEnum
    MacSize             As Long '--- always 0 (not used w/ AEAD ciphers)
    KeySize             As Long
    IvSize              As Long
    IvDynamicSize       As Long '--- only for AES in TLS 1.2
    TagSize             As Long
    DigestAlgo          As UcsTlsCryptoAlgorithmsEnum
    DigestSize          As Long
    '--- bulk secrets
    HandshakeMessages() As Byte '--- ToDo: reduce to HandshakeHash only
    HandshakeSecret()   As Byte
    MasterSecret()      As Byte
    ServerTrafficSecret() As Byte
    ServerTrafficKey()  As Byte
    ServerTrafficIV()   As Byte
    ServerTrafficSeqNo  As Long
    ClientTrafficSecret() As Byte
    ClientTrafficKey()  As Byte
    ClientTrafficIV()   As Byte
    ClientTrafficSeqNo  As Long
    '--- I/O buffers
    RecvBuffer()        As Byte
    RecvPos             As Long
    DecrBuffer()        As Byte
    DecrPos             As Long
    SendBuffer()        As Byte
    SendPos             As Long
    MessBuffer()        As Byte
    MessPos             As Long
    MessSize            As Long
End Type

'=========================================================================
' Methods
'=========================================================================

Public Function TlsInitClient( _
            uCtx As UcsTlsContext, _
            Optional TargetHost As String, _
            Optional ByVal ClientFeatures As UcsTlsClientFeaturesEnum = ucsTlsSupportAll) As Boolean
    Dim uEmpty          As UcsTlsContext
    
    On Error GoTo EH
    If Not CryptoInit() Then
        GoTo QH
    End If
    With uEmpty
        pvSetLastError uEmpty, vbNullString
        .State = ucsTlsStateHandshakeStart
        .TargetHost = TargetHost
        .ClientFeatures = ClientFeatures
        .ClientRandom = pvCryptoRandomArray(TLS_HELLO_RANDOM_SIZE)
        .RandomSize = TLS_HELLO_RANDOM_SIZE
        If (ClientFeatures And ucsTlsSupportTls13) <> 0 Then
            '--- note: TLS 1.3 uses X25519 only and ClientPublic has to be ready for pvBuildClientHello
            If Not pvSetupKeyExchangeEccGroup(uEmpty, TLS_GROUP_X25519, .LastError, .LastAlertCode) Then
                pvSetLastError uCtx, .LastError, .LastAlertCode
                GoTo QH
            End If
        End If
    End With
    uCtx = uEmpty
    '--- success
    TlsInitClient = True
QH:
    Exit Function
EH:
    pvSetLastError uCtx, Err.Description
    Resume QH
End Function

Public Function TlsHandshake(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, baOutput() As Byte, lPos As Long) As Boolean
    On Error GoTo EH
    With uCtx
        If .State = ucsTlsStateClosed Then
            pvSetLastError uCtx, ERR_CONNECTION_CLOSED
            Exit Function
        End If
        pvSetLastError uCtx, vbNullString
        '--- swap-in
        pvArraySwap .SendBuffer, .SendPos, baOutput, lPos
        If .State = ucsTlsStateHandshakeStart Then
            .SendPos = pvBuildClientHello(uCtx, .SendBuffer, .SendPos)
            .State = ucsTlsStateExpectServerHello
        Else
            If lSize < 0 Then
                lSize = pvArraySize(baInput)
            End If
            If Not pvParsePayload(uCtx, baInput, lSize, .LastError, .LastAlertCode) Then
                pvSetLastError uCtx, .LastError, .LastAlertCode
                GoTo QH
            End If
        End If
        '--- success
        TlsHandshake = True
QH:
        '--- swap-out
        pvArraySwap baOutput, lPos, .SendBuffer, .SendPos
    End With
    Exit Function
EH:
    pvSetLastError uCtx, Err.Description
    Resume QH
End Function

Public Function TlsSend(uCtx As UcsTlsContext, baPlainText() As Byte, ByVal lSize As Long, baOutput() As Byte, lPos As Long) As Boolean
    On Error GoTo EH
    With uCtx
        If .State = ucsTlsStateClosed Then
            pvSetLastError uCtx, ERR_CONNECTION_CLOSED
            Exit Function
        End If
        pvSetLastError uCtx, vbNullString
        '--- swap-in
        pvArraySwap .SendBuffer, .SendPos, baOutput, lPos
        If lSize < 0 Then
            lSize = pvArraySize(baPlainText)
        End If
        .SendPos = pvBuildClientApplicationData(uCtx, .SendBuffer, .SendPos, baPlainText, lSize, .LastError, .LastAlertCode)
        If LenB(.LastError) <> 0 Then
            pvSetLastError uCtx, .LastError, .LastAlertCode
            GoTo QH
        End If
        '--- success
        TlsSend = True
QH:
        '--- swap-out
        pvArraySwap baOutput, lPos, .SendBuffer, .SendPos
    End With
    Exit Function
EH:
    pvSetLastError uCtx, Err.Description
    Resume QH
End Function

Public Function TlsReceive(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, baPlainText() As Byte, lPos As Long) As Boolean
    On Error GoTo EH
    With uCtx
        If .State = ucsTlsStateClosed Then
            pvSetLastError uCtx, ERR_CONNECTION_CLOSED
            Exit Function
        End If
        pvSetLastError uCtx, vbNullString
        '--- swap-in
        pvArraySwap .DecrBuffer, .DecrPos, baPlainText, lPos
        If lSize < 0 Then
            lSize = pvArraySize(baInput)
        End If
        If Not pvParsePayload(uCtx, baInput, lSize, .LastError, .LastAlertCode) Then
            pvSetLastError uCtx, .LastError, .LastAlertCode
            GoTo QH
        End If
        '--- success
        TlsReceive = True
QH:
        '--- swap-out
        pvArraySwap baPlainText, lPos, .DecrBuffer, .DecrPos
    End With
    Exit Function
EH:
    pvSetLastError uCtx, Err.Description
    Resume QH
End Function

Public Function TlsShutdown(uCtx As UcsTlsContext, baOutput() As Byte, lPos As Long) As Boolean
    On Error GoTo EH
    With uCtx
        If .State = ucsTlsStateClosed Then
            Exit Function
        End If
        pvSetLastError uCtx, vbNullString
        '--- swap-in
        pvArraySwap .SendBuffer, .SendPos, baOutput, lPos
        .SendPos = pvBuildClientAlert(uCtx, .SendBuffer, .SendPos, uscTlsAlertCloseNotify, TLS_ALERT_LEVEL_WARNING, .LastError, .LastAlertCode)
        If LenB(.LastError) <> 0 Then
            pvSetLastError uCtx, .LastError, .LastAlertCode
            GoTo QH
        End If
        .State = ucsTlsStateClosed
        '--- success
        TlsShutdown = True
QH:
        '--- swap-out
        pvArraySwap baOutput, lPos, .SendBuffer, .SendPos
    End With
    Exit Function
EH:
    pvSetLastError uCtx, Err.Description
    Resume QH
End Function

Public Function TlsIsClosed(uCtx As UcsTlsContext) As Boolean
    TlsIsClosed = (uCtx.State = ucsTlsStateClosed)
End Function

Public Function TlsIsReady(uCtx As UcsTlsContext) As Boolean
    TlsIsReady = (uCtx.State = ucsTlsStatePostHandshake)
End Function

Public Function TlsGetLastError(uCtx As UcsTlsContext) As String
    TlsGetLastError = uCtx.LastError
    If uCtx.LastAlertCode <> -1 Then
        TlsGetLastError = IIf(LenB(TlsGetLastError) <> 0, TlsGetLastError & ". ", vbNullString) & Replace(STR_FORMAT_ALERT, "%1", TlsGetLastAlert(uCtx))
    End If
End Function

Public Function TlsGetLastAlert(uCtx As UcsTlsContext, Optional AlertCode As UcsTlsAlertDescriptionsEnum) As String
    Static vTexts       As Variant
    
    AlertCode = uCtx.LastAlertCode
    If AlertCode >= 0 Then
        If IsEmpty(vTexts) Then
            vTexts = SplitOrReindex(STR_VL_ALERTS, "|")
        End If
        If AlertCode <= UBound(vTexts) Then
            TlsGetLastAlert = vTexts(AlertCode)
        End If
        If LenB(TlsGetLastAlert) = 0 Then
            TlsGetLastAlert = Replace(STR_UNKNOWN_ALERT, "%1", AlertCode)
        End If
    End If
End Function

'= private ===============================================================

Private Function pvSetupKeyExchangeEccGroup(uCtx As UcsTlsContext, ByVal lExchangeGroup As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    With uCtx
        If .ExchangeGroup <> lExchangeGroup Then
            .ExchangeGroup = lExchangeGroup
            Select Case lExchangeGroup
            Case TLS_GROUP_X25519
                .ExchangeAlgo = ucsTlsAlgoKeyX25519
                If Not CryptoEccCurve25519MakeKey(.ClientPrivate, .ClientPublic) Then
                    sError = Replace(ERR_GEN_KEYPAIR_FAILED, "%1", "Curve25519")
                    eAlertCode = uscTlsAlertInternalError
                    GoTo QH
                End If
            Case TLS_GROUP_SECP256R1
                .ExchangeAlgo = ucsTlsAlgoKeySecp256r1
                If Not CryptoEccSecp256r1MakeKey(.ClientPrivate, .ClientPublic) Then
                    sError = Replace(ERR_GEN_KEYPAIR_FAILED, "%1", "secp256r1")
                    eAlertCode = uscTlsAlertInternalError
                    GoTo QH
                End If
            Case Else
                sError = Replace(ERR_UNSUPPORTED_EX_GROUP, "%1", "0x" & Hex$(.ExchangeGroup))
                eAlertCode = uscTlsAlertInternalError
                GoTo QH
            End Select
        End If
    End With
    '--- success
    pvSetupKeyExchangeEccGroup = True
QH:
End Function

Private Function pvSetupKeyExchangeRsaCertificate(uCtx As UcsTlsContext, baCert() As Byte, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    On Error GoTo EH
    With uCtx
        .ExchangeAlgo = ucsTlsAlgoKeyCertificate
        .ClientPrivate = pvCryptoRandomArray(48)
        pvWriteLong .ClientPrivate, 0, TLS_CLIENT_LEGACY_VERSION, Size:=2
        .ClientEncrPrivate = CryptoRsaEncrypt(baCert, .ClientPrivate)
    End With
    '--- success
    pvSetupKeyExchangeRsaCertificate = True
    Exit Function
EH:
    sError = Trim$(Replace(Err.Description, vbLf, ". "))
    If Right$(sError, 1) = "." Then
        sError = Left$(sError, Len(sError) - 1)
    End If
    sError = sError & " in " & Err.Source
    eAlertCode = uscTlsAlertInternalError
End Function

Private Function pvSetupCipherSuite(uCtx As UcsTlsContext, ByVal lCipherSuite As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    With uCtx
        If .CipherSuite <> lCipherSuite Then
            .CipherSuite = lCipherSuite
            Select Case lCipherSuite
            Case TLS_CIPHER_SUITE_CHACHA20_POLY1305_SHA256, TLS_CIPHER_SUITE_ECDHE_RSA_WITH_CHACHA20_POLY1305_SHA256, TLS_CIPHER_SUITE_ECDHE_ECDSA_WITH_CHACHA20_POLY1305_SHA256
                .AeadAlgo = ucsTlsAlgoAeadChacha20Poly1305
                .KeySize = TLS_CHACHA20_KEY_SIZE
                .IvSize = TLS_CHACHA20POLY1305_IV_SIZE
                .TagSize = TLS_CHACHA20POLY1305_TAG_SIZE
                .DigestAlgo = ucsTlsAlgoDigestSha256
                .DigestSize = TLS_SHA256_DIGEST_SIZE
            Case TLS_CIPHER_SUITE_AES_128_GCM_SHA256, TLS_CIPHER_SUITE_ECDHE_RSA_WITH_AES_128_GCM_SHA256, TLS_CIPHER_SUITE_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256, TLS_CIPHER_SUITE_RSA_WITH_AES_128_GCM_SHA256
                .AeadAlgo = ucsTlsAlgoAeadAes128
                .KeySize = TLS_AES128_KEY_SIZE
                .IvSize = TLS_AESGCM_IV_SIZE
                If lCipherSuite <> TLS_CIPHER_SUITE_AES_128_GCM_SHA256 Then
                    .IvDynamicSize = 8 '--- AES in TLS 1.2
                End If
                .TagSize = TLS_AESGCM_TAG_SIZE
                .DigestAlgo = ucsTlsAlgoDigestSha256
                .DigestSize = TLS_SHA256_DIGEST_SIZE
            Case TLS_CIPHER_SUITE_AES_256_GCM_SHA384, TLS_CIPHER_SUITE_ECDHE_RSA_WITH_AES_256_GCM_SHA384, TLS_CIPHER_SUITE_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384, TLS_CIPHER_SUITE_RSA_WITH_AES_256_GCM_SHA384
                .AeadAlgo = ucsTlsAlgoAeadAes256
                .KeySize = TLS_AES256_KEY_SIZE
                .IvSize = TLS_AESGCM_IV_SIZE
                If lCipherSuite <> TLS_CIPHER_SUITE_AES_256_GCM_SHA384 Then
                    .IvDynamicSize = 8 '--- AES in TLS 1.2
                End If
                .TagSize = TLS_AESGCM_TAG_SIZE
                .DigestAlgo = ucsTlsAlgoDigestSha384
                .DigestSize = TLS_SHA384_DIGEST_SIZE
            Case Else
                sError = Replace(ERR_UNSUPPORTED_CIPHER_SUITE, "%1", "0x" & Hex$(.CipherSuite))
                eAlertCode = uscTlsAlertInternalError
                GoTo QH
            End Select
        End If
    End With
    '--- success
    pvSetupCipherSuite = True
QH:
End Function

Private Function pvBuildClientHello(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long) As Long
    Dim lMessagePos     As Long
    Dim vElem           As Variant
    
    With uCtx
        '--- Record Header
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            lMessagePos = lPos
            '--- Handshake Header
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_CLIENT_HELLO)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                lPos = pvWriteLong(baOutput, lPos, TLS_CLIENT_LEGACY_VERSION, Size:=2)
                lPos = pvWriteArray(baOutput, lPos, uCtx.ClientRandom)
                '--- Legacy Session ID
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack)
                    lPos = pvWriteArray(baOutput, lPos, uCtx.ResumeSessionID)
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                '--- Cipher Suites
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                    For Each vElem In pvCryptoOrderCiphers(.ClientFeatures)
                        lPos = pvWriteLong(baOutput, lPos, vElem, Size:=2)
                    Next
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                '--- Legacy Compression Methods
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack)
                    lPos = pvWriteLong(baOutput, lPos, TLS_COMPRESS_NULL)
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                '--- Extensions
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                    If LenB(uCtx.TargetHost) <> 0 Then
                        '--- Extension - Server Name
                        lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_SERVER_NAME, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                                lPos = pvWriteLong(baOutput, lPos, TLS_SERVER_NAME_TYPE_HOSTNAME) '--- FQDN
                                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                                    lPos = pvWriteString(baOutput, lPos, uCtx.TargetHost)
                                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    End If
                    '--- Extension - Supported Groups
                    lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_SUPPORTED_GROUPS, Size:=2)
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            If CryptoIsSupported(ucsTlsAlgoKeyX25519) Then
                                If .HelloRetryExchangeGroup = 0 Or .HelloRetryExchangeGroup = TLS_GROUP_X25519 Then
                                    lPos = pvWriteLong(baOutput, lPos, TLS_GROUP_X25519, Size:=2)
                                End If
                            End If
                            If (.ClientFeatures And ucsTlsSupportTls12) <> 0 And CryptoIsSupported(ucsTlsAlgoKeySecp256r1) Then
                                If .HelloRetryExchangeGroup = 0 Or .HelloRetryExchangeGroup = TLS_GROUP_SECP256R1 Then
                                    lPos = pvWriteLong(baOutput, lPos, TLS_GROUP_SECP256R1, Size:=2)
                                End If
                            End If
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    If (.ClientFeatures And ucsTlsSupportTls12) <> 0 Then
                        '--- Extension - EC Point Formats
                        lPos = pvWriteArray(baOutput, lPos, FromHex("00:0B:00:02:01:00"))   '--- uncompressed only
                        '--- Extension - Renegotiation Info
                        lPos = pvWriteArray(baOutput, lPos, FromHex("FF:01:00:01:00"))      '--- empty info
                    End If
                    '--- Extension - Signature Algorithms
                    lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_SIGNATURE_ALGORITHMS, Size:=2)
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, Size:=2)
                            lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_ECDSA_SECP256R1_SHA256, Size:=2)
                            lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PKCS1_SHA256, Size:=2)
                            lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PKCS1_SHA1, Size:=2)
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    If (.ClientFeatures And ucsTlsSupportTls13) <> 0 Then
                        '--- Extension - Key Share
                        lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_KEY_SHARE, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                                lPos = pvWriteLong(baOutput, lPos, TLS_GROUP_X25519, Size:=2)
                                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                                    lPos = pvWriteArray(baOutput, lPos, uCtx.ClientPublic)
                                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        '--- Extension - Supported Versions
                        lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_SUPPORTED_VERSIONS, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack)
                                lPos = pvWriteLong(baOutput, lPos, TLS_PROTOCOL_VERSION_TLS13_FINAL, Size:=2)
                                If (.ClientFeatures And ucsTlsSupportTls12) <> 0 Then
                                    lPos = pvWriteLong(baOutput, lPos, TLS_PROTOCOL_VERSION_TLS12, Size:=2)
                                End If
                            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        If pvArraySize(.HelloRetryCookie) > 0 Then
                            '--- Extension - Cookie
                            lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_COOKIE, Size:=2)
                            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack)
                                    lPos = pvWriteArray(baOutput, lPos, .HelloRetryCookie)
                                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        End If
                    End If
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lMessagePos)), lPos - lMessagePos
    End With
    pvBuildClientHello = lPos
End Function

Private Function pvBuildLegacyClientKeyExchange(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Long
    Dim baClientIV()    As Byte
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim baHandshakeHash() As Byte
    Dim baVerifyData()  As Byte
    Dim baAad()         As Byte
    Dim lAadPos         As Long
    Dim lRecordPos      As Long
    
    With uCtx
        '--- Record Header
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            lMessagePos = lPos
            '--- Handshake Client Key Exchange
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_CLIENT_KEY_EXCHANGE)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                If pvArraySize(.ClientEncrPrivate) > 0 Then
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                        lPos = pvWriteArray(baOutput, lPos, .ClientEncrPrivate)
                    lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                Else
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack)
                        lPos = pvWriteArray(baOutput, lPos, .ClientPublic)
                    lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                End If
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lMessagePos)), lPos - lMessagePos
        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        '--- Legacy Change Cipher Spec
        lPos = pvWriteArray(baOutput, lPos, FromHex("14:03:03:00:01:01"))
        '--- Record Header
        lRecordPos = lPos
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            If .IvDynamicSize > 0 Then '--- AES in TLS 1.2
                ReDim baClientIV(0 To .IvSize - 1) As Byte
                pvWriteArray baClientIV, 0, .ClientTrafficIV
                pvWriteArray baClientIV, .IvSize - .IvDynamicSize, pvCryptoRandomArray(.IvDynamicSize)
                lPos = pvWriteBuffer(baOutput, lPos, VarPtr(baClientIV(.IvSize - .IvDynamicSize)), .IvDynamicSize)
            Else
                baClientIV = pvArrayXor(.ClientTrafficIV, .ClientTrafficSeqNo)
            End If
            lMessagePos = lPos
            '--- Handshake Finish
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_FINISHED)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                baHandshakeHash = pvCryptoHash(.DigestAlgo, .HandshakeMessages, 0)
                baVerifyData = pvKdfLegacyTls1Prf(.DigestAlgo, .MasterSecret, "client finished", baHandshakeHash, 12)
                lPos = pvWriteArray(baOutput, lPos, baVerifyData)
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            lMessageSize = lPos - lMessagePos
            '--- note: *before* allocating space for the authentication tag
            pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lMessagePos)), lPos - lMessagePos
            lPos = pvWriteReserved(baOutput, lPos, .TagSize)
            '--- encrypt message
            ReDim baAad(0 To LNG_LEGACY_AAD_SIZE - 1) As Byte
            lAadPos = pvWriteLong(baAad, 0, 0, Size:=4)
            lAadPos = pvWriteLong(baAad, lAadPos, .ClientTrafficSeqNo, Size:=4)
            lAadPos = pvWriteBuffer(baAad, lAadPos, VarPtr(baOutput(lRecordPos)), 3)
            lAadPos = pvWriteLong(baAad, lAadPos, lMessageSize, Size:=2)
            Debug.Assert lAadPos = LNG_LEGACY_AAD_SIZE
            If Not pvCryptoAeadEncrypt(.AeadAlgo, baClientIV, .ClientTrafficKey, baAad, 0, UBound(baAad) + 1, baOutput, lMessagePos, lMessageSize) Then
                sError = ERR_ENCRYPTION_FAILED
                eAlertCode = uscTlsAlertInternalError
                GoTo QH
            End If
            .ClientTrafficSeqNo = .ClientTrafficSeqNo + 1
            lMessagePos = lRecordPos + 5
        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
    End With
    pvBuildLegacyClientKeyExchange = lPos
QH:
End Function

Private Function pvBuildClientHandshakeFinished(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Long
    Dim lRecordPos      As Long
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim baClientIV()    As Byte
    Dim baHandshakeHash() As Byte
    Dim baVerifyData()  As Byte
    
    With uCtx
        '--- Legacy Change Cipher Spec
        lPos = pvWriteArray(baOutput, lPos, FromHex("14:03:03:00:01:01"))
        '--- Record Header
        lRecordPos = lPos
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_APPDATA)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            lMessagePos = lPos
            '--- Handshake Finish
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_FINISHED)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                baHandshakeHash = pvCryptoHash(.DigestAlgo, .HandshakeMessages, 0)
                baVerifyData = pvHkdfExpandLabel(.DigestAlgo, .ClientTrafficSecret, "finished", EmptyByteArray, .DigestSize)
                baVerifyData = pvHkdfExtract(.DigestAlgo, baVerifyData, baHandshakeHash)
                lPos = pvWriteArray(baOutput, lPos, baVerifyData)
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
            lMessageSize = lPos - lMessagePos
            lPos = pvWriteReserved(baOutput, lPos, .TagSize)
        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        baClientIV = pvArrayXor(.ClientTrafficIV, .ClientTrafficSeqNo)
        If Not pvCryptoAeadEncrypt(.AeadAlgo, baClientIV, .ClientTrafficKey, baOutput, lRecordPos, LNG_AAD_SIZE, baOutput, lMessagePos, lMessageSize) Then
            sError = ERR_ENCRYPTION_FAILED
            eAlertCode = uscTlsAlertInternalError
            GoTo QH
        End If
        .ClientTrafficSeqNo = .ClientTrafficSeqNo + 1
    End With
    pvBuildClientHandshakeFinished = lPos
QH:
End Function

Private Function pvBuildClientApplicationData(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long, baData() As Byte, ByVal lSize As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Long
    Dim lRecordPos      As Long
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim baClientIV()    As Byte
    Dim baAad()         As Byte
    Dim lAadPos         As Long
    Dim bResult         As Boolean
    
    With uCtx
        lRecordPos = lPos
        '--- Record Header
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_APPDATA)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            If .IvDynamicSize > 0 Then '--- AES in TLS 1.2
                ReDim baClientIV(0 To .IvSize - 1) As Byte
                pvWriteArray baClientIV, 0, .ClientTrafficIV
                pvWriteArray baClientIV, .IvSize - .IvDynamicSize, pvCryptoRandomArray(.IvDynamicSize)
                lPos = pvWriteBuffer(baOutput, lPos, VarPtr(baClientIV(.IvSize - .IvDynamicSize)), .IvDynamicSize)
            Else
                baClientIV = pvArrayXor(.ClientTrafficIV, .ClientTrafficSeqNo)
            End If
            lMessagePos = lPos
            If lSize > 0 Then
                lPos = pvWriteBuffer(baOutput, lPos, VarPtr(baData(0)), lSize)
            End If
            If .ServerProtocol = TLS_PROTOCOL_VERSION_TLS13_FINAL Then
                lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_APPDATA)
            End If
            lMessageSize = lPos - lMessagePos
            lPos = pvWriteReserved(baOutput, lPos, .TagSize)
        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        '--- encrypt message
        If .ServerProtocol = TLS_PROTOCOL_VERSION_TLS13_FINAL Then
            bResult = pvCryptoAeadEncrypt(.AeadAlgo, baClientIV, .ClientTrafficKey, baOutput, lRecordPos, LNG_AAD_SIZE, baOutput, lMessagePos, lMessageSize)
        ElseIf .ServerProtocol = TLS_PROTOCOL_VERSION_TLS12 Then
            ReDim baAad(0 To LNG_LEGACY_AAD_SIZE - 1) As Byte
            lAadPos = pvWriteLong(baAad, 0, 0, Size:=4)
            lAadPos = pvWriteLong(baAad, lAadPos, .ClientTrafficSeqNo, Size:=4)
            lAadPos = pvWriteBuffer(baAad, lAadPos, VarPtr(baOutput(lRecordPos)), 3)
            lAadPos = pvWriteLong(baAad, lAadPos, lMessageSize, Size:=2)
            Debug.Assert lAadPos = LNG_LEGACY_AAD_SIZE
            bResult = pvCryptoAeadEncrypt(.AeadAlgo, baClientIV, .ClientTrafficKey, baAad, 0, UBound(baAad) + 1, baOutput, lMessagePos, lMessageSize)
        End If
        If Not bResult Then
            sError = ERR_ENCRYPTION_FAILED
            eAlertCode = uscTlsAlertInternalError
            GoTo QH
        End If
        .ClientTrafficSeqNo = .ClientTrafficSeqNo + 1
    End With
    pvBuildClientApplicationData = lPos
QH:
End Function

Private Function pvBuildClientAlert(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long, ByVal eAlertDesc As UcsTlsAlertDescriptionsEnum, ByVal lAlertLevel As Long, Optional sError As String, Optional eAlertCode As UcsTlsAlertDescriptionsEnum) As Long
    Dim lRecordPos      As Long
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim baClientIV()    As Byte
    Dim baAad()         As Byte
    Dim lAadPos         As Long
    
    With uCtx
        '--- for TLS 1.3 -> tunnel alert through application data encryption
        If .State = ucsTlsStatePostHandshake And .ServerProtocol = TLS_PROTOCOL_VERSION_TLS13_FINAL Then
            ReDim baClientIV(0 To 3) As Byte
            baClientIV(0) = eAlertDesc
            baClientIV(1) = lAlertLevel
            baClientIV(2) = TLS_CONTENT_TYPE_ALERT
            pvBuildClientAlert = pvBuildClientApplicationData(uCtx, baOutput, lPos, baClientIV, UBound(baClientIV) + 1, sError, eAlertCode)
            GoTo QH
        End If
        lRecordPos = lPos
        '--- Record Header
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_ALERT)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            If .State = ucsTlsStatePostHandshake And .ServerProtocol = TLS_PROTOCOL_VERSION_TLS12 Then
                If .IvDynamicSize > 0 Then '--- AES in TLS 1.2
                    ReDim baClientIV(0 To .IvSize - 1) As Byte
                    pvWriteArray baClientIV, 0, .ClientTrafficIV
                    pvWriteArray baClientIV, .IvSize - .IvDynamicSize, pvCryptoRandomArray(.IvDynamicSize)
                    lPos = pvWriteBuffer(baOutput, lPos, VarPtr(baClientIV(.IvSize - .IvDynamicSize)), .IvDynamicSize)
                Else
                    baClientIV = pvArrayXor(.ClientTrafficIV, .ClientTrafficSeqNo)
                End If
            End If
            lMessagePos = lPos
            lPos = pvWriteLong(baOutput, lPos, eAlertDesc)
            lPos = pvWriteLong(baOutput, lPos, lAlertLevel)
            lMessageSize = lPos - lMessagePos
            If .State = ucsTlsStatePostHandshake And .ServerProtocol = TLS_PROTOCOL_VERSION_TLS12 Then
                lPos = pvWriteReserved(baOutput, lPos, .TagSize)
            End If
        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        If .State = ucsTlsStatePostHandshake And .ServerProtocol = TLS_PROTOCOL_VERSION_TLS12 Then
            '--- encrypt message
            ReDim baAad(0 To LNG_LEGACY_AAD_SIZE - 1) As Byte
            lAadPos = pvWriteLong(baAad, 0, 0, Size:=4)
            lAadPos = pvWriteLong(baAad, lAadPos, .ClientTrafficSeqNo, Size:=4)
            lAadPos = pvWriteBuffer(baAad, lAadPos, VarPtr(baOutput(lRecordPos)), 3)
            lAadPos = pvWriteLong(baAad, lAadPos, lMessageSize, Size:=2)
            Debug.Assert lAadPos = LNG_LEGACY_AAD_SIZE
            If Not pvCryptoAeadEncrypt(.AeadAlgo, baClientIV, .ClientTrafficKey, baAad, 0, UBound(baAad) + 1, baOutput, lMessagePos, lMessageSize) Then
                sError = ERR_ENCRYPTION_FAILED
                eAlertCode = uscTlsAlertInternalError
                GoTo QH
            End If
            .ClientTrafficSeqNo = .ClientTrafficSeqNo + 1
        End If
    End With
    pvBuildClientAlert = lPos
QH:
End Function

Private Function pvParsePayload(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Dim lRecvPos        As Long
    Dim lRecvSize       As Long
    
    If lSize > 0 Then
    With uCtx
        .RecvPos = pvWriteBuffer(.RecvBuffer, .RecvPos, VarPtr(baInput(0)), lSize)
        lRecvPos = pvParseRecord(uCtx, .RecvBuffer, .RecvPos, sError, eAlertCode)
        If LenB(sError) <> 0 Then
            GoTo QH
        End If
        lRecvSize = .RecvPos - lRecvPos
        If lRecvPos > 0 And lRecvSize > 0 Then
            Call CopyMemory(.RecvBuffer(0), .RecvBuffer(lRecvPos), lRecvSize)
        End If
        .RecvPos = IIf(lRecvSize > 0, lRecvSize, 0)
    End With
    End If
    '--- success
    pvParsePayload = True
QH:
End Function

Private Function pvParseRecord(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Long
    Dim lRecordPos      As Long
    Dim lRecordSize     As Long
    Dim lRecordType     As Long
    Dim lRecordProtocol As Long
    Dim baServerIV()    As Byte
    Dim lPos            As Long
    Dim lEnd            As Long
    Dim baAad()         As Byte
    Dim bResult         As Byte
    
    With uCtx
    Do While lPos + 6 <= lSize
        lRecordPos = lPos
        lPos = pvReadLong(baInput, lPos, lRecordType)
        lPos = pvReadLong(baInput, lPos, lRecordProtocol, Size:=2)
        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lRecordSize)
            If lRecordSize > IIf(lRecordType = TLS_CONTENT_TYPE_APPDATA, TLS_MAX_ENCRYPTED_RECORD_SIZE, TLS_MAX_PLAINTEXT_RECORD_SIZE) Then
                sError = ERR_RECORD_TOO_BIG
                eAlertCode = uscTlsAlertDecodeError
                GoTo QH
            End If
            If lPos + lRecordSize > lSize Then
                '--- back off and bail out early
                lPos = pvReadEndOfBlock(baInput, lPos + lRecordSize, .BlocksStack)
                lPos = lRecordPos
                Exit Do
            End If
            Select Case lRecordType
            Case TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC
                lPos = lPos + lRecordSize
            Case TLS_CONTENT_TYPE_ALERT
                lEnd = lPos + lRecordSize
                If lRecordSize > 2 Then
                    If .ServerProtocol = TLS_PROTOCOL_VERSION_TLS13_FINAL Then
                        '--- note: TLS_CONTENT_TYPE_ALERT encryption is tunneled through TLS_CONTENT_TYPE_APPDATA
                        sError = ERR_RECORD_TOO_BIG
                        eAlertCode = uscTlsAlertDecodeError
                        GoTo QH
                    ElseIf .ServerProtocol = TLS_PROTOCOL_VERSION_TLS12 Then
                        pvPrepareLegacyDecryptParams uCtx, baInput, lRecordPos, lRecordSize, lPos, lEnd, baServerIV, baAad
                        bResult = pvCryptoAeadDecrypt(.AeadAlgo, baServerIV, .ServerTrafficKey, baAad, 0, UBound(baAad) + 1, baInput, lPos, lEnd - lPos + .TagSize)
                    Else
                        bResult = False
                    End If
                    If Not bResult Then
                        sError = ERR_DECRYPTION_FAILED
                        eAlertCode = uscTlsAlertBadRecordMac
                        GoTo QH
                    End If
                    .ServerTrafficSeqNo = .ServerTrafficSeqNo + 1
                End If
HandleAlertContent:
                If lPos + 1 < lEnd Then
                    Select Case baInput(lPos)
                    Case TLS_ALERT_LEVEL_FATAL
                        sError = ERR_FATAL_ALERT
                        eAlertCode = baInput(lPos + 1)
                        GoTo QH
                    Case TLS_ALERT_LEVEL_WARNING
                        .LastAlertCode = baInput(lPos + 1)
                        Debug.Print TlsGetLastAlert(uCtx) & " (TLS_ALERT_LEVEL_WARNING)", Timer
                        If .LastAlertCode = uscTlsAlertCloseNotify Then
                            .State = ucsTlsStateClosed
                        End If
                    End Select
                End If
                '--- note: skip AEAD's authentication tag
                lPos = lRecordPos + lRecordSize + 5
            Case TLS_CONTENT_TYPE_HANDSHAKE
                lEnd = lPos + lRecordSize
                If .State = ucsTlsStateExpectServerFinish Then
                    If .ServerProtocol = TLS_PROTOCOL_VERSION_TLS13_FINAL Then
                        '--- note: ucsTlsStateExpectServerFinish is TLS 1.2 state only
                        sError = Replace(Replace(ERR_UNEXPECTED_PROTOCOL, "%1", "ucsTlsStateExpectServerFinish"), "%2", .ServerProtocol)
                        eAlertCode = uscTlsAlertInternalError
                        GoTo QH
                    ElseIf .ServerProtocol = TLS_PROTOCOL_VERSION_TLS12 Then
                        pvPrepareLegacyDecryptParams uCtx, baInput, lRecordPos, lRecordSize, lPos, lEnd, baServerIV, baAad
                        bResult = pvCryptoAeadDecrypt(.AeadAlgo, baServerIV, .ServerTrafficKey, baAad, 0, UBound(baAad) + 1, baInput, lPos, lEnd - lPos + .TagSize)
                    Else
                        bResult = False
                    End If
                    If Not bResult Then
                        sError = ERR_DECRYPTION_FAILED
                        eAlertCode = uscTlsAlertBadRecordMac
                        GoTo QH
                    End If
                    .ServerTrafficSeqNo = .ServerTrafficSeqNo + 1
                End If
HandleHandshakeContent:
                If .MessSize > 0 Then
                    .MessSize = pvWriteBuffer(.MessBuffer, .MessSize, VarPtr(baInput(lPos)), lEnd - lPos)
                    .MessPos = pvParseHandshakeContent(uCtx, .MessBuffer, .MessPos, .MessSize, lRecordProtocol, sError, eAlertCode)
                    If LenB(sError) <> 0 Then
                        GoTo QH
                    End If
                    If .MessPos >= .MessSize Then
                        Erase .MessBuffer
                        .MessSize = 0
                        .MessPos = 0
                    End If
                Else
                    lPos = pvParseHandshakeContent(uCtx, baInput, lPos, lEnd, lRecordProtocol, sError, eAlertCode)
                    If LenB(sError) <> 0 Then
                        GoTo QH
                    End If
                    If lPos < lEnd Then
                        .MessSize = pvWriteBuffer(.MessBuffer, .MessSize, VarPtr(baInput(lPos)), lEnd - lPos)
                        .MessPos = 0
                    End If
                End If
                '--- note: skip AEAD's authentication tag
                lPos = lRecordPos + lRecordSize + 5
            Case TLS_CONTENT_TYPE_APPDATA
                If .ServerProtocol = TLS_PROTOCOL_VERSION_TLS13_FINAL Then
                    baServerIV = pvArrayXor(.ServerTrafficIV, .ServerTrafficSeqNo)
                    bResult = pvCryptoAeadDecrypt(.AeadAlgo, baServerIV, .ServerTrafficKey, baInput, lRecordPos, LNG_AAD_SIZE, baInput, lPos, lRecordSize)
                ElseIf .ServerProtocol = TLS_PROTOCOL_VERSION_TLS12 Then
                    pvPrepareLegacyDecryptParams uCtx, baInput, lRecordPos, lRecordSize, lPos, lEnd, baServerIV, baAad
                    bResult = pvCryptoAeadDecrypt(.AeadAlgo, baServerIV, .ServerTrafficKey, baAad, 0, UBound(baAad) + 1, baInput, lPos, lEnd - lPos + .TagSize)
                Else
                    bResult = False
                End If
                If Not bResult Then
                    sError = ERR_DECRYPTION_FAILED
                    eAlertCode = uscTlsAlertBadRecordMac
                    GoTo QH
                End If
                .ServerTrafficSeqNo = .ServerTrafficSeqNo + 1
                If .ServerProtocol = TLS_PROTOCOL_VERSION_TLS13_FINAL Then
                    lEnd = lPos + lRecordSize - .TagSize - 1
                    '--- trim zero padding at the end of decrypted record
                    Do While baInput(lEnd) = 0
                        lEnd = lEnd - 1
                    Loop
                    lRecordType = baInput(lEnd)
                    Select Case lRecordType
                    Case TLS_CONTENT_TYPE_ALERT
                        GoTo HandleAlertContent
                    Case TLS_CONTENT_TYPE_HANDSHAKE
                        GoTo HandleHandshakeContent
                    Case TLS_CONTENT_TYPE_APPDATA
                        '--- do nothing
                    Case Else
                        sError = Replace(ERR_UNEXPECTED_RECORD_TYPE, "%1", lRecordType)
                        eAlertCode = uscTlsAlertHandshakeFailure
                        GoTo QH
                    End Select
                End If
                .DecrPos = pvWriteBuffer(.DecrBuffer, .DecrPos, VarPtr(baInput(lPos)), lEnd - lPos)
                '--- note: skip AEAD's authentication tag or zero padding
                lPos = lRecordPos + lRecordSize + 5
            Case Else
                sError = Replace(ERR_UNEXPECTED_RECORD_TYPE, "%1", lRecordType)
                eAlertCode = uscTlsAlertHandshakeFailure
                GoTo QH
            End Select
        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
    Loop
    End With
    pvParseRecord = lPos
QH:
End Function

Private Function pvParseHandshakeContent(uCtx As UcsTlsContext, baInput() As Byte, ByVal lPos As Long, ByVal lEnd As Long, ByVal lRecordProtocol As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Long
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim lMessageType    As Long
    Dim baMessage()     As Byte
    Dim baHandshakeHash() As Byte
    Dim baVerifyData()  As Byte
    Dim lVerifyPos      As Long
    Dim lRequestUpdate  As Long
    Dim lCurveType      As Long
    Dim lNamedCurve     As Long
    Dim lSignatureType  As Long
    Dim lSignatureSize  As Long
    Dim baSignature()   As Byte
    Dim baCert()        As Byte
    Dim lCertSize       As Long
    Dim lCertEnd        As Long
    
    With uCtx
        Do While lPos < lEnd
            lMessagePos = lPos
            lPos = pvReadLong(baInput, lPos, lMessageType)
            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=3, BlockSize:=lMessageSize)
                If lPos + lMessageSize > lEnd Then
                    '--- back off and bail out early
                    lPos = pvReadEndOfBlock(baInput, lPos + lMessageSize, .BlocksStack)
                    lPos = lMessagePos
                    Exit Do
                End If
                Select Case .State
                Case ucsTlsStateExpectServerHello
                    Select Case lMessageType
                    Case TLS_HANDSHAKE_TYPE_SERVER_HELLO
                        lPos = pvReadArray(baInput, lPos, baMessage, lMessageSize)
                        If Not pvParseHandshakeServerHello(uCtx, baMessage, lRecordProtocol, sError, eAlertCode) Then
                            GoTo QH
                        End If
                        If pvArraySize(.HelloRetryCookie) <> 0 Then
                            '--- after HelloRetryRequest -> replace HandshakeMessages w/ "synthetic handshake message"
                            baHandshakeHash = pvCryptoHash(.DigestAlgo, .HandshakeMessages, 0)
                            Erase .HandshakeMessages
                            lVerifyPos = pvWriteLong(.HandshakeMessages, 0, TLS_HANDSHAKE_TYPE_MESSAGE_HASH)
                            lVerifyPos = pvWriteLong(.HandshakeMessages, lVerifyPos, .DigestSize, Size:=3)
                            lVerifyPos = pvWriteArray(.HandshakeMessages, lVerifyPos, baHandshakeHash)
                        Else
                            .State = ucsTlsStateExpectExtensions
                        End If
                    Case Else
                        sError = Replace(Replace(ERR_UNEXPECTED_MSG_TYPE, "%1", "ucsTlsStateExpectServerHello"), "%2", lMessageType)
                        eAlertCode = uscTlsAlertUnexpectedMessage
                        GoTo QH
                    End Select
                    pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baInput(lMessagePos)), lMessageSize + 4
                    '--- post-process ucsTlsStateExpectServerHello
                    If .State = ucsTlsStateExpectServerHello And pvArraySize(.HelloRetryCookie) <> 0 Then
                        .SendPos = pvBuildClientHello(uCtx, .SendBuffer, .SendPos)
                    End If
                    If .State = ucsTlsStateExpectExtensions And .ServerProtocol = TLS_PROTOCOL_VERSION_TLS13_FINAL Then
                        If Not pvDeriveHandshakeSecrets(uCtx, sError, eAlertCode) Then
                            GoTo QH
                        End If
                    End If
                Case ucsTlsStateExpectExtensions
                    Select Case lMessageType
                    Case TLS_HANDSHAKE_TYPE_CERTIFICATE
                        If .ServerProtocol = TLS_PROTOCOL_VERSION_TLS13_FINAL Then
                            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, BlockSize:=lCertSize)
                                lPos = pvReadArray(baInput, lPos, .ServerCertReqContext, lCertSize)
                            lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                        End If
                        Set .ServerCertificates = New Collection
                        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=3, BlockSize:=lCertSize)
                            lCertEnd = lPos + lCertSize
                            Do While lPos < lCertEnd
                                lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=3, BlockSize:=lCertSize)
                                    lPos = pvReadArray(baInput, lPos, baCert, lCertSize)
                                    .ServerCertificates.Add baCert
                                lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                                If .ServerProtocol = TLS_PROTOCOL_VERSION_TLS13_FINAL Then
                                    lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lCertSize)
                                        '--- certificate extensions -> skip
                                        lPos = lPos + lCertSize
                                    lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                                End If
                            Loop
                        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                    Case TLS_HANDSHAKE_TYPE_CERTIFICATE_VERIFY
                        baHandshakeHash = pvCryptoHash(.DigestAlgo, .HandshakeMessages, 0)
                        lVerifyPos = pvWriteString(baVerifyData, 0, Space$(64) & "TLS 1.3, server CertificateVerify" & Chr$(0))
                        lVerifyPos = pvWriteArray(baVerifyData, lVerifyPos, baHandshakeHash)
                        '--- ToDo: verify .ServerCertificate signature
                        '--- ShellExecute("openssl x509 -pubkey -noout -in server.crt > server.pub")
                        lPos = lPos + lMessageSize
                    Case TLS_HANDSHAKE_TYPE_FINISHED
                        lPos = pvReadArray(baInput, lPos, baMessage, lMessageSize)
                        baHandshakeHash = pvCryptoHash(.DigestAlgo, .HandshakeMessages, 0)
                        baVerifyData = pvHkdfExpandLabel(.DigestAlgo, .ServerTrafficSecret, "finished", EmptyByteArray, .DigestSize)
                        baVerifyData = pvHkdfExtract(.DigestAlgo, baVerifyData, baHandshakeHash)
                        Debug.Assert StrConv(baVerifyData, vbUnicode) = StrConv(baMessage, vbUnicode)
                        If StrConv(baVerifyData, vbUnicode) <> StrConv(baMessage, vbUnicode) Then
                            sError = ERR_SERVER_HANDSHAKE_FAILED
                            eAlertCode = uscTlsAlertHandshakeFailure
                            GoTo QH
                        End If
                        .State = ucsTlsStatePostHandshake
                    Case TLS_HANDSHAKE_TYPE_SERVER_KEY_EXCHANGE
                        If .ServerProtocol = TLS_PROTOCOL_VERSION_TLS12 Then
                            lPos = pvReadLong(baInput, lPos, lCurveType)
                            Debug.Assert lCurveType = 3 '--- 3 = named_curve
                            lPos = pvReadLong(baInput, lPos, lNamedCurve, Size:=2)
                            If Not pvSetupKeyExchangeEccGroup(uCtx, lNamedCurve, sError, eAlertCode) Then
                                GoTo QH
                            End If
                            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, BlockSize:=lSignatureSize)
                                lPos = pvReadArray(baInput, lPos, .ServerPublic, lSignatureSize)
                            lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                            lPos = pvReadLong(baInput, lPos, lSignatureType, Size:=2)
                            '-- &H401 = RSA signature with SHA256 hash
                            Debug.Print "Using signature type &H" & Hex$(lSignatureType), Timer
                            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lSignatureSize)
                                lPos = pvReadArray(baInput, lPos, baSignature, lSignatureSize)
                                '--- ToDo: verify .ServerCertificate signature
                            lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                            If Not pvDeriveLegacySecrets(uCtx, sError, eAlertCode) Then
                                GoTo QH
                            End If
                        End If
                    Case TLS_HANDSHAKE_TYPE_SERVER_HELLO_DONE
                        If .ServerProtocol = TLS_PROTOCOL_VERSION_TLS12 Then
                            .State = ucsTlsStateExpectServerFinish
                        End If
                        lPos = lPos + lMessageSize
                    Case Else
                        '--- do nothing
                        lPos = lPos + lMessageSize
                    End Select
                    pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baInput(lMessagePos)), lMessageSize + 4
                    '--- post-process ucsTlsStateExpectExtensions
                    If .State = ucsTlsStateExpectServerFinish And .ServerProtocol = TLS_PROTOCOL_VERSION_TLS12 Then
                        If pvCryptoCipherSuiteUseRsaCertificate(.CipherSuite) Then
                            On Error Resume Next
                            baCert = .ServerCertificates.Item(1)
                            On Error GoTo 0
                            If pvArraySize(baCert) = 0 Then
                                sError = ERR_NO_SERVER_CERTIFICATE
                                eAlertCode = uscTlsAlertHandshakeFailure
                                GoTo QH
                            End If
                            If Not pvSetupKeyExchangeRsaCertificate(uCtx, baCert, sError, eAlertCode) Then
                                GoTo QH
                            End If
                            If Not pvDeriveLegacySecrets(uCtx, sError, eAlertCode) Then
                                GoTo QH
                            End If
                        End If
                        .SendPos = pvBuildLegacyClientKeyExchange(uCtx, .SendBuffer, .SendPos, sError, eAlertCode)
                        If LenB(sError) <> 0 Then
                            GoTo QH
                        End If
                    End If
                    If .State = ucsTlsStatePostHandshake And .ServerProtocol = TLS_PROTOCOL_VERSION_TLS13_FINAL Then
                        .SendPos = pvBuildClientHandshakeFinished(uCtx, .SendBuffer, .SendPos, sError, eAlertCode)
                        If LenB(sError) <> 0 Then
                            GoTo QH
                        End If
                        If Not pvDeriveApplicationSecrets(uCtx, sError, eAlertCode) Then
                            GoTo QH
                        End If
                        '--- not used past handshake
                        Erase .HandshakeMessages
                    End If
                Case ucsTlsStateExpectServerFinish
                    Select Case lMessageType
                    Case TLS_HANDSHAKE_TYPE_FINISHED
                        If .ServerProtocol = TLS_PROTOCOL_VERSION_TLS12 Then
                            lPos = pvReadArray(baInput, lPos, baMessage, lMessageSize)
                            baHandshakeHash = pvCryptoHash(.DigestAlgo, .HandshakeMessages, 0)
                            baVerifyData = pvKdfLegacyTls1Prf(.DigestAlgo, .MasterSecret, "server finished", baHandshakeHash, 12)
                            Debug.Assert StrConv(baVerifyData, vbUnicode) = StrConv(baMessage, vbUnicode)
                            If StrConv(baVerifyData, vbUnicode) <> StrConv(baMessage, vbUnicode) Then
                                sError = ERR_SERVER_HANDSHAKE_FAILED
                                eAlertCode = uscTlsAlertHandshakeFailure
                                GoTo QH
                            End If
                            .State = ucsTlsStatePostHandshake
                            '--- not used past handshake
                            Erase .HandshakeMessages
                        Else
                            GoTo InvalidState
                        End If
                    Case Else
                        sError = Replace(Replace(ERR_UNEXPECTED_MSG_TYPE, "%1", "ucsTlsStateExpectServerFinish"), "%2", lMessageType)
                        eAlertCode = uscTlsAlertUnexpectedMessage
                        GoTo QH
                    End Select
                Case ucsTlsStatePostHandshake
                    Select Case lMessageType
                    Case TLS_HANDSHAKE_TYPE_NEW_SESSION_TICKET
                        '--- don't store tickets for now
                    Case TLS_HANDSHAKE_TYPE_KEY_UPDATE
                        Debug.Print "Received TLS_HANDSHAKE_TYPE_KEY_UPDATE", Timer
                        If lMessageSize = 1 Then
                            lRequestUpdate = baInput(lPos)
                        Else
                            lRequestUpdate = -1
                        End If
                        If Not pvDeriveKeyUpdate(uCtx, lRequestUpdate <> 0, sError, eAlertCode) Then
                            GoTo QH
                        End If
                        If lRequestUpdate <> 0 Then
                            '--- ack by TLS_HANDSHAKE_TYPE_KEY_UPDATE w/ update_not_requested(0)
                            If pvBuildClientApplicationData(uCtx, baMessage, 0, FromHex("18:00:00:01:00"), -1, sError, eAlertCode) = 0 Then
                                GoTo QH
                            End If
                            .SendPos = pvWriteArray(.SendBuffer, .SendPos, baMessage)
                        End If
                    Case Else
                        sError = Replace(Replace(ERR_UNEXPECTED_MSG_TYPE, "%1", "ucsTlsStatePostHandshake"), "%2", lMessageType)
                        eAlertCode = uscTlsAlertUnexpectedMessage
                        GoTo QH
                    End Select
                    lPos = lPos + lMessageSize
                Case Else
InvalidState:
                    sError = Replace(ERR_INVALID_STATE_HANDSHAKE, "%1", .State)
                    eAlertCode = uscTlsAlertHandshakeFailure
                    GoTo QH
                End Select
            lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
        Loop
    End With
    pvParseHandshakeContent = lPos
QH:
End Function

Private Function pvParseHandshakeServerHello(uCtx As UcsTlsContext, baMessage() As Byte, ByVal lRecordProtocol As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Static baHelloRetryRandom() As Byte
    Dim lPos            As Long
    Dim lBlockSize      As Long
    Dim lLegacyVersion  As Long
    Dim lCipherSuite    As Long
    Dim lLegacyCompress As Long
    Dim lExtType        As Long
    Dim lExchangeGroup  As Long
    Dim lEnd            As Long
    Dim bHelloRetryRequest As Boolean
    
    If pvArraySize(baHelloRetryRandom) = 0 Then
        baHelloRetryRandom = FromHex(STR_HELLO_RETRY_RANDOM)
    End If
    With uCtx
        '--- clear HelloRetryRequest
        .HelloRetryCipherSuite = 0
        .HelloRetryExchangeGroup = 0
        Erase .HelloRetryCookie
        .ServerProtocol = lRecordProtocol
        lPos = pvReadLong(baMessage, lPos, lLegacyVersion, Size:=2)
        lPos = pvReadArray(baMessage, lPos, .ServerRandom, .RandomSize)
        bHelloRetryRequest = (StrConv(.ServerRandom, vbUnicode) = StrConv(baHelloRetryRandom, vbUnicode))
        lPos = pvReadBeginOfBlock(baMessage, lPos, .BlocksStack, BlockSize:=lBlockSize)
            lPos = pvReadArray(baMessage, lPos, .ServerSessionID, lBlockSize)
        lPos = pvReadEndOfBlock(baMessage, lPos, .BlocksStack)
        lPos = pvReadLong(baMessage, lPos, lCipherSuite, Size:=2)
        If Not pvSetupCipherSuite(uCtx, lCipherSuite, sError, eAlertCode) Then
            GoTo QH
        End If
        Debug.Print "Using " & pvCryptoCipherSuiteName(.CipherSuite) & " from " & .TargetHost, Timer
        If bHelloRetryRequest Then
            .HelloRetryCipherSuite = lCipherSuite
        End If
        lPos = pvReadLong(baMessage, lPos, lLegacyCompress)
        Debug.Assert lLegacyCompress = 0
        lPos = pvReadBeginOfBlock(baMessage, lPos, .BlocksStack, Size:=2, BlockSize:=lBlockSize)
            lEnd = lPos + lBlockSize
            Do While lPos < lEnd
                lPos = pvReadLong(baMessage, lPos, lExtType, Size:=2)
                lPos = pvReadBeginOfBlock(baMessage, lPos, .BlocksStack, Size:=2, BlockSize:=lBlockSize)
                    Select Case lExtType
                    Case TLS_EXTENSION_TYPE_KEY_SHARE
                        .ServerProtocol = TLS_PROTOCOL_VERSION_TLS13_FINAL
                        If lBlockSize >= 2 Then
                            lPos = pvReadLong(baMessage, lPos, lExchangeGroup, Size:=2)
                        Else
                            sError = ERR_INVALID_SIZE_KEY_SHARE
                            eAlertCode = uscTlsAlertDecodeError
                            GoTo QH
                        End If
                        If bHelloRetryRequest Then
                            .HelloRetryExchangeGroup = lExchangeGroup
                        Else
                            Debug.Assert lExchangeGroup = TLS_GROUP_X25519
                            If Not pvSetupKeyExchangeEccGroup(uCtx, lExchangeGroup, sError, eAlertCode) Then
                                GoTo QH
                            End If
                            If lBlockSize > 4 Then
                                lPos = pvReadBeginOfBlock(baMessage, lPos, .BlocksStack, Size:=2, BlockSize:=lBlockSize)
                                    Debug.Assert lBlockSize = TLS_X25519_KEY_SIZE
                                    If lBlockSize <> TLS_X25519_KEY_SIZE Then
                                        sError = ERR_INVALID_SERVER_KEY
                                        eAlertCode = uscTlsAlertIllegalParameter
                                        GoTo QH
                                    End If
                                    lPos = pvReadArray(baMessage, lPos, .ServerPublic, lBlockSize)
                                lPos = pvReadEndOfBlock(baMessage, lPos, .BlocksStack)
                            Else
                                sError = ERR_INVALID_SIZE_SERVER_KEY
                                eAlertCode = uscTlsAlertDecodeError
                                GoTo QH
                            End If
                        End If
                    Case TLS_EXTENSION_TYPE_SUPPORTED_VERSIONS
                        If lBlockSize <> 2 Then
                            sError = ERR_INVALID_SIZE_VERSIONS
                            eAlertCode = uscTlsAlertDecodeError
                            GoTo QH
                        End If
                        lPos = pvReadLong(baMessage, lPos, .ServerProtocol, Size:=2)
                    Case TLS_EXTENSION_TYPE_COOKIE
                        If bHelloRetryRequest Then
                            lPos = pvReadArray(baMessage, lPos, .HelloRetryCookie, lBlockSize)
                        Else
                            sError = ERR_COOKIE_NOT_ALLOWED
                            eAlertCode = uscTlsAlertIllegalParameter
                            GoTo QH
                        End If
                    Case Else
                        lPos = lPos + lBlockSize
                    End Select
                lPos = pvReadEndOfBlock(baMessage, lPos, .BlocksStack)
            Loop
        lPos = pvReadEndOfBlock(baMessage, lPos, .BlocksStack)
    End With
    '--- success
    pvParseHandshakeServerHello = True
QH:
End Function

Private Sub pvPrepareLegacyDecryptParams(uCtx As UcsTlsContext, baInput() As Byte, ByVal lRecordPos As Long, ByVal lRecordSize As Long, lPos As Long, lEnd As Long, baServerIV() As Byte, baAad() As Byte)
    Dim lAadPos         As Long
    
    With uCtx
        lEnd = lPos + lRecordSize - .TagSize
        If .IvDynamicSize > 0 Then '--- AES in TLS 1.2
            ReDim baServerIV(0 To .IvSize - 1) As Byte
            pvWriteArray baServerIV, 0, .ServerTrafficIV
            pvWriteBuffer baServerIV, .IvSize - .IvDynamicSize, VarPtr(baInput(lPos)), .IvDynamicSize
            lPos = lPos + .IvDynamicSize
        Else
            baServerIV = pvArrayXor(.ServerTrafficIV, .ServerTrafficSeqNo)
        End If
        ReDim baAad(0 To LNG_LEGACY_AAD_SIZE - 1) As Byte
        lAadPos = pvWriteLong(baAad, 0, 0, Size:=4)
        lAadPos = pvWriteLong(baAad, lAadPos, .ServerTrafficSeqNo, Size:=4)
        lAadPos = pvWriteBuffer(baAad, lAadPos, VarPtr(baInput(lRecordPos)), 3)
        lAadPos = pvWriteLong(baAad, lAadPos, lEnd - lPos, Size:=2)
        Debug.Assert lAadPos = LNG_LEGACY_AAD_SIZE
    End With
End Sub

Private Sub pvSetLastError(uCtx As UcsTlsContext, sError As String, Optional ByVal AlertDesc As UcsTlsAlertDescriptionsEnum = -1)
    With uCtx
        .LastError = sError
        .LastAlertCode = AlertDesc
        If LenB(sError) = 0 Then
            Set .BlocksStack = Nothing
        Else
            If AlertDesc >= 0 Then
                .SendPos = pvBuildClientAlert(uCtx, .SendBuffer, .SendPos, AlertDesc, TLS_ALERT_LEVEL_FATAL)
            End If
            .State = ucsTlsStateClosed
        End If
    End With
End Sub

'= HMAC-based key derivation functions ===================================

Private Function pvDeriveHandshakeSecrets(uCtx As UcsTlsContext, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Dim baHandshakeHash() As Byte
    Dim baEarlySecret() As Byte
    Dim baEmptyHash()   As Byte
    Dim baDerivedSecret() As Byte
    Dim baSharedSecret() As Byte
    
    With uCtx
        If pvArraySize(.HandshakeMessages) = 0 Then
            sError = ERR_NO_HANDSHAKE_MESSAGES
            eAlertCode = uscTlsAlertInternalError
            GoTo QH
        End If
        baHandshakeHash = pvCryptoHash(.DigestAlgo, .HandshakeMessages, 0)
        baEarlySecret = pvHkdfExtract(.DigestAlgo, EmptyByteArray(.DigestSize), EmptyByteArray(.DigestSize))
        baEmptyHash = pvCryptoHash(.DigestAlgo, EmptyByteArray, 0)
        baDerivedSecret = pvHkdfExpandLabel(.DigestAlgo, baEarlySecret, "derived", baEmptyHash, .DigestSize)
        baSharedSecret = pvCryptoSharedSecret(.ExchangeAlgo, .ClientPrivate, .ServerPublic)
        .HandshakeSecret = pvHkdfExtract(.DigestAlgo, baDerivedSecret, baSharedSecret)
        .ServerTrafficSecret = pvHkdfExpandLabel(.DigestAlgo, .HandshakeSecret, "s hs traffic", baHandshakeHash, .DigestSize)
        .ServerTrafficKey = pvHkdfExpandLabel(.DigestAlgo, .ServerTrafficSecret, "key", EmptyByteArray, .KeySize)
        .ServerTrafficIV = pvHkdfExpandLabel(.DigestAlgo, .ServerTrafficSecret, "iv", EmptyByteArray, .IvSize)
        .ServerTrafficSeqNo = 0
        .ClientTrafficSecret = pvHkdfExpandLabel(.DigestAlgo, .HandshakeSecret, "c hs traffic", baHandshakeHash, .DigestSize)
        .ClientTrafficKey = pvHkdfExpandLabel(.DigestAlgo, .ClientTrafficSecret, "key", EmptyByteArray, .KeySize)
        .ClientTrafficIV = pvHkdfExpandLabel(.DigestAlgo, .ClientTrafficSecret, "iv", EmptyByteArray, .IvSize)
        .ClientTrafficSeqNo = 0
    End With
    '--- success
    pvDeriveHandshakeSecrets = True
QH:
End Function

Private Function pvDeriveApplicationSecrets(uCtx As UcsTlsContext, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Dim baHandshakeHash() As Byte
    Dim baEmptyHash()   As Byte
    Dim baDerivedSecret() As Byte
    
    With uCtx
        If pvArraySize(.HandshakeMessages) = 0 Then
            sError = ERR_NO_HANDSHAKE_MESSAGES
            eAlertCode = uscTlsAlertInternalError
            GoTo QH
        End If
        baHandshakeHash = pvCryptoHash(.DigestAlgo, .HandshakeMessages, 0)
        baEmptyHash = pvCryptoHash(.DigestAlgo, EmptyByteArray, 0)
        baDerivedSecret = pvHkdfExpandLabel(.DigestAlgo, .HandshakeSecret, "derived", baEmptyHash, .DigestSize)
        .MasterSecret = pvHkdfExtract(.DigestAlgo, baDerivedSecret, EmptyByteArray(.DigestSize))
        .ServerTrafficSecret = pvHkdfExpandLabel(.DigestAlgo, .MasterSecret, "s ap traffic", baHandshakeHash, .DigestSize)
        .ServerTrafficKey = pvHkdfExpandLabel(.DigestAlgo, .ServerTrafficSecret, "key", EmptyByteArray, .KeySize)
        .ServerTrafficIV = pvHkdfExpandLabel(.DigestAlgo, .ServerTrafficSecret, "iv", EmptyByteArray, .IvSize)
        .ServerTrafficSeqNo = 0
        .ClientTrafficSecret = pvHkdfExpandLabel(.DigestAlgo, .MasterSecret, "c ap traffic", baHandshakeHash, .DigestSize)
        .ClientTrafficKey = pvHkdfExpandLabel(.DigestAlgo, .ClientTrafficSecret, "key", EmptyByteArray, .KeySize)
        .ClientTrafficIV = pvHkdfExpandLabel(.DigestAlgo, .ClientTrafficSecret, "iv", EmptyByteArray, .IvSize)
        .ClientTrafficSeqNo = 0
    End With
    '--- success
    pvDeriveApplicationSecrets = True
QH:
End Function

Private Function pvDeriveKeyUpdate(uCtx As UcsTlsContext, ByVal bUpdateClient As Boolean, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    With uCtx
        If pvArraySize(.ServerTrafficSecret) = 0 Then
            sError = ERR_NO_PREV_SERVER_SECRET
            eAlertCode = uscTlsAlertInternalError
            GoTo QH
        End If
        .ServerTrafficSecret = pvHkdfExpandLabel(.DigestAlgo, .ServerTrafficSecret, "traffic upd", EmptyByteArray, .DigestSize)
        .ServerTrafficKey = pvHkdfExpandLabel(.DigestAlgo, .ServerTrafficSecret, "key", EmptyByteArray, .KeySize)
        .ServerTrafficIV = pvHkdfExpandLabel(.DigestAlgo, .ServerTrafficSecret, "iv", EmptyByteArray, .IvSize)
        .ServerTrafficSeqNo = 0
        If bUpdateClient Then
            If pvArraySize(.ClientTrafficSecret) = 0 Then
                sError = ERR_NO_PREV_CLIENT_SECRET
                eAlertCode = uscTlsAlertInternalError
                GoTo QH
            End If
            .ClientTrafficSecret = pvHkdfExpandLabel(.DigestAlgo, .ClientTrafficSecret, "traffic upd", EmptyByteArray, .DigestSize)
            .ClientTrafficKey = pvHkdfExpandLabel(.DigestAlgo, .ClientTrafficSecret, "key", EmptyByteArray, .KeySize)
            .ClientTrafficIV = pvHkdfExpandLabel(.DigestAlgo, .ClientTrafficSecret, "iv", EmptyByteArray, .IvSize)
            .ClientTrafficSeqNo = 0
        End If
    End With
    '--- success
    pvDeriveKeyUpdate = True
QH:
End Function

Private Function pvHkdfExtract(ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baKey() As Byte, baInput() As Byte) As Byte()
    pvHkdfExtract = pvCryptoHmac(eHash, baKey, baInput, 0)
End Function

Private Function pvHkdfExpandLabel(ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baKey() As Byte, ByVal sLabel As String, baContext() As Byte, ByVal lSize As Long) As Byte()
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
        baLast = pvCryptoHmac(eHash, baKey, baInput, 0, Size:=lInputPos)
        lRetValPos = pvWriteArray(baRetVal, lRetValPos, baLast)
        lIdx = lIdx + 1
    Loop
    If UBound(baRetVal) <> lSize - 1 Then
        ReDim Preserve baRetVal(0 To lSize - 1) As Byte
    End If
    pvHkdfExpandLabel = baRetVal
End Function

'= legacy PRF-based key derivation functions =============================

Private Function pvDeriveLegacySecrets(uCtx As UcsTlsContext, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Dim baPreMasterSecret() As Byte
    Dim baRandom()      As Byte
    Dim baExpanded()    As Byte
    Dim lPos            As Long
    
    With uCtx
        If pvArraySize(.ServerRandom) = 0 Then
            sError = ERR_NO_SERVER_RANDOM
            eAlertCode = uscTlsAlertInternalError
            GoTo QH
        End If
        Debug.Assert pvArraySize(.ClientRandom) = TLS_HELLO_RANDOM_SIZE
        Debug.Assert pvArraySize(.ServerRandom) = TLS_HELLO_RANDOM_SIZE
        baPreMasterSecret = pvCryptoSharedSecret(.ExchangeAlgo, .ClientPrivate, .ServerPublic)
        ReDim baRandom(0 To pvArraySize(.ClientRandom) + pvArraySize(.ServerRandom) - 1) As Byte
        lPos = pvWriteArray(baRandom, 0, .ClientRandom)
        lPos = pvWriteArray(baRandom, lPos, .ServerRandom)
        Debug.Assert lPos = UBound(baRandom) + 1
        .MasterSecret = pvKdfLegacyTls1Prf(.DigestAlgo, baPreMasterSecret, "master secret", baRandom, .RandomSize + .RandomSize \ 2) '--- always 48
        lPos = pvWriteArray(baRandom, 0, .ServerRandom)
        lPos = pvWriteArray(baRandom, lPos, .ClientRandom)
        Debug.Assert lPos = UBound(baRandom) + 1
        baExpanded = pvKdfLegacyTls1Prf(.DigestAlgo, .MasterSecret, "key expansion", baRandom, 2 * (.MacSize + .KeySize + .IvSize))
        lPos = pvReadArray(baExpanded, 0, EmptyByteArray, .MacSize) '--- ClientMacKey not used w/ AEAD
        lPos = pvReadArray(baExpanded, lPos, EmptyByteArray, .MacSize) '--- ServerMacKey not used w/ AEAD
        lPos = pvReadArray(baExpanded, lPos, .ClientTrafficKey, .KeySize)
        lPos = pvReadArray(baExpanded, lPos, .ServerTrafficKey, .KeySize)
        lPos = pvReadArray(baExpanded, lPos, .ClientTrafficIV, .IvSize - .IvDynamicSize)
        lPos = pvReadArray(baExpanded, lPos, .ServerTrafficIV, .IvSize - .IvDynamicSize)
    End With
    '--- success
    pvDeriveLegacySecrets = True
QH:
End Function

Private Function pvKdfLegacyTls1Prf(ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baSecret() As Byte, ByVal sLabel As String, baContext() As Byte, ByVal lSize As Long) As Byte()
    Dim baSeed()        As Byte
    Dim baRetVal()      As Byte
    Dim lRetValPos      As Long
    Dim baInput()       As Byte
    Dim lInputPos       As Long
    Dim baLast()        As Byte
    Dim baHmac()        As Byte
    
    lInputPos = pvWriteString(baSeed, 0, sLabel)
    lInputPos = pvWriteArray(baSeed, lInputPos, baContext)
    baLast = baSeed
    Do While lRetValPos < lSize
        baLast = pvCryptoHmac(eHash, baSecret, baLast, 0)
        lInputPos = pvWriteArray(baInput, 0, baLast)
        lInputPos = pvWriteArray(baInput, lInputPos, baSeed)
        baHmac = pvCryptoHmac(eHash, baSecret, baInput, 0, Size:=lInputPos)
        lRetValPos = pvWriteArray(baRetVal, lRetValPos, baHmac)
    Loop
    If lRetValPos <> lSize Then
        ReDim Preserve baRetVal(0 To lSize - 1) As Byte
    End If
    pvKdfLegacyTls1Prf = baRetVal
End Function

'= crypto wrappers =======================================================

Private Function pvCryptoOrderCiphers(ByVal eFilter As UcsTlsClientFeaturesEnum) As Collection
    Const PREF      As Long = &H1000
    Dim oRetVal     As Collection
    
    Set oRetVal = New Collection
    If (eFilter And ucsTlsSupportTls13) <> 0 Then
        If CryptoIsSupported(ucsTlsAlgoKeyX25519) Then
            '--- first if AES preferred over Chacha20
            If CryptoIsSupported(PREF + ucsTlsAlgoAeadAes128) And CryptoIsSupported(ucsTlsAlgoAeadAes128) Then
                oRetVal.Add TLS_CIPHER_SUITE_AES_128_GCM_SHA256
            End If
            If CryptoIsSupported(PREF + ucsTlsAlgoAeadAes256) And CryptoIsSupported(ucsTlsAlgoAeadAes256) Then
                oRetVal.Add TLS_CIPHER_SUITE_AES_256_GCM_SHA384
            End If
            If CryptoIsSupported(ucsTlsAlgoAeadChacha20Poly1305) Then
                oRetVal.Add TLS_CIPHER_SUITE_CHACHA20_POLY1305_SHA256
            End If
            '--- least preferred AES
            If Not CryptoIsSupported(PREF + ucsTlsAlgoAeadAes128) And CryptoIsSupported(ucsTlsAlgoAeadAes128) Then
                oRetVal.Add TLS_CIPHER_SUITE_AES_128_GCM_SHA256
            End If
            If Not CryptoIsSupported(PREF + ucsTlsAlgoAeadAes256) And CryptoIsSupported(ucsTlsAlgoAeadAes256) Then
                oRetVal.Add TLS_CIPHER_SUITE_AES_256_GCM_SHA384
            End If
        End If
    End If
    If (eFilter And ucsTlsSupportTls12) <> 0 Then
        If CryptoIsSupported(ucsTlsAlgoKeySecp256r1) Then
            '--- first if AES preferred over Chacha20
            If CryptoIsSupported(PREF + ucsTlsAlgoAeadAes128) And CryptoIsSupported(ucsTlsAlgoAeadAes128) Then
                oRetVal.Add TLS_CIPHER_SUITE_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256
                oRetVal.Add TLS_CIPHER_SUITE_ECDHE_RSA_WITH_AES_128_GCM_SHA256
            End If
            If CryptoIsSupported(PREF + ucsTlsAlgoAeadAes128) And CryptoIsSupported(ucsTlsAlgoAeadAes256) Then
                oRetVal.Add TLS_CIPHER_SUITE_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384
                oRetVal.Add TLS_CIPHER_SUITE_ECDHE_RSA_WITH_AES_256_GCM_SHA384
            End If
            If CryptoIsSupported(ucsTlsAlgoAeadChacha20Poly1305) Then
                oRetVal.Add TLS_CIPHER_SUITE_ECDHE_ECDSA_WITH_CHACHA20_POLY1305_SHA256
                oRetVal.Add TLS_CIPHER_SUITE_ECDHE_RSA_WITH_CHACHA20_POLY1305_SHA256
            End If
            '--- least preferred AES
            If Not CryptoIsSupported(PREF + ucsTlsAlgoAeadAes128) And CryptoIsSupported(ucsTlsAlgoAeadAes128) Then
                oRetVal.Add TLS_CIPHER_SUITE_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256
                oRetVal.Add TLS_CIPHER_SUITE_ECDHE_RSA_WITH_AES_128_GCM_SHA256
            End If
            If Not CryptoIsSupported(PREF + ucsTlsAlgoAeadAes128) And CryptoIsSupported(ucsTlsAlgoAeadAes256) Then
                oRetVal.Add TLS_CIPHER_SUITE_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384
                oRetVal.Add TLS_CIPHER_SUITE_ECDHE_RSA_WITH_AES_256_GCM_SHA384
            End If
        End If
        '--- no "perfect forward secrecy" -> least preferred
        If CryptoIsSupported(ucsTlsAlgoAeadAes128) Then
            oRetVal.Add TLS_CIPHER_SUITE_RSA_WITH_AES_128_GCM_SHA256
        End If
        If CryptoIsSupported(ucsTlsAlgoAeadAes256) Then
            oRetVal.Add TLS_CIPHER_SUITE_RSA_WITH_AES_256_GCM_SHA384
        End If
    End If
    Set pvCryptoOrderCiphers = oRetVal
End Function

Private Function pvCryptoAeadDecrypt(ByVal eAead As UcsTlsCryptoAlgorithmsEnum, baServerIV() As Byte, baServerKey() As Byte, baAad() As Byte, ByVal lAadPos As Long, ByVal lAdSize As Long, baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Select Case eAead
    Case ucsTlsAlgoAeadChacha20Poly1305
        If Not CryptoAeadChacha20Poly1305Decrypt(baServerIV, baServerKey, baAad, lAadPos, lAdSize, baBuffer, lPos, lSize) Then
            GoTo QH
        End If
    Case ucsTlsAlgoAeadAes128, ucsTlsAlgoAeadAes256
        If Not CryptoAeadAesGcmDecrypt(baServerIV, baServerKey, baAad, lAadPos, lAdSize, baBuffer, lPos, lSize) Then
            GoTo QH
        End If
    Case Else
        Err.Raise vbObjectError, "pvCryptoAeadDecrypt", "Unsupported AEAD type " & eAead
    End Select
    '--- success
    pvCryptoAeadDecrypt = True
QH:
End Function

Private Function pvCryptoAeadEncrypt(ByVal eAead As UcsTlsCryptoAlgorithmsEnum, baClientIV() As Byte, baClientKey() As Byte, baAad() As Byte, ByVal lAadPos As Long, ByVal lAdSize As Long, baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Select Case eAead
    Case ucsTlsAlgoAeadChacha20Poly1305
        If Not CryptoAeadChacha20Poly1305Encrypt(baClientIV, baClientKey, baAad, lAadPos, lAdSize, baBuffer, lPos, lSize) Then
            GoTo QH
        End If
    Case ucsTlsAlgoAeadAes128, ucsTlsAlgoAeadAes256
        If Not CryptoAeadAesGcmEncrypt(baClientIV, baClientKey, baAad, lAadPos, lAdSize, baBuffer, lPos, lSize) Then
            GoTo QH
        End If
    Case Else
        Err.Raise vbObjectError, "pvCryptoAeadEncrypt", "Unsupported AEAD type " & eAead
    End Select
    '--- success
    pvCryptoAeadEncrypt = True
QH:
End Function

Private Function pvCryptoHash(ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baInput() As Byte, ByVal lPos As Long, Optional ByVal Size As Long = -1) As Byte()
    Select Case eHash
    Case ucsTlsAlgoDigestSha256
        pvCryptoHash = CryptoHashSha256(baInput, lPos, Size)
    Case ucsTlsAlgoDigestSha384
        pvCryptoHash = CryptoHashSha384(baInput, lPos, Size)
    Case Else
        Err.Raise vbObjectError, "pvCryptoHash", "Unsupported hash type " & eHash
    End Select
End Function

Private Function pvCryptoHmac(ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baKey() As Byte, baInput() As Byte, ByVal lPos As Long, Optional ByVal Size As Long = -1) As Byte()
    Select Case eHash
    Case ucsTlsAlgoDigestSha256
        pvCryptoHmac = CryptoHmacSha256(baKey, baInput, lPos, Size)
    Case ucsTlsAlgoDigestSha384
        pvCryptoHmac = CryptoHmacSha384(baKey, baInput, lPos, Size)
    Case Else
        Err.Raise vbObjectError, "pvCryptoHmac", "Unsupported hash type " & eHash
    End Select
End Function

Private Function pvCryptoSharedSecret(ByVal eKeyX As UcsTlsCryptoAlgorithmsEnum, baPriv() As Byte, baPub() As Byte) As Byte()
    Dim baRetVal()      As Byte
    
    Select Case eKeyX
    Case ucsTlsAlgoKeyX25519
        baRetVal = CryptoEccCurve25519SharedSecret(baPriv, baPub)
    Case ucsTlsAlgoKeySecp256r1
        baRetVal = CryptoEccSecp256r1SharedSecret(baPriv, baPub)
    Case ucsTlsAlgoKeyCertificate
        baRetVal = baPriv
    Case Else
        Err.Raise vbObjectError, "pvCryptoSharedSecret", "Unsupported exchange curve " & eKeyX
    End Select
    pvCryptoSharedSecret = baRetVal
End Function

Private Function pvCryptoRandomArray(ByVal lSize As Long) As Byte()
    Dim baRetVal()      As Byte
    
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
        CryptoRandomBytes VarPtr(baRetVal(0)), lSize
    End If
    pvCryptoRandomArray = baRetVal
End Function

Private Function pvCryptoCipherSuiteName(ByVal lCipherSuite As Long) As String
    Select Case lCipherSuite
    Case TLS_CIPHER_SUITE_AES_128_GCM_SHA256
        pvCryptoCipherSuiteName = "TLS_AES_128_GCM_SHA256"
    Case TLS_CIPHER_SUITE_AES_256_GCM_SHA384
        pvCryptoCipherSuiteName = "TLS_AES_256_GCM_SHA384"
    Case TLS_CIPHER_SUITE_CHACHA20_POLY1305_SHA256
        pvCryptoCipherSuiteName = "TLS_CHACHA20_POLY1305_SHA256"
    Case TLS_CIPHER_SUITE_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256
        pvCryptoCipherSuiteName = "ECDHE-ECDSA-AES128-GCM-SHA256"
    Case TLS_CIPHER_SUITE_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384
        pvCryptoCipherSuiteName = "ECDHE-ECDSA-AES256-GCM-SHA384"
    Case TLS_CIPHER_SUITE_ECDHE_RSA_WITH_AES_128_GCM_SHA256
        pvCryptoCipherSuiteName = "ECDHE-RSA-AES128-GCM-SHA256"
    Case TLS_CIPHER_SUITE_ECDHE_RSA_WITH_AES_256_GCM_SHA384
        pvCryptoCipherSuiteName = "ECDHE-RSA-AES256-GCM-SHA384"
    Case TLS_CIPHER_SUITE_ECDHE_RSA_WITH_CHACHA20_POLY1305_SHA256
        pvCryptoCipherSuiteName = "ECDHE-RSA-CHACHA20-POLY1305"
    Case TLS_CIPHER_SUITE_ECDHE_ECDSA_WITH_CHACHA20_POLY1305_SHA256
        pvCryptoCipherSuiteName = "ECDHE-ECDSA-CHACHA20-POLY1305"
    Case TLS_CIPHER_SUITE_RSA_WITH_AES_128_GCM_SHA256
        pvCryptoCipherSuiteName = "AES128-GCM-SHA256"
    Case TLS_CIPHER_SUITE_RSA_WITH_AES_256_GCM_SHA384
        pvCryptoCipherSuiteName = "AES256-GCM-SHA384"
    End Select
End Function

Private Function pvCryptoCipherSuiteUseRsaCertificate(ByVal lCipherSuite As Long) As Boolean
    Select Case lCipherSuite
    Case TLS_CIPHER_SUITE_RSA_WITH_AES_128_GCM_SHA256, TLS_CIPHER_SUITE_RSA_WITH_AES_256_GCM_SHA384
        pvCryptoCipherSuiteUseRsaCertificate = True
    End Select
End Function

'= buffer management =====================================================

Private Function pvWriteBeginOfBlock(baBuffer() As Byte, ByVal lPos As Long, cStack As Collection, Optional ByVal Size As Long = 1) As Long
    If cStack Is Nothing Then
        Set cStack = New Collection
    End If
    If cStack.Count = 0 Then
        cStack.Add lPos
    Else
        cStack.Add lPos, Before:=1
    End If
    pvWriteBeginOfBlock = pvWriteReserved(baBuffer, lPos, Size)
    '--- note: keep Size in baBuffer
    baBuffer(lPos) = (Size And &HFF)
End Function

Private Function pvWriteEndOfBlock(baBuffer() As Byte, ByVal lPos As Long, cStack As Collection) As Long
    Dim lStart          As Long
    
    lStart = cStack.Item(1)
    cStack.Remove 1
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
        pvWriteLong = pvWriteReserved(baBuffer, lPos, Size)
        Call CopyMemory(baTemp(0), lValue, 4)
        baBuffer(lPos) = baTemp(Size - 1)
        baBuffer(lPos + 1) = baTemp(Size - 2)
        If Size >= 3 Then baBuffer(lPos + 2) = baTemp(Size - 3)
        If Size >= 4 Then baBuffer(lPos + 3) = baTemp(Size - 4)
    End If
End Function

Private Function pvWriteReserved(baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Long
    pvWriteReserved = pvWriteBuffer(baBuffer, lPos, 0, lSize)
End Function

Private Function pvWriteBuffer(baBuffer() As Byte, ByVal lPos As Long, ByVal lPtr As Long, ByVal lSize As Long) As Long
    Dim lBufPtr         As Long
    
    '--- peek long at ArrPtr(baBuffer)
    Call CopyMemory(lBufPtr, ByVal ArrPtr(baBuffer), 4)
    If lBufPtr = 0 Then
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

Private Function pvReadBeginOfBlock(baBuffer() As Byte, ByVal lPos As Long, cStack As Collection, Optional ByVal Size As Long = 1, Optional BlockSize As Long) As Long
    If cStack Is Nothing Then
        Set cStack = New Collection
    End If
    pvReadBeginOfBlock = pvReadLong(baBuffer, lPos, BlockSize, Size)
    If cStack.Count = 0 Then
        cStack.Add pvReadBeginOfBlock + BlockSize
    Else
        cStack.Add pvReadBeginOfBlock + BlockSize, Before:=1
    End If
End Function

Private Function pvReadEndOfBlock(baBuffer() As Byte, ByVal lPos As Long, cStack As Collection) As Long
    Dim lEnd          As Long
    
    #If baBuffer Then '--- touch args
    #End If
    lEnd = cStack.Item(1)
    cStack.Remove 1
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

'= arrays helpers ========================================================

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

Private Sub pvArraySwap(baBuffer() As Byte, lBufferPos As Long, baInput() As Byte, lInputPos As Long)
    Dim lTemp           As Long
    
    Call CopyMemory(lTemp, ByVal ArrPtr(baBuffer), 4)
    Call CopyMemory(ByVal ArrPtr(baBuffer), ByVal ArrPtr(baInput), 4)
    Call CopyMemory(ByVal ArrPtr(baInput), lTemp, 4)
    lTemp = lBufferPos
    lBufferPos = lInputPos
    lInputPos = lTemp
End Sub

'= global helpers ========================================================

Private Function ToHex(baText() As Byte, Optional Delimiter As String = "-") As String
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

Private Function FromHex(sText As String) As Byte()
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

Private Function EmptyByteArray(Optional ByVal Size As Long) As Byte()
    Dim baRetVal()      As Byte
    
    If Size > 0 Then
        ReDim baRetVal(0 To Size - 1) As Byte
    End If
    EmptyByteArray = baRetVal
End Function

Private Function SplitOrReindex(Expression As String, Delimiter As String) As Variant
    Dim vResult         As Variant
    Dim vTemp           As Variant
    Dim lIdx            As Long
    Dim lSize           As Long
    
    vResult = Split(Expression, Delimiter)
    '--- check if reindex needed
    If IsNumeric(vResult(0)) Then
        vTemp = vResult
        For lIdx = 0 To UBound(vTemp) Step 2
            If lSize < vTemp(lIdx) Then
                lSize = vTemp(lIdx)
            End If
        Next
        ReDim vResult(0 To lSize) As Variant
        For lIdx = 0 To UBound(vTemp) Step 2
            vResult(vTemp(lIdx)) = vTemp(lIdx + 1)
        Next
        SplitOrReindex = vResult
    End If
End Function

