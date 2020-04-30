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
Private Const TLS_HANDSHAKE_TYPE_ENCRYPTED_EXTENSIONS   As Long = 8
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
Private Const TLS_EXTENSION_TYPE_EC_POINT_FORMAT        As Long = 11
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
Private Const TLS_GROUP_SECP384R1                       As Long = 24
'Private Const TLS_GROUP_SECP521R1                       As Long = 25
Private Const TLS_GROUP_X25519                          As Long = 29
'Private Const TLS_GROUP_X448                            As Long = 30
Private Const TLS_SIGNATURE_RSA_PKCS1_SHA1              As Long = &H201
Private Const TLS_SIGNATURE_RSA_PKCS1_SHA256            As Long = &H401
Private Const TLS_SIGNATURE_RSA_PKCS1_SHA384            As Long = &H501
Private Const TLS_SIGNATURE_RSA_PKCS1_SHA512            As Long = &H601
Private Const TLS_SIGNATURE_ECDSA_SECP256R1_SHA256      As Long = &H403
Private Const TLS_SIGNATURE_ECDSA_SECP384R1_SHA384      As Long = &H503
Private Const TLS_SIGNATURE_ECDSA_SECP521R1_SHA512      As Long = &H603
Private Const TLS_SIGNATURE_RSA_PSS_RSAE_SHA256         As Long = &H804
Private Const TLS_SIGNATURE_RSA_PSS_RSAE_SHA384         As Long = &H805
Private Const TLS_SIGNATURE_RSA_PSS_RSAE_SHA512         As Long = &H806
Private Const TLS_SIGNATURE_RSA_PSS_PSS_SHA256          As Long = &H809
Private Const TLS_SIGNATURE_RSA_PSS_PSS_SHA384          As Long = &H80A
Private Const TLS_SIGNATURE_RSA_PSS_PSS_SHA512          As Long = &H80B
'Private Const TLS_PSK_KE_MODE_PSK_DHE                   As Long = 1
Private Const TLS_PROTOCOL_VERSION_TLS12                As Long = &H303
Private Const TLS_PROTOCOL_VERSION_TLS13                As Long = &H304
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
Private Const TLS_SECP256R1_KEY_SIZE                    As Long = 32
Private Const TLS_SECP384R1_KEY_SIZE                    As Long = 48
Private Const TLS_MAX_PLAINTEXT_RECORD_SIZE             As Long = 16384
Private Const TLS_MAX_ENCRYPTED_RECORD_SIZE             As Long = (16384 + 256)
Private Const TLS_RECORD_VERSION                        As Long = TLS_PROTOCOL_VERSION_TLS12 '--- always legacy version
Private Const TLS_LOCAL_LEGACY_VERSION                  As Long = &H303
Private Const TLS_HELLO_RANDOM_SIZE                     As Long = 32
Private Const szOID_RSA_RSA                             As String = "1.2.840.113549.1.1.1"
Private Const szOID_RSA_SSA_PSS                         As String = "1.2.840.113549.1.1.10"
Private Const szOID_ECC_PUBLIC_KEY                      As String = "1.2.840.10045.2.1"

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_VL_ALERTS             As String = "0|Close notify|10|Unexpected message|20|Bad record mac|40|Handshake failure|42|Bad certificate|44|Certificate revoked|45|Certificate expired|46|Certificate unknown|47|Illegal parameter|48|Unknown CA|50|Decode error|51|Decrypt error|70|Protocol version|80|Internal error|90|User canceled|109|Missing extension|112|Unrecognized name|116|Certificate required|120|No application protocol"
Private Const STR_UNKNOWN               As String = "Unknown (%1)"
Private Const STR_FORMAT_ALERT          As String = """%1"" alert"
'--- numeric
Private Const LNG_AAD_SIZE              As Long = 5     '--- size of additional authenticated data for TLS 1.3
Private Const LNG_LEGACY_AAD_SIZE       As Long = 13    '--- for TLS 1.2
Private Const LNG_ANS1_TYPE_SEQUENCE    As Long = &H30
Private Const LNG_ANS1_TYPE_INTEGER     As Long = &H2
'--- errors
Private Const ERR_CONNECTION_CLOSED     As String = "Connection closed"
Private Const ERR_GENER_KEYPAIR_FAILED  As String = "Failed generating key pair (%1)"
Private Const ERR_UNSUPPORTED_EXCH_GROUP As String = "Unsupported exchange group (%1)"
Private Const ERR_UNSUPPORTED_CIPHER_SUITE As String = "Unsupported cipher suite (%1)"
Private Const ERR_UNSUPPORTED_SIGNATURE_TYPE As String = "Unsupported signature type (%1)"
Private Const ERR_UNSUPPORTED_PUBLIC_KEY As String = "Unsupported public key OID (%1)"
Private Const ERR_UNSUPPORTED_CURVE_SIZE As String = "Unsupported curve size (%1)"
Private Const ERR_UNSUPPORTED_PROTOCOL  As String = "Invalid protocol version"
Private Const ERR_ENCRYPTION_FAILED     As String = "Encryption failed"
Private Const ERR_SIGNATURE_FAILED      As String = "Certificate signature failed (%1)"
Private Const ERR_DECRYPTION_FAILED     As String = "Decryption failed"
Private Const ERR_SERVER_HANDSHAKE_FAILED As String = "Handshake verification failed"
Private Const ERR_NEGOTIATE_SIGNATURE_FAILED As String = "Negotiate signature type failed"
Private Const ERR_RECORD_TOO_BIG        As String = "Record size too big"
Private Const ERR_FATAL_ALERT           As String = "Fatal alert"
Private Const ERR_UNEXPECTED_RECORD_TYPE As String = "Unexpected record type (%1)"
Private Const ERR_UNEXPECTED_MSG_TYPE   As String = "Unexpected message type for %1 (%2)"
Private Const ERR_UNEXPECTED_PROTOCOL   As String = "Unexpected protocol for %1 (%2)"
Private Const ERR_INVALID_STATE_HANDSHAKE As String = "Invalid state for handshake content (%1)"
Private Const ERR_INVALID_SIZE_KEY_SHARE As String = "Invalid data size for key share"
Private Const ERR_INVALID_REMOTE_KEY    As String = "Invalid remote key size"
Private Const ERR_INVALID_SIZE_REMOTE_KEY As String = "Invalid data size for remote key"
Private Const ERR_INVALID_SIZE_VERSIONS As String = "Invalid data size for supported versions"
Private Const ERR_INVALID_SIGNATURE     As String = "Invalid certificate signature"
Private Const ERR_COOKIE_NOT_ALLOWED    As String = "Cookie not allowed outside HelloRetryRequest"
Private Const ERR_NO_HANDSHAKE_MESSAGES As String = "Missing handshake messages"
Private Const ERR_NO_PREV_REMOTE_SECRET As String = "Missing previous remote secret"
Private Const ERR_NO_PREV_LOCAL_SECRET  As String = "Missing previous local secret"
Private Const ERR_NO_REMOTE_RANDOM      As String = "Missing remote random"
Private Const ERR_NO_SERVER_CERTIFICATE As String = "Missing server certificate"
Private Const ERR_NO_SUPPORTED_CIPHER_SUITE As String = "Missing supported ciphersuite"

Public Enum UcsTlsLocalFeaturesEnum '--- bitmask
    ucsTlsSupportTls12 = 2 ^ 0
    ucsTlsSupportTls13 = 2 ^ 1
    ucsTlsSupportAll = -1
End Enum

Public Enum UcsTlsStatesEnum
    ucsTlsStateClosed
    ucsTlsStateHandshakeStart
    ucsTlsStateExpectServerHello
    ucsTlsStateExpectExtensions
    ucsTlsStateExpectServerFinished     '--- not used in TLS 1.3
    '--- server states
    ucsTlsStateExpectClientHello
    ucsTlsStateExpectClientFinished
    ucsTlsStatePostHandshake
End Enum

Public Enum UcsTlsCryptoAlgorithmsEnum
    '--- key exchange
    ucsTlsAlgoExchX25519 = 1
    ucsTlsAlgoExchSecp256r1 = 2
    ucsTlsAlgoExchSecp384r1 = 3
    ucsTlsAlgoExchCertificate = 4
    '--- authenticated encryption w/ additional data
    ucsTlsAlgoAeadChacha20Poly1305 = 11
    ucsTlsAlgoAeadAes128 = 12
    ucsTlsAlgoAeadAes256 = 13
    '--- hash
    ucsTlsAlgoDigestSha256 = 21
    ucsTlsAlgoDigestSha384 = 22
    ucsTlsAlgoDigestSha512 = 23
    '--- verify signature
    ucsTlsAlgoSignaturePss = 31
    ucsTlsAlgoSignaturePkcsSha2 = 32
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
    IsServer            As Boolean
    RemoteHostName      As String
    LocalFeatures       As UcsTlsLocalFeaturesEnum
    '--- state
    State               As UcsTlsStatesEnum
    LastError           As String
    LastAlertCode       As UcsTlsAlertDescriptionsEnum
    BlocksStack         As Collection
    '--- handshake
    LocalSessionID()    As Byte
    LocalExchRandom()   As Byte
    LocalExchPrivate()  As Byte
    LocalExchPublic()   As Byte
    LocalExchRsaEncrPriv() As Byte
    LocalCertificates   As Collection
    LocalPrivateKey()   As Byte
    LocalSignatureType  As Long
    RemoteSessionID()   As Byte
    RemoteExchRandom()  As Byte
    RemoteExchPublic()  As Byte
    RemoteCertificates  As Collection
    RemoteExtensions    As Collection
    '--- crypto settings
    ProtocolVersion     As Long
    ExchGroup           As Long
    ExchAlgo            As UcsTlsCryptoAlgorithmsEnum
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
    RemoteTrafficSecret() As Byte
    RemoteTrafficKey()  As Byte
    RemoteTrafficIV()   As Byte
    RemoteTrafficSeqNo  As Long
    LocalTrafficSecret() As Byte
    LocalTrafficKey()   As Byte
    LocalTrafficIV()    As Byte
    LocalTrafficSeqNo   As Long
    '--- hello retry request
    HelloRetryRequest   As Boolean
    HelloRetryCipherSuite As Long
    HelloRetryExchGroup As Long
    HelloRetryCookie()  As Byte
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
            Optional RemoteHostName As String, _
            Optional ByVal LocalFeatures As UcsTlsLocalFeaturesEnum = ucsTlsSupportAll) As Boolean
    Dim uEmpty          As UcsTlsContext
    
    On Error GoTo EH
    If Not CryptoInit() Then
        GoTo QH
    End If
    With uEmpty
        pvTlsSetLastError uEmpty, vbNullString
        .State = ucsTlsStateHandshakeStart
        .RemoteHostName = RemoteHostName
        .LocalFeatures = LocalFeatures
        .LocalExchRandom = pvTlsArrayRandom(TLS_HELLO_RANDOM_SIZE)
        If (LocalFeatures And ucsTlsSupportTls13) <> 0 Then
            '--- note: uCtx.ClientPublic has to be ready for pvTlsBuildClientHello
            If Not pvTlsSetupExchEccGroup(uEmpty, TLS_GROUP_X25519, .LastError, .LastAlertCode) Then
                pvTlsSetLastError uCtx, .LastError, .LastAlertCode
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
    pvTlsSetLastError uCtx, Err.Description
    Resume QH
End Function

Public Function TlsInitServer( _
            uCtx As UcsTlsContext, _
            Optional RemoteHostName As String, _
            Optional Certificates As Collection, _
            Optional PrivateKey As Variant) As Boolean
    Dim uEmpty          As UcsTlsContext
    
    On Error GoTo EH
    If Not CryptoInit() Then
        GoTo QH
    End If
    With uEmpty
        pvTlsSetLastError uEmpty, vbNullString
        .IsServer = True
        .State = ucsTlsStateExpectClientHello
        .RemoteHostName = RemoteHostName
        .LocalFeatures = ucsTlsSupportTls13
        Set .LocalCertificates = Certificates
        If Not IsMissing(PrivateKey) Then
            .LocalPrivateKey = PrivateKey
        End If
        .LocalExchRandom = pvTlsArrayRandom(TLS_HELLO_RANDOM_SIZE)
    End With
    uCtx = uEmpty
    '--- success
    TlsInitServer = True
QH:
    Exit Function
EH:
    pvTlsSetLastError uCtx, Err.Description
    Resume QH
End Function

Public Function TlsHandshake(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, baOutput() As Byte, lPos As Long) As Boolean
    On Error GoTo EH
    With uCtx
        If .State = ucsTlsStateClosed Then
            pvTlsSetLastError uCtx, ERR_CONNECTION_CLOSED
            Exit Function
        End If
        pvTlsSetLastError uCtx, vbNullString
        '--- swap-in
        pvArraySwap .SendBuffer, .SendPos, baOutput, lPos
        If .State = ucsTlsStateHandshakeStart Then
            .SendPos = pvTlsBuildClientHello(uCtx, .SendBuffer, .SendPos)
            .State = ucsTlsStateExpectServerHello
        Else
            If lSize < 0 Then
                lSize = pvArraySize(baInput)
            End If
            If Not pvTlsParsePayload(uCtx, baInput, lSize, .LastError, .LastAlertCode) Then
                pvTlsSetLastError uCtx, .LastError, .LastAlertCode
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
    pvTlsSetLastError uCtx, Err.Description
    Resume QH
End Function

Public Function TlsSend(uCtx As UcsTlsContext, baPlainText() As Byte, ByVal lSize As Long, baOutput() As Byte, lPos As Long) As Boolean
    On Error GoTo EH
    With uCtx
        If lSize < 0 Then
            lSize = pvArraySize(baPlainText)
        End If
        If lSize = 0 Then
            '--- flush
            pvArraySwap .SendBuffer, .SendPos, baOutput, lPos
            Erase .SendBuffer
            .SendPos = 0
            '--- success
            TlsSend = True
            Exit Function
        End If
        If .State = ucsTlsStateClosed Then
            pvTlsSetLastError uCtx, ERR_CONNECTION_CLOSED
            Exit Function
        End If
        pvTlsSetLastError uCtx, vbNullString
        '--- swap-in
        pvArraySwap .SendBuffer, .SendPos, baOutput, lPos
        .SendPos = pvTlsBuildApplicationData(uCtx, .SendBuffer, .SendPos, baPlainText, lSize, TLS_CONTENT_TYPE_APPDATA, .LastError, .LastAlertCode)
        If LenB(.LastError) <> 0 Then
            pvTlsSetLastError uCtx, .LastError, .LastAlertCode
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
    pvTlsSetLastError uCtx, Err.Description
    Resume QH
End Function

Public Function TlsReceive(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, baPlainText() As Byte, lPos As Long) As Boolean
    On Error GoTo EH
    With uCtx
        If lSize < 0 Then
            lSize = pvArraySize(baInput)
        End If
        If lSize = 0 And lPos = 0 Then
            '--- flush
            pvArraySwap .DecrBuffer, .DecrPos, baPlainText, lPos
            Erase .DecrBuffer
            .DecrPos = 0
            '--- success
            TlsReceive = True
            Exit Function
        End If
        If .State = ucsTlsStateClosed Then
            pvTlsSetLastError uCtx, ERR_CONNECTION_CLOSED
            Exit Function
        End If
        pvTlsSetLastError uCtx, vbNullString
        '--- swap-in
        pvArraySwap .DecrBuffer, .DecrPos, baPlainText, lPos
        If Not pvTlsParsePayload(uCtx, baInput, lSize, .LastError, .LastAlertCode) Then
            pvTlsSetLastError uCtx, .LastError, .LastAlertCode
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
    pvTlsSetLastError uCtx, Err.Description
    Resume QH
End Function

Public Function TlsShutdown(uCtx As UcsTlsContext, baOutput() As Byte, lPos As Long) As Boolean
    On Error GoTo EH
    With uCtx
        If .State = ucsTlsStateClosed Then
            Exit Function
        End If
        pvTlsSetLastError uCtx, vbNullString
        '--- swap-in
        pvArraySwap .SendBuffer, .SendPos, baOutput, lPos
        .SendPos = pvTlsBuildAlert(uCtx, .SendBuffer, .SendPos, uscTlsAlertCloseNotify, TLS_ALERT_LEVEL_WARNING, .LastError, .LastAlertCode)
        If LenB(.LastError) <> 0 Then
            pvTlsSetLastError uCtx, .LastError, .LastAlertCode
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
    pvTlsSetLastError uCtx, Err.Description
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
            TlsGetLastAlert = Replace(STR_UNKNOWN, "%1", AlertCode)
        End If
    End If
End Function

'= private ===============================================================

Private Function pvTlsBuildClientHello(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long) As Long
    Dim lMessagePos     As Long
    Dim vElem           As Variant
    
    With uCtx
        '--- Record Header
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            '--- Handshake Header
            lMessagePos = lPos
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_CLIENT_HELLO)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                lPos = pvWriteLong(baOutput, lPos, TLS_LOCAL_LEGACY_VERSION, Size:=2)
                lPos = pvWriteArray(baOutput, lPos, .LocalExchRandom)
                '--- Legacy Session ID
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack)
                    If pvArraySize(.LocalSessionID) = 0 And (.LocalFeatures And ucsTlsSupportTls12) <> 0 Then
                        '--- non-empty for TLS 1.2 compatibility
                        lPos = pvWriteArray(baOutput, lPos, pvTlsArrayRandom(TLS_HELLO_RANDOM_SIZE))
                    Else
                        lPos = pvWriteArray(baOutput, lPos, .LocalSessionID)
                    End If
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                '--- Cipher Suites
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                    For Each vElem In pvTlsPrepareCipherSuitsOrder(.LocalFeatures)
                        lPos = pvWriteLong(baOutput, lPos, vElem, Size:=2)
                    Next
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                '--- Legacy Compression Methods
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack)
                    lPos = pvWriteLong(baOutput, lPos, TLS_COMPRESS_NULL)
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                '--- Extensions
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                    If LenB(.RemoteHostName) <> 0 Then
                        '--- Extension - Server Name
                        lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_SERVER_NAME, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                                lPos = pvWriteLong(baOutput, lPos, TLS_SERVER_NAME_TYPE_HOSTNAME) '--- FQDN
                                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                                    lPos = pvWriteString(baOutput, lPos, .RemoteHostName)
                                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    End If
                    '--- Extension - Supported Groups
                    lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_SUPPORTED_GROUPS, Size:=2)
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            If CryptoIsSupported(ucsTlsAlgoExchX25519) Then
                                If .HelloRetryExchGroup = 0 Or .HelloRetryExchGroup = TLS_GROUP_X25519 Then
                                    lPos = pvWriteLong(baOutput, lPos, TLS_GROUP_X25519, Size:=2)
                                End If
                            End If
                            If CryptoIsSupported(ucsTlsAlgoExchSecp256r1) Then
                                If .HelloRetryExchGroup = 0 Or .HelloRetryExchGroup = TLS_GROUP_SECP256R1 Then
                                    lPos = pvWriteLong(baOutput, lPos, TLS_GROUP_SECP256R1, Size:=2)
                                End If
                            End If
                            If CryptoIsSupported(ucsTlsAlgoExchSecp384r1) Then
                                If .HelloRetryExchGroup = 0 Or .HelloRetryExchGroup = TLS_GROUP_SECP384R1 Then
                                    lPos = pvWriteLong(baOutput, lPos, TLS_GROUP_SECP384R1, Size:=2)
                                End If
                            End If
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    If (.LocalFeatures And ucsTlsSupportTls12) <> 0 Then
                        '--- Extension - EC Point Formats
                        lPos = pvWriteArray(baOutput, lPos, pvArrayByte(0, TLS_EXTENSION_TYPE_EC_POINT_FORMAT, 0, 2, 1, 0))   '--- uncompressed only
                        '--- Extension - Renegotiation Info
                        lPos = pvWriteArray(baOutput, lPos, pvArrayByte(&HFF, 1, 0, 1, 0))     '--- empty info
                    End If
                    '--- Extension - Signature Algorithms
                    lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_SIGNATURE_ALGORITHMS, Size:=2)
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_ECDSA_SECP256R1_SHA256, Size:=2)
                            lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_ECDSA_SECP384R1_SHA384, Size:=2)
                            lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_ECDSA_SECP521R1_SHA512, Size:=2)
                            If CryptoIsSupported(ucsTlsAlgoSignaturePss) Then
                                lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, Size:=2)
                                lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, Size:=2)
                                lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PSS_RSAE_SHA512, Size:=2)
                                lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PSS_PSS_SHA256, Size:=2)
                                lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PSS_PSS_SHA384, Size:=2)
                                lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PSS_PSS_SHA512, Size:=2)
                            End If
                            If CryptoIsSupported(ucsTlsAlgoSignaturePkcsSha2) Then
                                lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PKCS1_SHA256, Size:=2)
                                lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PKCS1_SHA384, Size:=2)
                                lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PKCS1_SHA512, Size:=2)
                            End If
                            lPos = pvWriteLong(baOutput, lPos, TLS_SIGNATURE_RSA_PKCS1_SHA1, Size:=2)
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    If (.LocalFeatures And ucsTlsSupportTls13) <> 0 Then
                        '--- Extension - Key Share
                        lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_KEY_SHARE, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                                lPos = pvWriteLong(baOutput, lPos, .ExchGroup, Size:=2)
                                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                                    lPos = pvWriteArray(baOutput, lPos, .LocalExchPublic)
                                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        '--- Extension - Supported Versions
                        lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_SUPPORTED_VERSIONS, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack)
                                lPos = pvWriteLong(baOutput, lPos, TLS_PROTOCOL_VERSION_TLS13, Size:=2)
                                If (.LocalFeatures And ucsTlsSupportTls12) <> 0 Then
                                    lPos = pvWriteLong(baOutput, lPos, TLS_PROTOCOL_VERSION_TLS12, Size:=2)
                                End If
                            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        If .HelloRetryRequest Then
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
    pvTlsBuildClientHello = lPos
End Function

Private Function pvTlsBuildClientLegacyKeyExchange(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Long
    Dim baLocalIV()     As Byte
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
            '--- Handshake Client Key Exchange
            lMessagePos = lPos
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_CLIENT_KEY_EXCHANGE)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                If pvArraySize(.LocalExchRsaEncrPriv) > 0 Then
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                        lPos = pvWriteArray(baOutput, lPos, .LocalExchRsaEncrPriv)
                    lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                Else
                    lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack)
                        lPos = pvWriteArray(baOutput, lPos, .LocalExchPublic)
                    lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                End If
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lMessagePos)), lPos - lMessagePos
        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        '--- Legacy Change Cipher Spec
        lPos = pvWriteArray(baOutput, lPos, pvArrayByte(TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC, TLS_RECORD_VERSION \ &H100, TLS_RECORD_VERSION, 0, 1, 1))
        '--- Record Header
        lRecordPos = lPos
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            baLocalIV = pvArrayXor(.LocalTrafficIV, .LocalTrafficSeqNo)
            If .IvDynamicSize > 0 Then '--- AES in TLS 1.2
                lPos = pvWriteBuffer(baOutput, lPos, VarPtr(baLocalIV(.IvSize - .IvDynamicSize)), .IvDynamicSize)
            End If
            '--- Handshake Finish
            lMessagePos = lPos
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_FINISHED)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                baHandshakeHash = pvTlsArrayHash(.DigestAlgo, .HandshakeMessages, 0)
                baVerifyData = pvTlsKdfLegacyPrf(.DigestAlgo, .MasterSecret, "client finished", baHandshakeHash, 12)
                lPos = pvWriteArray(baOutput, lPos, baVerifyData)
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            lMessageSize = lPos - lMessagePos
            '--- note: *before* allocating space for the authentication tag
            pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lMessagePos)), lPos - lMessagePos
            lPos = pvWriteReserved(baOutput, lPos, .TagSize)
            '--- encrypt message
            ReDim baAad(0 To LNG_LEGACY_AAD_SIZE - 1) As Byte
            Debug.Assert RedimStats("pvTlsBuildClientLegacyKeyExchange.baAad", UBound(baAad) + 1)
            lAadPos = pvWriteLong(baAad, 0, 0, Size:=4)
            lAadPos = pvWriteLong(baAad, lAadPos, .LocalTrafficSeqNo, Size:=4)
            lAadPos = pvWriteBuffer(baAad, lAadPos, VarPtr(baOutput(lRecordPos)), 3)
            lAadPos = pvWriteLong(baAad, lAadPos, lMessageSize, Size:=2)
            Debug.Assert lAadPos = LNG_LEGACY_AAD_SIZE
            If Not pvTlsAeadEncrypt(.AeadAlgo, baLocalIV, .LocalTrafficKey, baAad, 0, UBound(baAad) + 1, baOutput, lMessagePos, lMessageSize) Then
                sError = ERR_ENCRYPTION_FAILED
                eAlertCode = uscTlsAlertInternalError
                GoTo QH
            End If
            .LocalTrafficSeqNo = UnsignedAdd(.LocalTrafficSeqNo, 1)
            lMessagePos = lRecordPos + 5
        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
    End With
    pvTlsBuildClientLegacyKeyExchange = lPos
QH:
End Function

Private Function pvTlsBuildClientHandshakeFinished(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Long
    Dim lRecordPos      As Long
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim baLocalIV()     As Byte
    Dim baHandshakeHash() As Byte
    Dim baVerifyData()  As Byte
    
    With uCtx
        '--- Legacy Change Cipher Spec
        lPos = pvWriteArray(baOutput, lPos, pvArrayByte(TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC, TLS_RECORD_VERSION \ &H100, TLS_RECORD_VERSION, 0, 1, 1))
        '--- Record Header
        lRecordPos = lPos
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_APPDATA)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            '--- Client Handshake Finished
            lMessagePos = lPos
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_FINISHED)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                baHandshakeHash = pvTlsArrayHash(.DigestAlgo, .HandshakeMessages, 0)
                baVerifyData = pvTlsHkdfExpandLabel(.DigestAlgo, .LocalTrafficSecret, "finished", EmptyByteArray, .DigestSize)
                baVerifyData = pvTlsHkdfExtract(.DigestAlgo, baVerifyData, baHandshakeHash)
                lPos = pvWriteArray(baOutput, lPos, baVerifyData)
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            '--- Record Type
            lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
            lMessageSize = lPos - lMessagePos
            lPos = pvWriteReserved(baOutput, lPos, .TagSize)
        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        baLocalIV = pvArrayXor(.LocalTrafficIV, .LocalTrafficSeqNo)
        If Not pvTlsAeadEncrypt(.AeadAlgo, baLocalIV, .LocalTrafficKey, baOutput, lRecordPos, LNG_AAD_SIZE, baOutput, lMessagePos, lMessageSize) Then
            sError = ERR_ENCRYPTION_FAILED
            eAlertCode = uscTlsAlertInternalError
            GoTo QH
        End If
        .LocalTrafficSeqNo = UnsignedAdd(.LocalTrafficSeqNo, 1)
    End With
    pvTlsBuildClientHandshakeFinished = lPos
QH:
End Function

Private Function pvTlsBuildServerHello(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long) As Long
    Dim lMessagePos     As Long
    
    With uCtx
        '--- Record Header
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            '--- Handshake Header
            lMessagePos = lPos
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_SERVER_HELLO)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                lPos = pvWriteLong(baOutput, lPos, TLS_LOCAL_LEGACY_VERSION, Size:=2)
                lPos = pvWriteArray(baOutput, lPos, .LocalExchRandom)
                '--- Legacy Session ID
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack)
                    lPos = pvWriteArray(baOutput, lPos, .RemoteSessionID)
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                '--- Cipher Suite
                lPos = pvWriteLong(baOutput, lPos, .CipherSuite, Size:=2)
                '--- Legacy Compression Method
                lPos = pvWriteLong(baOutput, lPos, TLS_COMPRESS_NULL)
                '--- Extensions
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                    '--- Extension - Key Share
                    If SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_TYPE_KEY_SHARE) Then
                        lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_KEY_SHARE, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            lPos = pvWriteLong(baOutput, lPos, .ExchGroup, Size:=2)
                            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                                lPos = pvWriteArray(baOutput, lPos, .LocalExchPublic)
                            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    End If
                    '--- Extension - Supported Versions
                    If SearchCollection(.RemoteExtensions, "#" & TLS_EXTENSION_TYPE_SUPPORTED_VERSIONS) Then
                        lPos = pvWriteLong(baOutput, lPos, TLS_EXTENSION_TYPE_SUPPORTED_VERSIONS, Size:=2)
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            lPos = pvWriteLong(baOutput, lPos, TLS_PROTOCOL_VERSION_TLS13, Size:=2)
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    End If
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lMessagePos)), lPos - lMessagePos
    End With
    pvTlsBuildServerHello = lPos
End Function

Private Function pvTlsBuildServerHandshakeFinished(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Long
    Dim lRecordPos      As Long
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim baLocalIV()     As Byte
    Dim baHandshakeHash() As Byte
    Dim lHandshakePos   As Long
    Dim baVerifyData()  As Byte
    Dim lVerifyPos      As Long
    Dim lIdx            As Long
    Dim baCert()        As Byte
    Dim baSignature()   As Byte
    
    With uCtx
        '--- Legacy Change Cipher Spec
        lPos = pvWriteArray(baOutput, lPos, pvArrayByte(TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC, TLS_RECORD_VERSION \ &H100, TLS_RECORD_VERSION, 0, 1, 1))
        '--- Record Header
        lRecordPos = lPos
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_APPDATA)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            lMessagePos = lPos
            '--- Server Encrypted Extensions
            lHandshakePos = lPos
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_ENCRYPTED_EXTENSIONS)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                    lPos = lPos '--- empty
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            '--- Server Certificate
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_CERTIFICATE)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                '--- certificate request context
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack)
                    lPos = lPos '--- empty
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                    For lIdx = 1 To .LocalCertificates.Count
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                            baCert = .LocalCertificates.Item(lIdx)
                            lPos = pvWriteArray(baOutput, lPos, baCert)
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                        '--- certificate extensions
                        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                            lPos = lPos '--- empty
                        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
                    Next
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lHandshakePos)), lPos - lHandshakePos
            '--- Server Certificate Verify
            lHandshakePos = lPos
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_CERTIFICATE_VERIFY)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                lPos = pvWriteLong(baOutput, lPos, .LocalSignatureType, Size:=2)
                lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
                    baHandshakeHash = pvTlsArrayHash(.DigestAlgo, .HandshakeMessages, 0)
                    lVerifyPos = pvWriteString(baVerifyData, 0, Space$(64) & "TLS 1.3, server CertificateVerify" & Chr$(0))
                    lVerifyPos = pvWriteArray(baVerifyData, lVerifyPos, baHandshakeHash)
                    If Not pvTlsSignatureSign(.LocalPrivateKey, .LocalSignatureType, baVerifyData, baSignature, sError, eAlertCode) Then
                        GoTo QH
                    End If
                    lPos = pvWriteArray(baOutput, lPos, baSignature)
                lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lHandshakePos)), lPos - lHandshakePos
            '--- Server Handshake Finished
            lHandshakePos = lPos
            lPos = pvWriteLong(baOutput, lPos, TLS_HANDSHAKE_TYPE_FINISHED)
            lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=3)
                baHandshakeHash = pvTlsArrayHash(.DigestAlgo, .HandshakeMessages, 0)
                baVerifyData = pvTlsHkdfExpandLabel(.DigestAlgo, .LocalTrafficSecret, "finished", EmptyByteArray, .DigestSize)
                baVerifyData = pvTlsHkdfExtract(.DigestAlgo, baVerifyData, baHandshakeHash)
                lPos = pvWriteArray(baOutput, lPos, baVerifyData)
            lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
            pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baOutput(lHandshakePos)), lPos - lHandshakePos
            '--- Record Type
            lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
            lMessageSize = lPos - lMessagePos
            lPos = pvWriteReserved(baOutput, lPos, .TagSize)
        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        baLocalIV = pvArrayXor(.LocalTrafficIV, .LocalTrafficSeqNo)
        If Not pvTlsAeadEncrypt(.AeadAlgo, baLocalIV, .LocalTrafficKey, baOutput, lRecordPos, LNG_AAD_SIZE, baOutput, lMessagePos, lMessageSize) Then
            sError = ERR_ENCRYPTION_FAILED
            eAlertCode = uscTlsAlertInternalError
            GoTo QH
        End If
        .LocalTrafficSeqNo = UnsignedAdd(.LocalTrafficSeqNo, 1)
    End With
    pvTlsBuildServerHandshakeFinished = lPos
QH:
End Function

Private Function pvTlsBuildApplicationData(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long, baData() As Byte, ByVal lSize As Long, ByVal lContentType As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Long
    Dim lRecordPos      As Long
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim baLocalIV()     As Byte
    Dim baAad()         As Byte
    Dim lAadPos         As Long
    Dim bResult         As Boolean
    
    With uCtx
        lRecordPos = lPos
        '--- Record Header
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_APPDATA)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            baLocalIV = pvArrayXor(.LocalTrafficIV, .LocalTrafficSeqNo)
            If .IvDynamicSize > 0 Then '--- AES in TLS 1.2
                lPos = pvWriteBuffer(baOutput, lPos, VarPtr(baLocalIV(.IvSize - .IvDynamicSize)), .IvDynamicSize)
            End If
            lMessagePos = lPos
            If lSize > 0 Then
                lPos = pvWriteBuffer(baOutput, lPos, VarPtr(baData(0)), lSize)
            End If
            If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                lPos = pvWriteLong(baOutput, lPos, lContentType)
            End If
            lMessageSize = lPos - lMessagePos
            lPos = pvWriteReserved(baOutput, lPos, .TagSize)
        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        '--- encrypt message
        If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
            bResult = pvTlsAeadEncrypt(.AeadAlgo, baLocalIV, .LocalTrafficKey, baOutput, lRecordPos, LNG_AAD_SIZE, baOutput, lMessagePos, lMessageSize)
        ElseIf .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
            ReDim baAad(0 To LNG_LEGACY_AAD_SIZE - 1) As Byte
            Debug.Assert RedimStats("pvTlsBuildApplicationData.baAad", UBound(baAad) + 1)
            lAadPos = pvWriteLong(baAad, 0, 0, Size:=4)
            lAadPos = pvWriteLong(baAad, lAadPos, .LocalTrafficSeqNo, Size:=4)
            lAadPos = pvWriteBuffer(baAad, lAadPos, VarPtr(baOutput(lRecordPos)), 3)
            lAadPos = pvWriteLong(baAad, lAadPos, lMessageSize, Size:=2)
            Debug.Assert lAadPos = LNG_LEGACY_AAD_SIZE
            bResult = pvTlsAeadEncrypt(.AeadAlgo, baLocalIV, .LocalTrafficKey, baAad, 0, UBound(baAad) + 1, baOutput, lMessagePos, lMessageSize)
        End If
        If Not bResult Then
            sError = ERR_ENCRYPTION_FAILED
            eAlertCode = uscTlsAlertInternalError
            GoTo QH
        End If
        .LocalTrafficSeqNo = UnsignedAdd(.LocalTrafficSeqNo, 1)
    End With
    pvTlsBuildApplicationData = lPos
QH:
End Function

Private Function pvTlsBuildAlert(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long, ByVal eAlertDesc As UcsTlsAlertDescriptionsEnum, ByVal lAlertLevel As Long, Optional sError As String, Optional eAlertCode As UcsTlsAlertDescriptionsEnum) As Long
    Dim lRecordPos      As Long
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim baLocalIV()     As Byte
    Dim baAad()         As Byte
    Dim lAadPos         As Long
    
    With uCtx
        '--- for TLS 1.3 -> tunnel alert through application data encryption
        If .State = ucsTlsStatePostHandshake And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
            ReDim baLocalIV(0 To 1) As Byte
            Debug.Assert RedimStats("pvTlsBuildAlert.baLocalIV", UBound(baLocalIV) + 1)
            baLocalIV(0) = lAlertLevel
            baLocalIV(1) = eAlertDesc
            pvTlsBuildAlert = pvTlsBuildApplicationData(uCtx, baOutput, lPos, baLocalIV, UBound(baLocalIV) + 1, TLS_CONTENT_TYPE_ALERT, sError, eAlertCode)
            GoTo QH
        End If
        lRecordPos = lPos
        '--- Record Header
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_ALERT)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, .BlocksStack, Size:=2)
            If .State = ucsTlsStatePostHandshake And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                baLocalIV = pvArrayXor(.LocalTrafficIV, .LocalTrafficSeqNo)
                If .IvDynamicSize > 0 Then '--- AES in TLS 1.2
                    lPos = pvWriteBuffer(baOutput, lPos, VarPtr(baLocalIV(.IvSize - .IvDynamicSize)), .IvDynamicSize)
                End If
            End If
            lMessagePos = lPos
            lPos = pvWriteLong(baOutput, lPos, lAlertLevel)
            lPos = pvWriteLong(baOutput, lPos, eAlertDesc)
            lMessageSize = lPos - lMessagePos
            If .State = ucsTlsStatePostHandshake And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                lPos = pvWriteReserved(baOutput, lPos, .TagSize)
            End If
        lPos = pvWriteEndOfBlock(baOutput, lPos, .BlocksStack)
        If .State = ucsTlsStatePostHandshake And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
            '--- encrypt message
            ReDim baAad(0 To LNG_LEGACY_AAD_SIZE - 1) As Byte
            Debug.Assert RedimStats("pvTlsBuildAlert.baAad", UBound(baAad) + 1)
            lAadPos = pvWriteLong(baAad, 0, 0, Size:=4)
            lAadPos = pvWriteLong(baAad, lAadPos, .LocalTrafficSeqNo, Size:=4)
            lAadPos = pvWriteBuffer(baAad, lAadPos, VarPtr(baOutput(lRecordPos)), 3)
            lAadPos = pvWriteLong(baAad, lAadPos, lMessageSize, Size:=2)
            Debug.Assert lAadPos = LNG_LEGACY_AAD_SIZE
            If Not pvTlsAeadEncrypt(.AeadAlgo, baLocalIV, .LocalTrafficKey, baAad, 0, UBound(baAad) + 1, baOutput, lMessagePos, lMessageSize) Then
                sError = ERR_ENCRYPTION_FAILED
                eAlertCode = uscTlsAlertInternalError
                GoTo QH
            End If
            .LocalTrafficSeqNo = UnsignedAdd(.LocalTrafficSeqNo, 1)
        End If
    End With
    pvTlsBuildAlert = lPos
QH:
End Function

Private Function pvTlsParsePayload(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Dim lPrevPos        As Long
    Dim lRecvSize       As Long
    
    If lSize > 0 Then
    With uCtx
        .RecvPos = pvWriteBuffer(.RecvBuffer, .RecvPos, VarPtr(baInput(0)), lSize)
        lPrevPos = .RecvPos
        .RecvPos = pvTlsParseRecord(uCtx, .RecvBuffer, .RecvPos, sError, eAlertCode)
        If LenB(sError) <> 0 Then
            GoTo QH
        End If
        lRecvSize = lPrevPos - .RecvPos
        If .RecvPos > 0 And lRecvSize > 0 Then
            Call CopyMemory(.RecvBuffer(0), .RecvBuffer(.RecvPos), lRecvSize)
        End If
        .RecvPos = IIf(lRecvSize > 0, lRecvSize, 0)
    End With
    End If
    '--- success
    pvTlsParsePayload = True
QH:
End Function

Private Function pvTlsParseRecord(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Long
    Dim lRecordPos      As Long
    Dim lRecordSize     As Long
    Dim lRecordType     As Long
    Dim lRecordProtocol As Long
    Dim baRemoteIV()    As Byte
    Dim lPos            As Long
    Dim lEnd            As Long
    Dim baAad()         As Byte
    Dim bResult         As Boolean
    
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
                    If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                        '--- note: TLS_CONTENT_TYPE_ALERT encryption is tunneled through TLS_CONTENT_TYPE_APPDATA
                        sError = ERR_RECORD_TOO_BIG
                        eAlertCode = uscTlsAlertDecodeError
                        GoTo QH
                    ElseIf .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                        If Not pvTlsPrepareLegacyDecryptParams(uCtx, baInput, lRecordPos, lRecordSize, lPos, lEnd, baRemoteIV, baAad) Then
                            sError = "pvTlsPrepareLegacyDecryptParams"
                            eAlertCode = uscTlsAlertInternalError
                            GoTo QH
                        End If
                        bResult = pvTlsAeadDecrypt(.AeadAlgo, baRemoteIV, .RemoteTrafficKey, baAad, 0, UBound(baAad) + 1, baInput, lPos, lEnd - lPos + .TagSize)
                    Else
                        bResult = False
                    End If
                    If Not bResult Then
                        sError = ERR_DECRYPTION_FAILED
                        eAlertCode = uscTlsAlertBadRecordMac
                        GoTo QH
                    End If
                    .RemoteTrafficSeqNo = UnsignedAdd(.RemoteTrafficSeqNo, 1)
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
                If .State = ucsTlsStateExpectServerFinished Then
                    If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                        '--- note: ucsTlsStateExpectServerFinished is TLS 1.2 state only
                        sError = Replace(Replace(ERR_UNEXPECTED_PROTOCOL, "%1", "ucsTlsStateExpectServerFinished"), "%2", .ProtocolVersion)
                        eAlertCode = uscTlsAlertInternalError
                        GoTo QH
                    ElseIf .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                        If Not pvTlsPrepareLegacyDecryptParams(uCtx, baInput, lRecordPos, lRecordSize, lPos, lEnd, baRemoteIV, baAad) Then
                            sError = "pvTlsPrepareLegacyDecryptParams"
                            eAlertCode = uscTlsAlertInternalError
                            GoTo QH
                        End If
                        bResult = pvTlsAeadDecrypt(.AeadAlgo, baRemoteIV, .RemoteTrafficKey, baAad, 0, UBound(baAad) + 1, baInput, lPos, lEnd - lPos + .TagSize)
                    Else
                        bResult = False
                    End If
                    If Not bResult Then
                        sError = ERR_DECRYPTION_FAILED
                        eAlertCode = uscTlsAlertBadRecordMac
                        GoTo QH
                    End If
                    .RemoteTrafficSeqNo = UnsignedAdd(.RemoteTrafficSeqNo, 1)
                End If
HandleHandshakeContent:
                If .MessSize > 0 Then
                    .MessSize = pvWriteBuffer(.MessBuffer, .MessSize, VarPtr(baInput(lPos)), lEnd - lPos)
                    If Not pvTlsParseHandshake(uCtx, .MessBuffer, .MessPos, .MessSize, lRecordProtocol, sError, eAlertCode) Then
                        GoTo QH
                    End If
                    If .MessPos >= .MessSize Then
                        Erase .MessBuffer
                        .MessSize = 0
                        .MessPos = 0
                    End If
                Else
                    If Not pvTlsParseHandshake(uCtx, baInput, lPos, lEnd, lRecordProtocol, sError, eAlertCode) Then
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
                If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                    baRemoteIV = pvArrayXor(.RemoteTrafficIV, .RemoteTrafficSeqNo)
                    bResult = pvTlsAeadDecrypt(.AeadAlgo, baRemoteIV, .RemoteTrafficKey, baInput, lRecordPos, LNG_AAD_SIZE, baInput, lPos, lRecordSize)
                ElseIf .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                    If Not pvTlsPrepareLegacyDecryptParams(uCtx, baInput, lRecordPos, lRecordSize, lPos, lEnd, baRemoteIV, baAad) Then
                        sError = "pvTlsPrepareLegacyDecryptParams"
                        eAlertCode = uscTlsAlertInternalError
                        GoTo QH
                    End If
                    bResult = pvTlsAeadDecrypt(.AeadAlgo, baRemoteIV, .RemoteTrafficKey, baAad, 0, UBound(baAad) + 1, baInput, lPos, lEnd - lPos + .TagSize)
                Else
                    bResult = False
                End If
                If Not bResult Then
                    sError = ERR_DECRYPTION_FAILED
                    eAlertCode = uscTlsAlertBadRecordMac
                    GoTo QH
                End If
                .RemoteTrafficSeqNo = UnsignedAdd(.RemoteTrafficSeqNo, 1)
                If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
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
    pvTlsParseRecord = lPos
QH:
End Function

Private Function pvTlsParseHandshake(uCtx As UcsTlsContext, baInput() As Byte, lPos As Long, ByVal lEnd As Long, ByVal lRecordProtocol As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
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
    Dim lSignPos        As Long
    Dim lSignSize       As Long
    
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
                        If Not pvTlsParseHandshakeServerHello(uCtx, baInput, lPos, lRecordProtocol, sError, eAlertCode) Then
                            GoTo QH
                        End If
                        If .HelloRetryRequest Then
                            '--- after HelloRetryRequest -> replace HandshakeMessages w/ 'synthetic handshake message'
                            baHandshakeHash = pvTlsArrayHash(.DigestAlgo, .HandshakeMessages, 0)
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
                    If .State = ucsTlsStateExpectServerHello And .HelloRetryRequest Then
                        .SendPos = pvTlsBuildClientHello(uCtx, .SendBuffer, .SendPos)
                    End If
                    If .State = ucsTlsStateExpectExtensions And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                        If Not pvTlsDeriveHandshakeSecrets(uCtx, sError, eAlertCode) Then
                            GoTo QH
                        End If
                    End If
                Case ucsTlsStateExpectExtensions
                    Select Case lMessageType
                    Case TLS_HANDSHAKE_TYPE_CERTIFICATE
                        If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, BlockSize:=lCertSize)
                                lPos = lPos + lCertSize ' skip RemoteCertReqContext
                            lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                        End If
                        Set .RemoteCertificates = New Collection
                        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=3, BlockSize:=lCertSize)
                            lCertEnd = lPos + lCertSize
                            Do While lPos < lCertEnd
                                lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=3, BlockSize:=lCertSize)
                                    lPos = pvReadArray(baInput, lPos, baCert, lCertSize)
                                    .RemoteCertificates.Add baCert
                                lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                                If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                                    lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lCertSize)
                                        '--- certificate extensions -> skip
                                        lPos = lPos + lCertSize
                                    lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                                End If
                            Loop
                        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                    Case TLS_HANDSHAKE_TYPE_CERTIFICATE_VERIFY
                        lPos = pvReadLong(baInput, lPos, lSignatureType, Size:=2)
                        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lCertSize)
                            lPos = pvReadArray(baInput, lPos, baSignature, lCertSize)
                        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                        If Not SearchCollection(.RemoteCertificates, 1, RetVal:=baCert) Then
                            sError = ERR_NO_SERVER_CERTIFICATE
                            eAlertCode = uscTlsAlertHandshakeFailure
                            GoTo QH
                        End If
                        baHandshakeHash = pvTlsArrayHash(.DigestAlgo, .HandshakeMessages, 0)
                        lVerifyPos = pvWriteString(baVerifyData, 0, Space$(64) & "TLS 1.3, server CertificateVerify" & Chr$(0))
                        lVerifyPos = pvWriteArray(baVerifyData, lVerifyPos, baHandshakeHash)
                        If Not pvTlsSignatureVerify(baCert, lSignatureType, baVerifyData, baSignature, sError, eAlertCode) Then
                            GoTo QH
                        End If
                    Case TLS_HANDSHAKE_TYPE_FINISHED
                        lPos = pvReadArray(baInput, lPos, baMessage, lMessageSize)
                        baHandshakeHash = pvTlsArrayHash(.DigestAlgo, .HandshakeMessages, 0)
                        baVerifyData = pvTlsHkdfExpandLabel(.DigestAlgo, .RemoteTrafficSecret, "finished", EmptyByteArray, .DigestSize)
                        baVerifyData = pvTlsHkdfExtract(.DigestAlgo, baVerifyData, baHandshakeHash)
                        If StrConv(baVerifyData, vbUnicode) <> StrConv(baMessage, vbUnicode) Then
                            sError = ERR_SERVER_HANDSHAKE_FAILED
                            eAlertCode = uscTlsAlertHandshakeFailure
                            GoTo QH
                        End If
                        .State = ucsTlsStatePostHandshake
                    Case TLS_HANDSHAKE_TYPE_SERVER_KEY_EXCHANGE
                        If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                            lSignPos = lPos
                            lPos = pvReadLong(baInput, lPos, lCurveType)
                            If lCurveType <> 3 Then '--- 3 = named_curve
                                sError = ERR_SERVER_HANDSHAKE_FAILED
                                eAlertCode = uscTlsAlertHandshakeFailure
                                GoTo QH
                            End If
                            lPos = pvReadLong(baInput, lPos, lNamedCurve, Size:=2)
                            If Not pvTlsSetupExchEccGroup(uCtx, lNamedCurve, sError, eAlertCode) Then
                                GoTo QH
                            End If
                            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, BlockSize:=lSignatureSize)
                                lPos = pvReadArray(baInput, lPos, .RemoteExchPublic, lSignatureSize)
                            lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                            lSignSize = lPos - lSignPos
                            '--- signature
                            lPos = pvReadLong(baInput, lPos, lSignatureType, Size:=2)
                            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lSignatureSize)
                                lPos = pvReadArray(baInput, lPos, baSignature, lSignatureSize)
                            lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                            If Not SearchCollection(.RemoteCertificates, 1, RetVal:=baCert) Then
                                sError = ERR_NO_SERVER_CERTIFICATE
                                eAlertCode = uscTlsAlertHandshakeFailure
                                GoTo QH
                            End If
                            lVerifyPos = pvWriteArray(baVerifyData, 0, .LocalExchRandom)
                            lVerifyPos = pvWriteArray(baVerifyData, lVerifyPos, .RemoteExchRandom)
                            lVerifyPos = pvWriteBuffer(baVerifyData, lVerifyPos, VarPtr(baInput(lSignPos)), lSignSize)
                            If Not pvTlsSignatureVerify(baCert, lSignatureType, baVerifyData, baSignature, sError, eAlertCode) Then
                                GoTo QH
                            End If
                            If Not pvTlsDeriveLegacySecrets(uCtx, sError, eAlertCode) Then
                                GoTo QH
                            End If
                        End If
                    Case TLS_HANDSHAKE_TYPE_SERVER_HELLO_DONE
                        If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                            .State = ucsTlsStateExpectServerFinished
                        End If
                        lPos = lPos + lMessageSize
                    Case Else
                        '--- do nothing
                        lPos = lPos + lMessageSize
                    End Select
                    pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baInput(lMessagePos)), lMessageSize + 4
                    '--- post-process ucsTlsStateExpectExtensions
                    If .State = ucsTlsStateExpectServerFinished And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                        If pvTlsCipherSuiteUseRsaCertificate(.CipherSuite) Then
                            If Not SearchCollection(.RemoteCertificates, 1, baCert) Then
                                sError = ERR_NO_SERVER_CERTIFICATE
                                eAlertCode = uscTlsAlertHandshakeFailure
                                GoTo QH
                            End If
                            If Not pvTlsSetupExchRsaCertificate(uCtx, baCert, sError, eAlertCode) Then
                                GoTo QH
                            End If
                            If Not pvTlsDeriveLegacySecrets(uCtx, sError, eAlertCode) Then
                                GoTo QH
                            End If
                        End If
                        .SendPos = pvTlsBuildClientLegacyKeyExchange(uCtx, .SendBuffer, .SendPos, sError, eAlertCode)
                        If LenB(sError) <> 0 Then
                            GoTo QH
                        End If
                    End If
                    If .State = ucsTlsStatePostHandshake And .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                        .SendPos = pvTlsBuildClientHandshakeFinished(uCtx, .SendBuffer, .SendPos, sError, eAlertCode)
                        If LenB(sError) <> 0 Then
                            GoTo QH
                        End If
                        If Not pvTlsDeriveApplicationSecrets(uCtx, sError, eAlertCode) Then
                            GoTo QH
                        End If
                        '--- not used past handshake
                        Erase .HandshakeMessages
                    End If
                Case ucsTlsStateExpectServerFinished
                    Select Case lMessageType
                    Case TLS_HANDSHAKE_TYPE_FINISHED
                        If .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS12 Then
                            lPos = pvReadArray(baInput, lPos, baMessage, lMessageSize)
                            baHandshakeHash = pvTlsArrayHash(.DigestAlgo, .HandshakeMessages, 0)
                            baVerifyData = pvTlsKdfLegacyPrf(.DigestAlgo, .MasterSecret, "server finished", baHandshakeHash, 12)
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
                        sError = Replace(Replace(ERR_UNEXPECTED_MSG_TYPE, "%1", "ucsTlsStateExpectServerFinished"), "%2", lMessageType)
                        eAlertCode = uscTlsAlertUnexpectedMessage
                        GoTo QH
                    End Select
                Case ucsTlsStateExpectClientHello
                    Select Case lMessageType
                    Case TLS_HANDSHAKE_TYPE_CLIENT_HELLO
                        If Not pvTlsParseHandshakeClientHello(uCtx, baInput, lPos, lRecordProtocol, sError, eAlertCode) Then
                            GoTo QH
                        End If
                        .State = ucsTlsStateExpectClientFinished
                    Case Else
                        sError = Replace(Replace(ERR_UNEXPECTED_MSG_TYPE, "%1", "ucsTlsStateExpectClientHello"), "%2", lMessageType)
                        eAlertCode = uscTlsAlertUnexpectedMessage
                        GoTo QH
                    End Select
                    pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baInput(lMessagePos)), lMessageSize + 4
                    '--- post-process ucsTlsStateExpectClientHello
                    If .State = ucsTlsStateExpectClientFinished Then
                        .SendPos = pvTlsBuildServerHello(uCtx, .SendBuffer, .SendPos)
                        If Not pvTlsDeriveHandshakeSecrets(uCtx, sError, eAlertCode) Then
                            GoTo QH
                        End If
                        .SendPos = pvTlsBuildServerHandshakeFinished(uCtx, .SendBuffer, .SendPos, sError, eAlertCode)
                        If LenB(sError) <> 0 Then
                            GoTo QH
                        End If
                    End If
                Case ucsTlsStateExpectClientFinished
                    Select Case lMessageType
                    Case TLS_HANDSHAKE_TYPE_FINISHED
                        lPos = pvReadArray(baInput, lPos, baMessage, lMessageSize)
                        baHandshakeHash = pvTlsArrayHash(.DigestAlgo, .HandshakeMessages, 0)
                        baVerifyData = pvTlsHkdfExpandLabel(.DigestAlgo, .RemoteTrafficSecret, "finished", EmptyByteArray, .DigestSize)
                        baVerifyData = pvTlsHkdfExtract(.DigestAlgo, baVerifyData, baHandshakeHash)
                        If StrConv(baVerifyData, vbUnicode) <> StrConv(baMessage, vbUnicode) Then
                            sError = ERR_SERVER_HANDSHAKE_FAILED
                            eAlertCode = uscTlsAlertHandshakeFailure
                            GoTo QH
                        End If
                        .State = ucsTlsStatePostHandshake
                    Case Else
                        sError = Replace(Replace(ERR_UNEXPECTED_MSG_TYPE, "%1", "ucsTlsStateExpectClientFinished"), "%2", lMessageType)
                        eAlertCode = uscTlsAlertUnexpectedMessage
                        GoTo QH
                    End Select
                    '--- post-process ucsTlsStateExpectClientFinished
                    If .State = ucsTlsStatePostHandshake Then
                        If Not pvTlsDeriveApplicationSecrets(uCtx, sError, eAlertCode) Then
                            GoTo QH
                        End If
                        .HandshakeMessages = vbNullString
                    End If
                Case ucsTlsStatePostHandshake
                    Select Case lMessageType
                    Case TLS_HANDSHAKE_TYPE_NEW_SESSION_TICKET
                        If Not .IsServer Then
                            '--- don't store tickets for now
                        End If
                    Case TLS_HANDSHAKE_TYPE_KEY_UPDATE
                        Debug.Print "Received TLS_HANDSHAKE_TYPE_KEY_UPDATE", Timer
                        If lMessageSize = 1 Then
                            lRequestUpdate = baInput(lPos)
                        Else
                            lRequestUpdate = -1
                        End If
                        If Not pvTlsDeriveKeyUpdate(uCtx, lRequestUpdate <> 0, sError, eAlertCode) Then
                            GoTo QH
                        End If
                        If lRequestUpdate <> 0 Then
                            '--- ack by TLS_HANDSHAKE_TYPE_KEY_UPDATE w/ update_not_requested(0)
                            If pvTlsBuildApplicationData(uCtx, baMessage, 0, pvArrayByte(TLS_HANDSHAKE_TYPE_KEY_UPDATE, 0, 0, 1, 0), -1, TLS_CONTENT_TYPE_APPDATA, sError, eAlertCode) = 0 Then
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
    '--- success
    pvTlsParseHandshake = True
QH:
End Function

Private Function pvTlsParseHandshakeServerHello(uCtx As UcsTlsContext, baInput() As Byte, lPos As Long, ByVal lRecordProtocol As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Static baHelloRetryRandom() As Byte
    Dim lSize           As Long
    Dim lEnd            As Long
    Dim lLegacyVersion  As Long
    Dim lCipherSuite    As Long
    Dim lLegacyCompress As Long
    Dim lExtType        As Long
    Dim lExtSize        As Long
    Dim lExchGroup      As Long
    Dim lBlockSize      As Long
    
    If pvArraySize(baHelloRetryRandom) = 0 Then
        baHelloRetryRandom = pvArrayByte(&HCF, &H21, &HAD, &H74, &HE5, &H9A, &H61, &H11, &HBE, &H1D, &H8C, &H2, &H1E, &H65, &HB8, &H91, &HC2, &HA2, &H11, &H16, &H7A, &HBB, &H8C, &H5E, &H7, &H9E, &H9, &HE2, &HC8, &HA8, &H33, &H9C)
    End If
    With uCtx
        .ProtocolVersion = lRecordProtocol
        lPos = pvReadLong(baInput, lPos, lLegacyVersion, Size:=2)
        lPos = pvReadArray(baInput, lPos, .RemoteExchRandom, TLS_HELLO_RANDOM_SIZE)
        If .HelloRetryRequest Then
            '--- clear HelloRetryRequest
            .HelloRetryRequest = False
            .HelloRetryCipherSuite = 0
            .HelloRetryExchGroup = 0
            Erase .HelloRetryCookie
        Else
            .HelloRetryRequest = (StrConv(.RemoteExchRandom, vbUnicode) = StrConv(baHelloRetryRandom, vbUnicode))
        End If
        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, BlockSize:=lSize)
            lPos = pvReadArray(baInput, lPos, .RemoteSessionID, lSize)
        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
        lPos = pvReadLong(baInput, lPos, lCipherSuite, Size:=2)
        If Not pvTlsSetupCipherSuite(uCtx, lCipherSuite, sError, eAlertCode) Then
            GoTo QH
        End If
        Debug.Print "Using " & pvTlsCipherSuiteName(.CipherSuite) & " from " & .RemoteHostName, Timer
        If .HelloRetryRequest Then
            .HelloRetryCipherSuite = lCipherSuite
        End If
        lPos = pvReadLong(baInput, lPos, lLegacyCompress)
        Debug.Assert lLegacyCompress = 0
        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lSize)
            lEnd = lPos + lSize
            Do While lPos < lEnd
                lPos = pvReadLong(baInput, lPos, lExtType, Size:=2)
                lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lExtSize)
                    Select Case lExtType
                    Case TLS_EXTENSION_TYPE_KEY_SHARE
                        .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13
                        If lExtSize < 2 Then
                            sError = ERR_INVALID_SIZE_KEY_SHARE
                            eAlertCode = uscTlsAlertDecodeError
                            GoTo QH
                        End If
                        lPos = pvReadLong(baInput, lPos, lExchGroup, Size:=2)
                        If Not pvTlsSetupExchEccGroup(uCtx, lExchGroup, sError, eAlertCode) Then
                            GoTo QH
                        End If
                        If .HelloRetryRequest Then
                            .HelloRetryExchGroup = lExchGroup
                        Else
                            If lExtSize <= 4 Then
                                sError = ERR_INVALID_SIZE_REMOTE_KEY
                                eAlertCode = uscTlsAlertDecodeError
                                GoTo QH
                            End If
                            lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lBlockSize)
                                lPos = pvReadArray(baInput, lPos, .RemoteExchPublic, lBlockSize)
                            lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                        End If
                    Case TLS_EXTENSION_TYPE_SUPPORTED_VERSIONS
                        If lExtSize <> 2 Then
                            sError = ERR_INVALID_SIZE_VERSIONS
                            eAlertCode = uscTlsAlertDecodeError
                            GoTo QH
                        End If
                        lPos = pvReadLong(baInput, lPos, .ProtocolVersion, Size:=2)
                    Case TLS_EXTENSION_TYPE_COOKIE
                        If Not .HelloRetryRequest Then
                            sError = ERR_COOKIE_NOT_ALLOWED
                            eAlertCode = uscTlsAlertIllegalParameter
                            GoTo QH
                        End If
                        lPos = pvReadArray(baInput, lPos, .HelloRetryCookie, lExtSize)
                    Case Else
                        lPos = lPos + lExtSize
                    End Select
                lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
            Loop
        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
    End With
    '--- success
    pvTlsParseHandshakeServerHello = True
QH:
End Function

Private Function pvTlsParseHandshakeClientHello(uCtx As UcsTlsContext, baInput() As Byte, lPos As Long, ByVal lRecordProtocol As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Dim lSize           As Long
    Dim lEnd            As Long
    Dim lLegacyVersion  As Long
    Dim lCipherSuite    As Long
    Dim lCipherPref     As Long
    Dim lLegacyCompress As Long
    Dim lExtType        As Long
    Dim lExtSize        As Long
    Dim lExtEnd         As Long
    Dim lExchGroup      As Long
    Dim lBlockSize      As Long
    Dim lBlockEnd       As Long
    Dim lProtocolVersion As Long
    Dim lSignatureType  As Long
    Dim cCipherPrefs    As Collection
    Dim vElem           As Variant
    Dim lIdx            As Long
    Dim baCert()        As Byte
    Dim sPubAlgoObjId   As String
    Dim lPubKeyLen      As Long
    
    Set cCipherPrefs = New Collection
    For Each vElem In pvTlsPrepareCipherSuitsOrder(ucsTlsSupportTls13)
        cCipherPrefs.Add cCipherPrefs.Count, "#" & vElem
    Next
    lCipherPref = 1000
    With uCtx
        If SearchCollection(.LocalCertificates, 1, RetVal:=baCert) Then
            Asn1DecodePublicKeyFromDer baCert, AlgoObjId:=sPubAlgoObjId, KeyLen:=lPubKeyLen
        End If
        .ProtocolVersion = lRecordProtocol
        lPos = pvReadLong(baInput, lPos, lLegacyVersion, Size:=2)
        lPos = pvReadArray(baInput, lPos, .RemoteExchRandom, TLS_HELLO_RANDOM_SIZE)
        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, BlockSize:=lSize)
            lPos = pvReadArray(baInput, lPos, .RemoteSessionID, lSize)
        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lSize)
            lEnd = lPos + lSize
            Do While lPos < lEnd
                lPos = pvReadLong(baInput, lPos, lIdx, Size:=2)
                If SearchCollection(cCipherPrefs, "#" & lIdx, RetVal:=vElem) Then
                    If vElem < lCipherPref Then
                        lCipherSuite = lIdx
                        lCipherPref = vElem
                    End If
                End If
            Loop
        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
        If lCipherSuite = 0 Then
            sError = ERR_NO_SUPPORTED_CIPHER_SUITE
            eAlertCode = uscTlsAlertHandshakeFailure
            GoTo QH
        End If
        If Not pvTlsSetupCipherSuite(uCtx, lCipherSuite, sError, eAlertCode) Then
            GoTo QH
        End If
        Debug.Print "Using " & pvTlsCipherSuiteName(.CipherSuite) & " from " & .RemoteHostName, Timer
        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack)
            lPos = pvReadLong(baInput, lPos, lLegacyCompress)
        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
        Debug.Assert lLegacyCompress = 0
        '--- extensions
        Set .RemoteExtensions = New Collection
        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lSize)
            lEnd = lPos + lSize
            Do While lPos < lEnd
                lPos = pvReadLong(baInput, lPos, lExtType, Size:=2)
                lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lExtSize)
                    lExtEnd = lPos + lExtSize
                    Select Case lExtType
                    Case TLS_EXTENSION_TYPE_KEY_SHARE
                        .ProtocolVersion = TLS_PROTOCOL_VERSION_TLS13
                        If lExtSize < 4 Then
                            sError = ERR_INVALID_SIZE_KEY_SHARE
                            eAlertCode = uscTlsAlertDecodeError
                            GoTo QH
                        End If
                        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lBlockSize)
                            lBlockEnd = lPos + lBlockSize
                            Do While lPos < lBlockEnd
                                lPos = pvReadLong(baInput, lPos, lExchGroup, Size:=2)
                                lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lBlockSize)
                                    If lExchGroup = TLS_GROUP_X25519 Then
                                        If lBlockSize <> TLS_X25519_KEY_SIZE Then
                                            sError = ERR_INVALID_REMOTE_KEY
                                            eAlertCode = uscTlsAlertIllegalParameter
                                            GoTo QH
                                        End If
                                        lPos = pvReadArray(baInput, lPos, .RemoteExchPublic, lBlockSize)
                                        If Not pvTlsSetupExchEccGroup(uCtx, lExchGroup, sError, eAlertCode) Then
                                            GoTo QH
                                        End If
                                    Else
                                        lPos = lPos + lBlockSize
                                    End If
                                lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                            Loop
                        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                    Case TLS_EXTENSION_TYPE_SIGNATURE_ALGORITHMS
                        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, Size:=2, BlockSize:=lBlockSize)
                            Do While lPos < lExtEnd
                                lPos = pvReadLong(baInput, lPos, lSignatureType, Size:=2)
                                Select Case lSignatureType
                                Case TLS_SIGNATURE_RSA_PKCS1_SHA1
                                    If sPubAlgoObjId = szOID_RSA_RSA And Not CryptoIsSupported(ucsTlsAlgoSignaturePkcsSha2) Then
                                        .LocalSignatureType = lSignatureType
                                    End If
                                Case TLS_SIGNATURE_RSA_PKCS1_SHA256, TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, _
                                        TLS_SIGNATURE_RSA_PKCS1_SHA384, TLS_SIGNATURE_RSA_PSS_RSAE_SHA384
                                    If sPubAlgoObjId = szOID_RSA_RSA And CryptoIsSupported(ucsTlsAlgoSignaturePkcsSha2) Then
                                        .LocalSignatureType = lSignatureType
                                    End If
                                Case TLS_SIGNATURE_RSA_PKCS1_SHA512, TLS_SIGNATURE_RSA_PSS_RSAE_SHA512
                                    '--- PSS w/ SHA512 fails on short key lengths
                                    If sPubAlgoObjId = szOID_RSA_RSA And lPubKeyLen > 1024 And CryptoIsSupported(ucsTlsAlgoSignaturePkcsSha2) Then
                                        .LocalSignatureType = lSignatureType
                                    End If
                                Case TLS_SIGNATURE_RSA_PSS_PSS_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA512
                                    If sPubAlgoObjId = szOID_RSA_SSA_PSS And CryptoIsSupported(ucsTlsAlgoSignaturePss) Then
                                        .LocalSignatureType = lSignatureType
                                    End If
                                Case TLS_SIGNATURE_ECDSA_SECP256R1_SHA256, TLS_SIGNATURE_ECDSA_SECP384R1_SHA384, TLS_SIGNATURE_ECDSA_SECP521R1_SHA512
                                    If sPubAlgoObjId = szOID_ECC_PUBLIC_KEY Then
                                        .LocalSignatureType = lSignatureType
                                    End If
                                End Select
                                If .LocalSignatureType <> 0 Then
                                    lPos = lExtEnd
                                End If
                            Loop
                            If .LocalSignatureType = 0 Then
                                sError = ERR_NEGOTIATE_SIGNATURE_FAILED
                                eAlertCode = uscTlsAlertHandshakeFailure
                                GoTo QH
                            End If
                        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                    Case TLS_EXTENSION_TYPE_SUPPORTED_VERSIONS
                        lPos = pvReadBeginOfBlock(baInput, lPos, .BlocksStack, BlockSize:=lBlockSize)
                            Do While lPos < lExtEnd
                                lPos = pvReadLong(baInput, lPos, lProtocolVersion, Size:=2)
                                If lProtocolVersion = TLS_PROTOCOL_VERSION_TLS13 Then
                                    lPos = lExtEnd
                                End If
                            Loop
                        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
                        If lProtocolVersion <> TLS_PROTOCOL_VERSION_TLS13 Then
                            sError = ERR_UNSUPPORTED_PROTOCOL
                            eAlertCode = uscTlsAlertProtocolVersion
                            GoTo QH
                        End If
                        .ProtocolVersion = lProtocolVersion
                    Case Else
                        lPos = lPos + lExtSize
                    End Select
                    If Not SearchCollection(.RemoteExtensions, "#" & lExtType) Then
                        .RemoteExtensions.Add lExtType, "#" & lExtType
                    End If
                lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
            Loop
        lPos = pvReadEndOfBlock(baInput, lPos, .BlocksStack)
    End With
    '--- success
    pvTlsParseHandshakeClientHello = True
QH:
End Function

Private Function pvTlsSetupExchEccGroup(uCtx As UcsTlsContext, ByVal lExchGroup As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    With uCtx
        If .ExchGroup <> lExchGroup Then
            .ExchGroup = lExchGroup
            Select Case lExchGroup
            Case TLS_GROUP_X25519
                .ExchAlgo = ucsTlsAlgoExchX25519
                If Not CryptoEccCurve25519MakeKey(.LocalExchPrivate, .LocalExchPublic) Then
                    sError = Replace(ERR_GENER_KEYPAIR_FAILED, "%1", "Curve25519")
                    eAlertCode = uscTlsAlertInternalError
                    GoTo QH
                End If
            Case TLS_GROUP_SECP256R1
                .ExchAlgo = ucsTlsAlgoExchSecp256r1
                If Not CryptoEccSecp256r1MakeKey(.LocalExchPrivate, .LocalExchPublic) Then
                    sError = Replace(ERR_GENER_KEYPAIR_FAILED, "%1", "secp256r1")
                    eAlertCode = uscTlsAlertInternalError
                    GoTo QH
                End If
            Case TLS_GROUP_SECP384R1
                .ExchAlgo = ucsTlsAlgoExchSecp384r1
                If Not CryptoEccSecp384r1MakeKey(.LocalExchPrivate, .LocalExchPublic) Then
                    sError = Replace(ERR_GENER_KEYPAIR_FAILED, "%1", "secp384r1")
                    eAlertCode = uscTlsAlertInternalError
                    GoTo QH
                End If
            Case Else
                sError = Replace(ERR_UNSUPPORTED_EXCH_GROUP, "%1", "0x" & Hex$(.ExchGroup))
                eAlertCode = uscTlsAlertInternalError
                GoTo QH
            End Select
        End If
    End With
    '--- success
    pvTlsSetupExchEccGroup = True
QH:
End Function

Private Function pvTlsSetupExchRsaCertificate(uCtx As UcsTlsContext, baCert() As Byte, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Dim uRsaCtx         As UcsRsaContextType
    
    On Error GoTo EH
    With uCtx
        .ExchAlgo = ucsTlsAlgoExchCertificate
        .LocalExchPrivate = pvTlsArrayRandom(TLS_HELLO_RANDOM_SIZE + TLS_HELLO_RANDOM_SIZE \ 2) '--- always 48
        pvWriteLong .LocalExchPrivate, 0, TLS_LOCAL_LEGACY_VERSION, Size:=2
        If Not CryptoRsaInitContext(uRsaCtx, EmptyByteArray, baCert, EmptyByteArray) Then
            sError = "CryptoRsaInitContext failed"
            eAlertCode = uscTlsAlertInternalError
            GoTo QH
        End If
        .LocalExchRsaEncrPriv = CryptoRsaEncrypt(uRsaCtx.hPubKey, .LocalExchPrivate)
    End With
    '--- success
    pvTlsSetupExchRsaCertificate = True
QH:
    If uRsaCtx.hProv <> 0 Then
        Call CryptoRsaTerminateContext(uRsaCtx)
    End If
    Exit Function
EH:
    sError = Trim$(Replace(Replace(Err.Description, vbCrLf, vbLf), vbLf, ". "))
    If Right$(sError, 1) = "." Then
        sError = Left$(sError, Len(sError) - 1)
    End If
    sError = sError & " in " & Err.Source
    eAlertCode = uscTlsAlertInternalError
End Function

Private Function pvTlsSetupCipherSuite(uCtx As UcsTlsContext, ByVal lCipherSuite As Long, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
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
    pvTlsSetupCipherSuite = True
QH:
End Function

Private Function pvTlsPrepareLegacyDecryptParams(uCtx As UcsTlsContext, baInput() As Byte, ByVal lRecordPos As Long, ByVal lRecordSize As Long, lPos As Long, lEnd As Long, baRemoteIV() As Byte, baAad() As Byte) As Boolean
    Dim lAadPos         As Long
    
    With uCtx
        lEnd = lPos + lRecordSize - .TagSize
        baRemoteIV = pvArrayXor(.RemoteTrafficIV, .RemoteTrafficSeqNo)
        If .IvDynamicSize > 0 Then '--- AES in TLS 1.2
            pvWriteBuffer baRemoteIV, .IvSize - .IvDynamicSize, VarPtr(baInput(lPos)), .IvDynamicSize
            lPos = lPos + .IvDynamicSize
        End If
        ReDim baAad(0 To LNG_LEGACY_AAD_SIZE - 1) As Byte
        Debug.Assert RedimStats("pvTlsPrepareLegacyDecryptParams.baAad", UBound(baAad) + 1)
        lAadPos = pvWriteLong(baAad, 0, 0, Size:=4)
        lAadPos = pvWriteLong(baAad, lAadPos, .RemoteTrafficSeqNo, Size:=4)
        lAadPos = pvWriteBuffer(baAad, lAadPos, VarPtr(baInput(lRecordPos)), 3)
        lAadPos = pvWriteLong(baAad, lAadPos, lEnd - lPos, Size:=2)
        Debug.Assert lAadPos = LNG_LEGACY_AAD_SIZE
    End With
    '--- success
    pvTlsPrepareLegacyDecryptParams = True
End Function

Private Function pvTlsPrepareCipherSuitsOrder(ByVal eFilter As UcsTlsLocalFeaturesEnum) As Collection
    Const PREF      As Long = &H1000
    Dim oRetVal     As Collection
    
    Set oRetVal = New Collection
    If (eFilter And ucsTlsSupportTls13) <> 0 Then
        If CryptoIsSupported(ucsTlsAlgoExchX25519) Then
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
        If CryptoIsSupported(ucsTlsAlgoExchSecp256r1) Then
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
        '--- no perfect forward secrecy -> least preferred
        If CryptoIsSupported(ucsTlsAlgoAeadAes128) Then
            oRetVal.Add TLS_CIPHER_SUITE_RSA_WITH_AES_128_GCM_SHA256
        End If
        If CryptoIsSupported(ucsTlsAlgoAeadAes256) Then
            oRetVal.Add TLS_CIPHER_SUITE_RSA_WITH_AES_256_GCM_SHA384
        End If
    End If
    Set pvTlsPrepareCipherSuitsOrder = oRetVal
End Function

Private Sub pvTlsSetLastError(uCtx As UcsTlsContext, sError As String, Optional ByVal AlertDesc As UcsTlsAlertDescriptionsEnum = -1)
    With uCtx
        .LastError = sError
        .LastAlertCode = AlertDesc
        If LenB(sError) = 0 Then
            Set .BlocksStack = Nothing
        Else
            If AlertDesc >= 0 Then
                .SendPos = pvTlsBuildAlert(uCtx, .SendBuffer, .SendPos, AlertDesc, TLS_ALERT_LEVEL_FATAL)
            End If
            .State = ucsTlsStateClosed
        End If
    End With
End Sub

'= HMAC-based key derivation functions ===================================

Private Function pvTlsDeriveHandshakeSecrets(uCtx As UcsTlsContext, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
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
        baHandshakeHash = pvTlsArrayHash(.DigestAlgo, .HandshakeMessages, 0)
        baEarlySecret = pvTlsHkdfExtract(.DigestAlgo, EmptyByteArray(.DigestSize), EmptyByteArray(.DigestSize))
        baEmptyHash = pvTlsArrayHash(.DigestAlgo, EmptyByteArray, 0)
        baDerivedSecret = pvTlsHkdfExpandLabel(.DigestAlgo, baEarlySecret, "derived", baEmptyHash, .DigestSize)
        baSharedSecret = pvTlsSharedSecret(.ExchAlgo, .LocalExchPrivate, .RemoteExchPublic)
        .HandshakeSecret = pvTlsHkdfExtract(.DigestAlgo, baDerivedSecret, baSharedSecret)
        .RemoteTrafficSecret = pvTlsHkdfExpandLabel(.DigestAlgo, .HandshakeSecret, IIf(.IsServer, "c", "s") & " hs traffic", baHandshakeHash, .DigestSize)
        .RemoteTrafficKey = pvTlsHkdfExpandLabel(.DigestAlgo, .RemoteTrafficSecret, "key", EmptyByteArray, .KeySize)
        .RemoteTrafficIV = pvTlsHkdfExpandLabel(.DigestAlgo, .RemoteTrafficSecret, "iv", EmptyByteArray, .IvSize)
        .RemoteTrafficSeqNo = 0
        .LocalTrafficSecret = pvTlsHkdfExpandLabel(.DigestAlgo, .HandshakeSecret, IIf(.IsServer, "s", "c") & " hs traffic", baHandshakeHash, .DigestSize)
        .LocalTrafficKey = pvTlsHkdfExpandLabel(.DigestAlgo, .LocalTrafficSecret, "key", EmptyByteArray, .KeySize)
        .LocalTrafficIV = pvTlsHkdfExpandLabel(.DigestAlgo, .LocalTrafficSecret, "iv", EmptyByteArray, .IvSize)
        .LocalTrafficSeqNo = 0
    End With
    '--- success
    pvTlsDeriveHandshakeSecrets = True
QH:
End Function

Private Function pvTlsDeriveApplicationSecrets(uCtx As UcsTlsContext, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Dim baHandshakeHash() As Byte
    Dim baEmptyHash()   As Byte
    Dim baDerivedSecret() As Byte
    
    With uCtx
        If pvArraySize(.HandshakeMessages) = 0 Then
            sError = ERR_NO_HANDSHAKE_MESSAGES
            eAlertCode = uscTlsAlertInternalError
            GoTo QH
        End If
        baHandshakeHash = pvTlsArrayHash(.DigestAlgo, .HandshakeMessages, 0)
        baEmptyHash = pvTlsArrayHash(.DigestAlgo, EmptyByteArray, 0)
        baDerivedSecret = pvTlsHkdfExpandLabel(.DigestAlgo, .HandshakeSecret, "derived", baEmptyHash, .DigestSize)
        .MasterSecret = pvTlsHkdfExtract(.DigestAlgo, baDerivedSecret, EmptyByteArray(.DigestSize))
        .RemoteTrafficSecret = pvTlsHkdfExpandLabel(.DigestAlgo, .MasterSecret, IIf(.IsServer, "c", "s") & " ap traffic", baHandshakeHash, .DigestSize)
        .RemoteTrafficKey = pvTlsHkdfExpandLabel(.DigestAlgo, .RemoteTrafficSecret, "key", EmptyByteArray, .KeySize)
        .RemoteTrafficIV = pvTlsHkdfExpandLabel(.DigestAlgo, .RemoteTrafficSecret, "iv", EmptyByteArray, .IvSize)
        .RemoteTrafficSeqNo = 0
        .LocalTrafficSecret = pvTlsHkdfExpandLabel(.DigestAlgo, .MasterSecret, IIf(.IsServer, "s", "c") & " ap traffic", baHandshakeHash, .DigestSize)
        .LocalTrafficKey = pvTlsHkdfExpandLabel(.DigestAlgo, .LocalTrafficSecret, "key", EmptyByteArray, .KeySize)
        .LocalTrafficIV = pvTlsHkdfExpandLabel(.DigestAlgo, .LocalTrafficSecret, "iv", EmptyByteArray, .IvSize)
        .LocalTrafficSeqNo = 0
    End With
    '--- success
    pvTlsDeriveApplicationSecrets = True
QH:
End Function

Private Function pvTlsDeriveKeyUpdate(uCtx As UcsTlsContext, ByVal bLocalUpdate As Boolean, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    With uCtx
        If pvArraySize(.RemoteTrafficSecret) = 0 Then
            sError = ERR_NO_PREV_REMOTE_SECRET
            eAlertCode = uscTlsAlertInternalError
            GoTo QH
        End If
        .RemoteTrafficSecret = pvTlsHkdfExpandLabel(.DigestAlgo, .RemoteTrafficSecret, "traffic upd", EmptyByteArray, .DigestSize)
        .RemoteTrafficKey = pvTlsHkdfExpandLabel(.DigestAlgo, .RemoteTrafficSecret, "key", EmptyByteArray, .KeySize)
        .RemoteTrafficIV = pvTlsHkdfExpandLabel(.DigestAlgo, .RemoteTrafficSecret, "iv", EmptyByteArray, .IvSize)
        .RemoteTrafficSeqNo = 0
        If bLocalUpdate Then
            If pvArraySize(.LocalTrafficSecret) = 0 Then
                sError = ERR_NO_PREV_LOCAL_SECRET
                eAlertCode = uscTlsAlertInternalError
                GoTo QH
            End If
            .LocalTrafficSecret = pvTlsHkdfExpandLabel(.DigestAlgo, .LocalTrafficSecret, "traffic upd", EmptyByteArray, .DigestSize)
            .LocalTrafficKey = pvTlsHkdfExpandLabel(.DigestAlgo, .LocalTrafficSecret, "key", EmptyByteArray, .KeySize)
            .LocalTrafficIV = pvTlsHkdfExpandLabel(.DigestAlgo, .LocalTrafficSecret, "iv", EmptyByteArray, .IvSize)
            .LocalTrafficSeqNo = 0
        End If
    End With
    '--- success
    pvTlsDeriveKeyUpdate = True
QH:
End Function

Private Function pvTlsHkdfExtract(ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baKey() As Byte, baInput() As Byte) As Byte()
    pvTlsHkdfExtract = pvTlsArrayHmac(eHash, baKey, baInput, 0)
End Function

Private Function pvTlsHkdfExpandLabel(ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baKey() As Byte, ByVal sLabel As String, baContext() As Byte, ByVal lSize As Long) As Byte()
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
        baLast = pvTlsArrayHmac(eHash, baKey, baInput, 0, Size:=lInputPos)
        lRetValPos = pvWriteArray(baRetVal, lRetValPos, baLast)
        lIdx = lIdx + 1
    Loop
    If UBound(baRetVal) <> lSize - 1 Then
        ReDim Preserve baRetVal(0 To lSize - 1) As Byte
        Debug.Assert RedimStats("pvTlsHkdfExpandLabel.baRetVal", UBound(baRetVal) + 1)
    End If
    pvTlsHkdfExpandLabel = baRetVal
End Function

'= legacy PRF-based key derivation functions =============================

Private Function pvTlsDeriveLegacySecrets(uCtx As UcsTlsContext, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Dim baPreMasterSecret() As Byte
    Dim baRandom()      As Byte
    Dim baExpanded()    As Byte
    Dim lPos            As Long
    
    With uCtx
        If pvArraySize(.RemoteExchRandom) = 0 Then
            sError = ERR_NO_REMOTE_RANDOM
            eAlertCode = uscTlsAlertInternalError
            GoTo QH
        End If
        Debug.Assert pvArraySize(.LocalExchRandom) = TLS_HELLO_RANDOM_SIZE
        Debug.Assert pvArraySize(.RemoteExchRandom) = TLS_HELLO_RANDOM_SIZE
        baPreMasterSecret = pvTlsSharedSecret(.ExchAlgo, .LocalExchPrivate, .RemoteExchPublic)
        ReDim baRandom(0 To pvArraySize(.LocalExchRandom) + pvArraySize(.RemoteExchRandom) - 1) As Byte
        Debug.Assert RedimStats("pvTlsDeriveLegacySecrets.baRandom", UBound(baRandom) + 1)
        lPos = pvWriteArray(baRandom, 0, .LocalExchRandom)
        lPos = pvWriteArray(baRandom, lPos, .RemoteExchRandom)
        .MasterSecret = pvTlsKdfLegacyPrf(.DigestAlgo, baPreMasterSecret, "master secret", baRandom, TLS_HELLO_RANDOM_SIZE + TLS_HELLO_RANDOM_SIZE \ 2) '--- always 48
        lPos = pvWriteArray(baRandom, 0, .RemoteExchRandom)
        lPos = pvWriteArray(baRandom, lPos, .LocalExchRandom)
        baExpanded = pvTlsKdfLegacyPrf(.DigestAlgo, .MasterSecret, "key expansion", baRandom, 2 * (.MacSize + .KeySize + .IvSize))
        lPos = pvReadArray(baExpanded, 0, EmptyByteArray, .MacSize) '--- LocalMacKey not used w/ AEAD
        lPos = pvReadArray(baExpanded, lPos, EmptyByteArray, .MacSize) '--- RemoteMacKey not used w/ AEAD
        lPos = pvReadArray(baExpanded, lPos, .LocalTrafficKey, .KeySize)
        lPos = pvReadArray(baExpanded, lPos, .RemoteTrafficKey, .KeySize)
        lPos = pvReadArray(baExpanded, lPos, .LocalTrafficIV, .IvSize - .IvDynamicSize)
        pvWriteArray .LocalTrafficIV, .IvSize - .IvDynamicSize, pvTlsArrayRandom(.IvDynamicSize)
        lPos = pvReadArray(baExpanded, lPos, .RemoteTrafficIV, .IvSize - .IvDynamicSize)
        pvWriteArray .RemoteTrafficIV, .IvSize - .IvDynamicSize, pvTlsArrayRandom(.IvDynamicSize)
    End With
    '--- success
    pvTlsDeriveLegacySecrets = True
QH:
End Function

Private Function pvTlsKdfLegacyPrf(ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baSecret() As Byte, ByVal sLabel As String, baContext() As Byte, ByVal lSize As Long) As Byte()
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
        baLast = pvTlsArrayHmac(eHash, baSecret, baLast, 0)
        lInputPos = pvWriteArray(baInput, 0, baLast)
        lInputPos = pvWriteArray(baInput, lInputPos, baSeed)
        baHmac = pvTlsArrayHmac(eHash, baSecret, baInput, 0, Size:=lInputPos)
        lRetValPos = pvWriteArray(baRetVal, lRetValPos, baHmac)
    Loop
    If lRetValPos <> lSize Then
        ReDim Preserve baRetVal(0 To lSize - 1) As Byte
        Debug.Assert RedimStats("pvTlsKdfLegacyPrf.baRetVal", UBound(baRetVal) + 1)
    End If
    pvTlsKdfLegacyPrf = baRetVal
End Function

'= crypto wrappers =======================================================

Private Function pvTlsArrayRandom(ByVal lSize As Long) As Byte()
    Dim baRetVal()      As Byte
    
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
        Debug.Assert RedimStats("pvTlsArrayRandom.baRetVal", UBound(baRetVal) + 1)
        CryptoRandomBytes VarPtr(baRetVal(0)), lSize
    End If
    pvTlsArrayRandom = baRetVal
End Function

Private Function pvTlsArrayHash(ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baInput() As Byte, ByVal lPos As Long, Optional ByVal Size As Long = -1) As Byte()
    Select Case eHash
    Case 0
        pvReadArray baInput, lPos, pvTlsArrayHash, Size
    Case ucsTlsAlgoDigestSha256
        pvTlsArrayHash = CryptoHashSha256(baInput, lPos, Size)
    Case ucsTlsAlgoDigestSha384
        pvTlsArrayHash = CryptoHashSha384(baInput, lPos, Size)
    Case ucsTlsAlgoDigestSha512
        pvTlsArrayHash = CryptoHashSha512(baInput, lPos, Size)
    Case Else
        Err.Raise vbObjectError, "pvTlsArrayHash", "Unsupported hash type " & eHash
    End Select
End Function

Private Function pvTlsArrayHmac(ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baKey() As Byte, baInput() As Byte, ByVal lPos As Long, Optional ByVal Size As Long = -1) As Byte()
    Select Case eHash
    Case ucsTlsAlgoDigestSha256
        pvTlsArrayHmac = CryptoHmacSha256(baKey, baInput, lPos, Size)
    Case ucsTlsAlgoDigestSha384
        pvTlsArrayHmac = CryptoHmacSha384(baKey, baInput, lPos, Size)
    Case Else
        Err.Raise vbObjectError, "pvTlsArrayHmac", "Unsupported hash type " & eHash
    End Select
End Function

Private Function pvTlsAeadDecrypt(ByVal eAead As UcsTlsCryptoAlgorithmsEnum, baRemoteIV() As Byte, baRemoteKey() As Byte, baAad() As Byte, ByVal lAadPos As Long, ByVal lAdSize As Long, baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Select Case eAead
    Case ucsTlsAlgoAeadChacha20Poly1305
        If Not CryptoAeadChacha20Poly1305Decrypt(baRemoteIV, baRemoteKey, baAad, lAadPos, lAdSize, baBuffer, lPos, lSize) Then
            GoTo QH
        End If
    Case ucsTlsAlgoAeadAes128, ucsTlsAlgoAeadAes256
        If Not CryptoAeadAesGcmDecrypt(baRemoteIV, baRemoteKey, baAad, lAadPos, lAdSize, baBuffer, lPos, lSize) Then
            GoTo QH
        End If
    Case Else
        Err.Raise vbObjectError, "pvTlsAeadDecrypt", "Unsupported AEAD type " & eAead
    End Select
    '--- success
    pvTlsAeadDecrypt = True
QH:
End Function

Private Function pvTlsAeadEncrypt(ByVal eAead As UcsTlsCryptoAlgorithmsEnum, baLocalIV() As Byte, baLocalKey() As Byte, baAad() As Byte, ByVal lAadPos As Long, ByVal lAdSize As Long, baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Select Case eAead
    Case ucsTlsAlgoAeadChacha20Poly1305
        If Not CryptoAeadChacha20Poly1305Encrypt(baLocalIV, baLocalKey, baAad, lAadPos, lAdSize, baBuffer, lPos, lSize) Then
            GoTo QH
        End If
    Case ucsTlsAlgoAeadAes128, ucsTlsAlgoAeadAes256
        If Not CryptoAeadAesGcmEncrypt(baLocalIV, baLocalKey, baAad, lAadPos, lAdSize, baBuffer, lPos, lSize) Then
            GoTo QH
        End If
    Case Else
        Err.Raise vbObjectError, "pvTlsAeadEncrypt", "Unsupported AEAD type " & eAead
    End Select
    '--- success
    pvTlsAeadEncrypt = True
QH:
End Function

Private Function pvTlsSharedSecret(ByVal eKeyX As UcsTlsCryptoAlgorithmsEnum, baPriv() As Byte, baPub() As Byte) As Byte()
    Dim baRetVal()      As Byte
    
    Select Case eKeyX
    Case ucsTlsAlgoExchX25519
        baRetVal = CryptoEccCurve25519SharedSecret(baPriv, baPub)
    Case ucsTlsAlgoExchSecp256r1
        baRetVal = CryptoEccSecp256r1SharedSecret(baPriv, baPub)
    Case ucsTlsAlgoExchSecp384r1
        baRetVal = CryptoEccSecp384r1SharedSecret(baPriv, baPub)
    Case ucsTlsAlgoExchCertificate
        baRetVal = baPriv
    Case Else
        Err.Raise vbObjectError, "pvTlsSharedSecret", "Unsupported exchange curve " & eKeyX
    End Select
    pvTlsSharedSecret = baRetVal
End Function

Private Function pvTlsCipherSuiteName(ByVal lCipherSuite As Long) As String
    Select Case lCipherSuite
    Case TLS_CIPHER_SUITE_AES_128_GCM_SHA256
        pvTlsCipherSuiteName = "TLS_AES_128_GCM_SHA256"
    Case TLS_CIPHER_SUITE_AES_256_GCM_SHA384
        pvTlsCipherSuiteName = "TLS_AES_256_GCM_SHA384"
    Case TLS_CIPHER_SUITE_CHACHA20_POLY1305_SHA256
        pvTlsCipherSuiteName = "TLS_CHACHA20_POLY1305_SHA256"
    Case TLS_CIPHER_SUITE_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256
        pvTlsCipherSuiteName = "ECDHE-ECDSA-AES128-GCM-SHA256"
    Case TLS_CIPHER_SUITE_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384
        pvTlsCipherSuiteName = "ECDHE-ECDSA-AES256-GCM-SHA384"
    Case TLS_CIPHER_SUITE_ECDHE_RSA_WITH_AES_128_GCM_SHA256
        pvTlsCipherSuiteName = "ECDHE-RSA-AES128-GCM-SHA256"
    Case TLS_CIPHER_SUITE_ECDHE_RSA_WITH_AES_256_GCM_SHA384
        pvTlsCipherSuiteName = "ECDHE-RSA-AES256-GCM-SHA384"
    Case TLS_CIPHER_SUITE_ECDHE_RSA_WITH_CHACHA20_POLY1305_SHA256
        pvTlsCipherSuiteName = "ECDHE-RSA-CHACHA20-POLY1305"
    Case TLS_CIPHER_SUITE_ECDHE_ECDSA_WITH_CHACHA20_POLY1305_SHA256
        pvTlsCipherSuiteName = "ECDHE-ECDSA-CHACHA20-POLY1305"
    Case TLS_CIPHER_SUITE_RSA_WITH_AES_128_GCM_SHA256
        pvTlsCipherSuiteName = "AES128-GCM-SHA256"
    Case TLS_CIPHER_SUITE_RSA_WITH_AES_256_GCM_SHA384
        pvTlsCipherSuiteName = "AES256-GCM-SHA384"
    Case Else
        pvTlsCipherSuiteName = Replace(STR_UNKNOWN, "%1", "0x" & Hex$(lCipherSuite))
    End Select
End Function

Private Function pvTlsCipherSuiteUseRsaCertificate(ByVal lCipherSuite As Long) As Boolean
    Select Case lCipherSuite
    Case TLS_CIPHER_SUITE_RSA_WITH_AES_128_GCM_SHA256, TLS_CIPHER_SUITE_RSA_WITH_AES_256_GCM_SHA384
        pvTlsCipherSuiteUseRsaCertificate = True
    End Select
End Function

Private Function pvTlsSignatureTypeName(ByVal lSignatureType As Long) As String
    Select Case lSignatureType
    Case TLS_SIGNATURE_RSA_PKCS1_SHA1
        pvTlsSignatureTypeName = "RSA_PKCS1_SHA1"
    Case TLS_SIGNATURE_RSA_PKCS1_SHA256
        pvTlsSignatureTypeName = "RSA_PKCS1_SHA256"
    Case TLS_SIGNATURE_RSA_PKCS1_SHA384
        pvTlsSignatureTypeName = "RSA_PKCS1_SHA384"
    Case TLS_SIGNATURE_RSA_PKCS1_SHA512
        pvTlsSignatureTypeName = "RSA_PKCS1_SHA512"
    Case TLS_SIGNATURE_ECDSA_SECP256R1_SHA256
        pvTlsSignatureTypeName = "ECDSA_SECP256R1_SHA256"
    Case TLS_SIGNATURE_ECDSA_SECP384R1_SHA384
        pvTlsSignatureTypeName = "ECDSA_SECP384R1_SHA384"
    Case TLS_SIGNATURE_ECDSA_SECP521R1_SHA512
        pvTlsSignatureTypeName = "ECDSA_SECP521R1_SHA512"
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256
        pvTlsSignatureTypeName = "RSA_PSS_RSAE_SHA256"
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA384
        pvTlsSignatureTypeName = "RSA_PSS_RSAE_SHA384"
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA512
        pvTlsSignatureTypeName = "RSA_PSS_RSAE_SHA512"
    Case TLS_SIGNATURE_RSA_PSS_PSS_SHA256
        pvTlsSignatureTypeName = "RSA_PSS_PSS_SHA256"
    Case TLS_SIGNATURE_RSA_PSS_PSS_SHA384
        pvTlsSignatureTypeName = "RSA_PSS_PSS_SHA384"
    Case TLS_SIGNATURE_RSA_PSS_PSS_SHA512
        pvTlsSignatureTypeName = "RSA_PSS_PSS_SHA512"
    Case Else
        pvTlsSignatureTypeName = Replace(STR_UNKNOWN, "%1", "0x" & Hex$(lSignatureType))
    End Select
End Function

Private Function pvTlsSignatureSign(baPrivKey() As Byte, ByVal lSignatureType As Long, baVerifyData() As Byte, baSignature() As Byte, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Dim baVerifyHash()      As Byte
    Dim uRsaCtx             As UcsRsaContextType
    
    Debug.Print "Signing with " & pvTlsSignatureTypeName(lSignatureType) & " signature", Timer
    Select Case lSignatureType
    Case TLS_SIGNATURE_RSA_PKCS1_SHA1, TLS_SIGNATURE_RSA_PKCS1_SHA256, TLS_SIGNATURE_RSA_PKCS1_SHA384, TLS_SIGNATURE_RSA_PKCS1_SHA512
        If Not CryptoRsaInitContext(uRsaCtx, baPrivKey, EmptyByteArray, EmptyByteArray, lSignatureType) Then
            sError = "CryptoRsaInitContext failed"
            eAlertCode = uscTlsAlertInternalError
            GoTo QH
        End If
        baSignature = CryptoRsaSign(uRsaCtx, baVerifyData)
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, TLS_SIGNATURE_RSA_PSS_RSAE_SHA512, _
            TLS_SIGNATURE_RSA_PSS_PSS_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA512
        baSignature = CryptoRsaPssSign(baPrivKey, baVerifyData, lSignatureType)
    Case Else
        baVerifyHash = pvTlsArrayHash(pvTlsSignatureDigestAlgo(lSignatureType), baVerifyData, 0)
        Select Case lSignatureType
        Case TLS_SIGNATURE_ECDSA_SECP256R1_SHA256
            baSignature = CryptoEccSecp256r1Sign(baPrivKey, baVerifyHash)
            baSignature = pvAsn1EncodeSignatureToDer(baSignature, TLS_SECP256R1_KEY_SIZE)
        Case TLS_SIGNATURE_ECDSA_SECP384R1_SHA384
            baSignature = CryptoEccSecp384r1Sign(baPrivKey, baVerifyHash)
            baSignature = pvAsn1EncodeSignatureToDer(baSignature, TLS_SECP384R1_KEY_SIZE)
        Case Else
            sError = Replace(ERR_UNSUPPORTED_SIGNATURE_TYPE, "%1", "0x" & Hex$(lSignatureType))
            eAlertCode = uscTlsAlertInternalError
            GoTo QH
        End Select
    End Select
    If pvArraySize(baSignature) = 0 Then
        sError = Replace(ERR_SIGNATURE_FAILED, "%1", pvTlsSignatureTypeName(lSignatureType))
        eAlertCode = uscTlsAlertInternalError
        GoTo QH
    End If
    pvTlsSignatureSign = True
QH:
    If uRsaCtx.hProv <> 0 Then
        Call CryptoRsaTerminateContext(uRsaCtx)
    End If
End Function

Private Function pvTlsSignatureVerify(baCert() As Byte, ByVal lSignatureType As Long, baVerifyData() As Byte, baSignature() As Byte, sError As String, eAlertCode As UcsTlsAlertDescriptionsEnum) As Boolean
    Dim uRsaCtx         As UcsRsaContextType
    Dim baPubKey()      As Byte
    Dim baVerifyHash()  As Byte
    Dim baPlainSig()    As Byte
    Dim sPubKeyObjId    As String
    Dim lCurveSize      As Long
    Dim bSkip           As Boolean
    Dim baTemp()        As Byte
    Dim bDeprecated     As Boolean
    
    Select Case lSignatureType
    Case TLS_SIGNATURE_RSA_PKCS1_SHA1, TLS_SIGNATURE_RSA_PKCS1_SHA256, TLS_SIGNATURE_RSA_PKCS1_SHA384, TLS_SIGNATURE_RSA_PKCS1_SHA512
        If Not CryptoRsaInitContext(uRsaCtx, EmptyByteArray, baCert, EmptyByteArray, lSignatureType) Then
            sError = "CryptoRsaInitContext failed"
            eAlertCode = uscTlsAlertInternalError
            GoTo QH
        End If
        If Not CryptoRsaVerify(uRsaCtx, baVerifyData, baSignature) Then
            GoTo InvalidSignature
        End If
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, TLS_SIGNATURE_RSA_PSS_RSAE_SHA512, _
            TLS_SIGNATURE_RSA_PSS_PSS_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA512
        If Not CryptoIsSupported(ucsTlsAlgoSignaturePss) Then
            bSkip = True
        Else
            If Not CryptoRsaPssVerify(baCert, baVerifyData, baSignature, lSignatureType) Then
                GoTo InvalidSignature
            End If
        End If
    Case TLS_SIGNATURE_ECDSA_SECP256R1_SHA256, TLS_SIGNATURE_ECDSA_SECP384R1_SHA384, TLS_SIGNATURE_ECDSA_SECP521R1_SHA512
        baPubKey = Asn1DecodePublicKeyFromDer(baCert, sPubKeyObjId)
        If sPubKeyObjId <> szOID_ECC_PUBLIC_KEY Then
            sError = Replace(ERR_UNSUPPORTED_PUBLIC_KEY, "%1", sPubKeyObjId)
            eAlertCode = uscTlsAlertHandshakeFailure
            GoTo QH
        End If
        baVerifyHash = pvTlsArrayHash(pvTlsSignatureDigestAlgo(lSignatureType), baVerifyData, 0)
        lCurveSize = UBound(baPubKey) \ 2
        baPlainSig = pvAsn1DecodeSignatureFromDer(baSignature, lCurveSize)
        If pvArraySize(baPlainSig) = 0 Then
            GoTo InvalidSignature
        End If
        If UBound(baVerifyHash) + 1 < lCurveSize Then
            '--- note: when hash size is less than curve size must left-pad w/ zeros (right-align hash) -> deprecated
            '---       incl. ECDSA_SECP384R1_SHA256 only
            baTemp = baVerifyHash
            ReDim baVerifyHash(0 To lCurveSize - 1) As Byte
            Debug.Assert RedimStats("pvTlsSignatureVerify.baVerifyHash", UBound(baVerifyHash) + 1)
            Call CopyMemory(baVerifyHash(lCurveSize - UBound(baTemp) - 1), baTemp(0), UBound(baTemp) + 1)
            bDeprecated = True
        ElseIf UBound(baVerifyHash) + 1 > lCurveSize Then
            '--- note: when hash size is above curve size the excess is ignored -> deprecated
            '---       incl. ECDSA_SECP256R1_SHA384, ECDSA_SECP256R1_SHA512 and ECDSA_SECP384R1_SHA512
            bDeprecated = True
        End If
        Select Case lCurveSize
        Case TLS_SECP256R1_KEY_SIZE
            If Not CryptoEccSecp256r1Verify(baPubKey, baVerifyHash, baPlainSig) Then
                GoTo InvalidSignature
            End If
        Case TLS_SECP384R1_KEY_SIZE
            If Not CryptoEccSecp384r1Verify(baPubKey, baVerifyHash, baPlainSig) Then
                GoTo InvalidSignature
            End If
        Case Else
            sError = Replace(ERR_UNSUPPORTED_CURVE_SIZE, "%1", lCurveSize)
            eAlertCode = uscTlsAlertHandshakeFailure
            GoTo QH
        End Select
    Case Else
        sError = Replace(ERR_UNSUPPORTED_SIGNATURE_TYPE, "%1", "0x" & Hex$(lSignatureType))
        eAlertCode = uscTlsAlertInternalError
        GoTo QH
    End Select
    '--- success
    pvTlsSignatureVerify = True
QH:
    Debug.Print IIf(pvTlsSignatureVerify, IIf(bSkip, "Skipping ", IIf(bDeprecated, "Deprecated ", "Valid ")), "Invalid ") & pvTlsSignatureTypeName(lSignatureType) & " signature" & IIf(bDeprecated, " (lCurveSize=" & lCurveSize & " from server's public key)", vbNullString), Timer
    If uRsaCtx.hProv <> 0 Then
        Call CryptoRsaTerminateContext(uRsaCtx)
    End If
    Exit Function
InvalidSignature:
    sError = ERR_INVALID_SIGNATURE
    eAlertCode = uscTlsAlertHandshakeFailure
    GoTo QH
End Function

Private Function pvTlsSignatureDigestAlgo(ByVal lSignatureType As Long) As UcsTlsCryptoAlgorithmsEnum
    Select Case lSignatureType And &HFF
    Case 1, 2, 3 '--- 1 - RSA, 2 - DSA, 3 - ECDSA
        Select Case lSignatureType \ &H100
        '--- Skipping: 1 - MD-5, 2 - SHA-1, 3 - SHA-224
        Case 4
            pvTlsSignatureDigestAlgo = ucsTlsAlgoDigestSha256
        Case 5
            pvTlsSignatureDigestAlgo = ucsTlsAlgoDigestSha384
        Case 6
            pvTlsSignatureDigestAlgo = ucsTlsAlgoDigestSha512
        End Select
    Case Else
        '--- 8 - Intrinsic for TLS 1.3
        Select Case lSignatureType
        Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA256
            pvTlsSignatureDigestAlgo = ucsTlsAlgoDigestSha256
        Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA384
            pvTlsSignatureDigestAlgo = ucsTlsAlgoDigestSha384
        Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA512, TLS_SIGNATURE_RSA_PSS_PSS_SHA512
            pvTlsSignatureDigestAlgo = ucsTlsAlgoDigestSha512
        End Select
    End Select
End Function

Private Function pvAsn1DecodeSignatureFromDer(baDerSig() As Byte, ByVal lCurveSize As Long) As Byte()
    Dim baRetVal()      As Byte
    Dim lType           As Long
    Dim lPos            As Long
    Dim lSize           As Long
    Dim cStack          As Collection
    Dim baTemp()        As Byte
    
    ReDim baRetVal(0 To 63) As Byte
    Debug.Assert RedimStats("pvAsn1DecodeSignatureFromDer.baRetVal", UBound(baRetVal) + 1)
    '--- ECDSA-Sig-Value ::= SEQUENCE { r INTEGER, s INTEGER }
    lPos = pvReadLong(baDerSig, 0, lType)
    If lType <> LNG_ANS1_TYPE_SEQUENCE Then
        GoTo QH
    End If
    lPos = pvReadBeginOfBlock(baDerSig, lPos, cStack)
        lPos = pvReadLong(baDerSig, lPos, lType)
        If lType <> LNG_ANS1_TYPE_INTEGER Then
            GoTo QH
        End If
        lPos = pvReadLong(baDerSig, lPos, lSize)
        lPos = pvReadArray(baDerSig, lPos, baTemp, lSize)
        If lSize <= lCurveSize Then
            pvWriteArray baRetVal, lCurveSize - lSize, baTemp
        Else
            pvWriteBuffer baRetVal, 0, VarPtr(baTemp(lSize - lCurveSize)), lCurveSize
        End If
        lPos = pvReadLong(baDerSig, lPos, lType)
        If lType <> LNG_ANS1_TYPE_INTEGER Then
            GoTo QH
        End If
        lPos = pvReadLong(baDerSig, lPos, lSize)
        lPos = pvReadArray(baDerSig, lPos, baTemp, lSize)
        If lSize <= lCurveSize Then
            pvWriteArray baRetVal, lCurveSize + lCurveSize - lSize, baTemp
        Else
            pvWriteBuffer baRetVal, lCurveSize, VarPtr(baTemp(lSize - lCurveSize)), lCurveSize
        End If
    lPos = pvReadEndOfBlock(baDerSig, lPos, cStack)
    pvAsn1DecodeSignatureFromDer = baRetVal
QH:
End Function

Private Function pvAsn1EncodeSignatureToDer(baPlainSig() As Byte, ByVal lPartSize As Long) As Byte()
    Dim baRetVal()      As Byte
    Dim lPos            As Long
    Dim cStack          As Collection
    Dim lStart          As Long
    
    lPos = pvWriteLong(baRetVal, lPos, LNG_ANS1_TYPE_SEQUENCE)
    lPos = pvWriteBeginOfBlock(baRetVal, lPos, cStack)
        lPos = pvWriteLong(baRetVal, lPos, LNG_ANS1_TYPE_INTEGER)
        lPos = pvWriteBeginOfBlock(baRetVal, lPos, cStack)
            For lStart = 0 To lPartSize - 1
                If baPlainSig(lStart) <> 0 Then
                    Exit For
                End If
            Next
            If (baPlainSig(lStart) And &H80) <> 0 Then
                lPos = pvWriteLong(baRetVal, lPos, 0)
            End If
            lPos = pvWriteBuffer(baRetVal, lPos, VarPtr(baPlainSig(lStart)), lPartSize - lStart)
        lPos = pvWriteEndOfBlock(baRetVal, lPos, cStack)
        lPos = pvWriteLong(baRetVal, lPos, LNG_ANS1_TYPE_INTEGER)
        lPos = pvWriteBeginOfBlock(baRetVal, lPos, cStack)
            For lStart = 0 To lPartSize - 1
                If baPlainSig(lPartSize + lStart) <> 0 Then
                    Exit For
                End If
            Next
            If (baPlainSig(lPartSize + lStart) And &H80) <> 0 Then
                lPos = pvWriteLong(baRetVal, lPos, 0)
            End If
            lPos = pvWriteBuffer(baRetVal, lPos, VarPtr(baPlainSig(lPartSize + lStart)), lPartSize - lStart)
        lPos = pvWriteEndOfBlock(baRetVal, lPos, cStack)
    lPos = pvWriteEndOfBlock(baRetVal, lPos, cStack)
    pvAsn1EncodeSignatureToDer = baRetVal
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
    
    lSize = pvArraySize(baSrc)
    If lSize > 0 Then
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
        Debug.Assert RedimStats("pvWriteBuffer.baBuffer", UBound(baBuffer) + 1)
    ElseIf UBound(baBuffer) < lPos + lSize - 1 Then
        ReDim Preserve baBuffer(0 To lPos + lSize - 1) As Byte
        Debug.Assert RedimStats("pvWriteBuffer.baBuffer", UBound(baBuffer) + 1)
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
        Debug.Assert RedimStats("pvReadArray.baDest", UBound(baDest) + 1)
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

Private Function pvArraySize(baArray() As Byte) As Long
    Dim lPtr            As Long
    
    '--- peek long at ArrPtr(baArray)
    Call CopyMemory(lPtr, ByVal ArrPtr(baArray), 4)
    If lPtr <> 0 Then
        pvArraySize = UBound(baArray) + 1
    End If
End Function

Private Function pvArrayXor(baArray() As Byte, ByVal lSeqNo As Long) As Byte()
    Dim baRetVal()      As Byte
    Dim lIdx            As Long
    
    baRetVal = baArray
    lIdx = pvArraySize(baRetVal)
    Do While lSeqNo <> 0 And lIdx > 0
        lIdx = lIdx - 1
        baRetVal(lIdx) = baRetVal(lIdx) Xor (lSeqNo And &HFF)
        lSeqNo = (lSeqNo And -&H100) \ &H100
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

Private Function pvArrayByte(ParamArray A() As Variant) As Byte()
    Dim baRetVal()      As Byte
    Dim vElem           As Variant
    Dim lIdx            As Long
    
    If UBound(A) >= 0 Then
        ReDim baRetVal(0 To UBound(A)) As Byte
        Debug.Assert RedimStats("pvArrayByte.baRetVal", UBound(baRetVal) + 1)
        For Each vElem In A
            baRetVal(lIdx) = vElem And &HFF
            lIdx = lIdx + 1
        Next
    End If
    pvArrayByte = baRetVal
End Function

'= global helpers ========================================================

Private Function EmptyByteArray(Optional ByVal Size As Long) As Byte()
    Dim baRetVal()      As Byte
    
    If Size > 0 Then
        ReDim baRetVal(0 To Size - 1) As Byte
        Debug.Assert RedimStats("EmptyByteArray.baRetVal", UBound(baRetVal) + 1)
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
        Debug.Assert RedimStats("SplitOrReindex.vResult", 0)
        For lIdx = 0 To UBound(vTemp) Step 2
            vResult(vTemp(lIdx)) = vTemp(lIdx + 1)
        Next
        SplitOrReindex = vResult
    End If
End Function

Private Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function
