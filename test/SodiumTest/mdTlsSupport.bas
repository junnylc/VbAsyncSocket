Attribute VB_Name = "mdTlsSupport"
'=========================================================================
'
' Based on RFC 8446 at https://tools.ietf.org/html/rfc8446
'   and illustrated traffic-dump at https://tls13.ulfheim.net/
'
' More TLS 1.3 implementations at https://github.com/h2o/picotls
'   and https://github.com/openssl/openssl
'
' List of resources at https://github.com/tlswg/tls13-spec/wiki/Implementations
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
'Private Const TLS_ALERT_LEVEL_WARNING                   As Long = 1
Private Const TLS_ALERT_LEVEL_FATAL                     As Long = 2
Private Const TLS_SHA256_DIGEST_SIZE                    As Long = 32
Private Const TLS_SHA384_DIGEST_SIZE                    As Long = 48
Private Const TLS_X25519_KEY_SIZE                       As Long = 32
Private Const TLS_MAX_PLAINTEXT_RECORD_SIZE             As Long = 16384
Private Const TLS_MAX_ENCRYPTED_RECORD_SIZE             As Long = (16384 + 256)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
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

Private Const STR_VL_ALERTS             As String = "0|Close notify|10|Unexpected message|20|Bad record mac|40|Handshake failure|42|Bad certificate|44|Certificate revoked|45|Certificate expired|46|Certificate unknown|47|Illegal parameter|48|Unknown CA|50|Decode error|51|Decrypt error|70|Protocol version|80|Internal error|90|User canceled|109|Missing extension|112|Unrecognized name|116|Certificate required|120|No application protocol"

Public Enum UcsTlsStatesEnum
    ucsTlsStateHandshakeStart
    ucsTlsStateExpectServerHello
    ucsTlsStateExpectEncryptedExtensions
    ucsTlsStatePostHandshake
End Enum

Public Enum UcsTlsCryptoAlgorithmsEnum
    '--- key exchange
    ucsTlsAlgoKeyX25519
    '--- authenticated encryption w/ additional data
    ucsTlsAlgoAeadAes128
    ucsTlsAlgoAeadChacha20Poly1305
    ucsTlsAlgoAeadAes256
    '--- digest
    ucsTlsAlgoDigestSha256
    ucsTlsAlgoDigestSha384
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
    ServerName                  As String
    LegacySessionID()           As Byte '--- not used
    ClientRandom()              As Byte
    ClientPrivate()             As Byte
    ClientPublic()              As Byte
    ServerRandom()              As Byte '--- not used
    ServerPublic()              As Byte
    ServerSupportedVersion      As Long '--- not used
    ServerCertificate()         As Byte
    '--- keys
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
    '--- state
    State                       As UcsTlsStatesEnum
    LastError                   As String
    LastAlertDesc               As UcsTlsAlertDescriptionsEnum
    '--- crypto settings
    KxAlgo                      As UcsTlsCryptoAlgorithmsEnum
    SecretSize                  As Long
    CipherSuite                 As Long
    AeadAlgo                    As UcsTlsCryptoAlgorithmsEnum
    KeySize                     As Long
    IvSize                      As Long
    TagSize                     As Long
    DigestAlgo                  As UcsTlsCryptoAlgorithmsEnum
    DigestSize                  As Long
    '--- buffers
    RecvBuffer()                As Byte
    RecvPos                     As Long
    DecrBuffer()                As Byte
    DecrPos                     As Long
    SendBuffer()                As Byte
    SendPos                     As Long
    MessBuffer()                As Byte
    MessPos                     As Long
    MessSize                    As Long
End Type

'=========================================================================
' Methods
'=========================================================================

Public Function TlsInitClient(sServerName As String) As UcsTlsContext
    Dim uCtx            As UcsTlsContext
    
    On Error GoTo EH
    With uCtx
        pvSetLastError uCtx, vbNullString
        .ServerName = sServerName
'        .LegacySessionID = pvCryptoRandomBytes(32)
        '--- setup key exchange ephemeral priv/pub keys
        .KxAlgo = ucsTlsAlgoKeyX25519
        .SecretSize = TLS_X25519_KEY_SIZE
        .ClientRandom = pvCryptoRandomBytes(.SecretSize)
        If .KxAlgo = ucsTlsAlgoKeyX25519 Then
            .ClientPrivate = pvCryptoRandomBytes(TLS_X25519_KEY_SIZE)
            '--- fix some issues w/ specific privkeys
            .ClientPrivate(0) = .ClientPrivate(0) And 248
            .ClientPrivate(UBound(.ClientPrivate)) = (.ClientPrivate(UBound(.ClientPrivate)) And 127) Or 64
            ReDim .ClientPublic(0 To TLS_X25519_KEY_SIZE - 1) As Byte
            Call crypto_scalarmult_curve25519_base(.ClientPublic(0), .ClientPrivate(0))
        Else
            Err.Raise vbObjectError, "TlsInitClient", "Unsupported key-exchange type " & .KxAlgo
        End If
    End With
QH:
    TlsInitClient = uCtx
    Exit Function
EH:
    pvSetLastError uCtx, Err.Description
    Resume QH
End Function

Public Function TlsHandshake(uCtx As UcsTlsContext, baInput() As Byte, lSize As Long, baOutput() As Byte, lPos As Long, bComplete As Boolean) As Boolean
    On Error GoTo EH
    With uCtx
        pvSetLastError uCtx, vbNullString
        '--- swap-in
        pvArraySwap .SendBuffer, .SendPos, baOutput, lPos
        If .State = ucsTlsStateHandshakeStart Then
            .SendPos = pvSendClientHello(uCtx, .SendBuffer, .SendPos)
            .State = ucsTlsStateExpectServerHello
        Else
            If lSize < 0 Then
                lSize = pvArraySize(baInput)
            End If
            If Not pvHandleInput(uCtx, baInput, lSize, .LastError) Then
                GoTo QH
            End If
        End If
        bComplete = (.State >= ucsTlsStatePostHandshake)
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

Public Function TlsSend(uCtx As UcsTlsContext, baData() As Byte, lSize As Long, baOutput() As Byte, lPos As Long) As Boolean
    On Error GoTo EH
    With uCtx
        pvSetLastError uCtx, vbNullString
        '--- swap-in
        pvArraySwap .SendBuffer, .SendPos, baOutput, lPos
        If lSize < 0 Then
            lSize = pvArraySize(baData)
        End If
        .SendPos = pvSendClientApplicationData(uCtx, .SendBuffer, .SendPos, baData, lSize, .LastError)
        If LenB(.LastError) <> 0 Then
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

Public Function TlsReceive(uCtx As UcsTlsContext, baInput() As Byte, lSize As Long, baOutput() As Byte, lPos As Long) As Boolean
    On Error GoTo EH
    With uCtx
        pvSetLastError uCtx, vbNullString
        '--- swap-in
        pvArraySwap .DecrBuffer, .DecrPos, baOutput, lPos
        If lSize < 0 Then
            lSize = pvArraySize(baInput)
        End If
        If Not pvHandleInput(uCtx, baInput, lSize, .LastError) Then
            GoTo QH
        End If
        '--- success
        TlsReceive = True
QH:
        '--- swap-out
        pvArraySwap baOutput, lPos, .DecrBuffer, .DecrPos
    End With
    Exit Function
EH:
    pvSetLastError uCtx, Err.Description
    Resume QH
End Function

Public Function TlsGetLastError(uCtx As UcsTlsContext) As String
    TlsGetLastError = uCtx.LastError
    If uCtx.LastAlertDesc <> -1 Then
        TlsGetLastError = IIf(LenB(TlsGetLastError) <> 0, TlsGetLastError & ": ", vbNullString) & TlsGetLastAlert(uCtx)
    End If
End Function

Public Function TlsGetLastAlert(uCtx As UcsTlsContext, Optional AlertCode As UcsTlsAlertDescriptionsEnum) As String
    Static vTexts       As Variant
    
    AlertCode = uCtx.LastAlertDesc
    If AlertCode >= 0 Then
        If IsEmpty(vTexts) Then
            vTexts = SplitOrReindex(STR_VL_ALERTS, "|")
        End If
        If AlertCode <= UBound(vTexts) Then
            TlsGetLastAlert = vTexts(AlertCode)
        End If
        If LenB(TlsGetLastAlert) = 0 Then
            TlsGetLastAlert = "Unknown (" & AlertCode & ")"
        End If
    End If
End Function

'= private ===============================================================

Private Function pvSendClientHello(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long) As Long
    Dim lMessagePos     As Long
    Dim cBlocks         As Collection
    
    With uCtx
        '--- Record Header
        lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
        lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
        lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks, Size:=2)
            lMessagePos = lPos
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
'                    lPos = pvWriteLong(baOutput, lPos, TLS_CIPHER_SUITE_AES_256_GCM_SHA384, Size:=2)
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
        pvWriteBuffer .HandshakeMessages, 0, VarPtr(baOutput(lMessagePos)), lPos - lMessagePos
    End With
    pvSendClientHello = lPos
End Function

Private Function pvSendClientHandshakeFinished(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long, sError As String) As Long
    Dim lRecordPos      As Long
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim cBlocks         As Collection
    Dim baVerifyData()  As Byte
    Dim baClientIV()    As Byte
    Dim baHandshakeHash() As Byte
    
    With uCtx
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
                baHandshakeHash = pvCryptoHash(.DigestAlgo, .HandshakeMessages, 0)
                baVerifyData = pvHkdfExpand(.DigestAlgo, .ClientTrafficSecret, "finished", EmptyByteArray, .DigestSize)
                baVerifyData = pvHkdfExtract(.DigestAlgo, baVerifyData, baHandshakeHash)
                lPos = pvWriteArray(baOutput, lPos, baVerifyData)
            lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
            lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
            lMessageSize = lPos - lMessagePos
            lPos = pvWriteReserved(baOutput, lPos, .TagSize)
        lPos = pvWriteEndOfBlock(baOutput, lPos, cBlocks)
        Debug.Assert cBlocks.Count = 0
        baClientIV = pvArrayXor(.ClientTrafficIV, .ClientTrafficSeqNo)
        If pvCryptoEncrypt(.AeadAlgo, baClientIV, .ClientTrafficKey, baOutput, lRecordPos, 5, baOutput, lMessagePos, lMessageSize) Then
            .ClientTrafficSeqNo = .ClientTrafficSeqNo + 1
        Else
            sError = "Encryption failed"
            GoTo QH
        End If
    End With
    pvSendClientHandshakeFinished = lPos
QH:
End Function

Private Function pvSendClientApplicationData(uCtx As UcsTlsContext, baOutput() As Byte, ByVal lPos As Long, baData() As Byte, ByVal lSize As Long, sError As String) As Long
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim cBlocks         As Collection
    Dim baClientIV()    As Byte
    
    '--- Record Header
    lPos = pvWriteLong(baOutput, lPos, TLS_CONTENT_TYPE_APPDATA)
    lPos = pvWriteLong(baOutput, lPos, TLS_RECORD_VERSION, Size:=2)
    lPos = pvWriteBeginOfBlock(baOutput, lPos, cBlocks, Size:=2)
        lMessagePos = lPos
        If lSize > 0 Then
            lPos = pvWriteBuffer(baOutput, lPos, VarPtr(baData(0)), lSize)
        End If
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
    pvSendClientApplicationData = lPos
QH:
End Function

Private Function pvHandleInput(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, sError As String) As Boolean
    Dim lRecvPos        As Long
    Dim lRecvSize       As Long
    
    If lSize > 0 Then
    With uCtx
        .RecvPos = pvWriteBuffer(.RecvBuffer, .RecvPos, VarPtr(baInput(0)), lSize)
        lRecvPos = pvHandleRecord(uCtx, .RecvBuffer, .RecvPos, sError)
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
    pvHandleInput = True
QH:
End Function

Private Function pvHandleRecord(uCtx As UcsTlsContext, baInput() As Byte, ByVal lSize As Long, sError As String) As Long
    Dim lRecordPos      As Long
    Dim lRecordSize     As Long
    Dim lRecordType     As Long
    Dim cBlocks         As Collection
    Dim baServerIV()    As Byte
    Dim lPos            As Long
    Dim lEnd            As Long
    
    With uCtx
    Do While lPos + 6 <= lSize
        lRecordPos = lPos
        lPos = pvReadLong(baInput, lPos, lRecordType)
        lPos = pvReadLong(baInput, lPos, TLS_CLIENT_LEGACY_VERSION, Size:=2)
        lPos = pvReadBeginOfBlock(baInput, lPos, cBlocks, Size:=2, BlockSize:=lRecordSize)
            If lRecordSize > IIf(lRecordType = TLS_CONTENT_TYPE_APPDATA, TLS_MAX_ENCRYPTED_RECORD_SIZE, TLS_MAX_PLAINTEXT_RECORD_SIZE) Then
                sError = "Record size too big"
                GoTo QH
            End If
            If lPos + lRecordSize > lSize Then
                '--- back off and bail out early
                lPos = lRecordPos
                Exit Do
            End If
            Select Case lRecordType
            Case TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC
                lPos = lPos + lRecordSize
            Case TLS_CONTENT_TYPE_ALERT
HandleAlertContent:
                If lRecordSize >= 2 Then
                    .LastAlertDesc = baInput(lPos + 1)
                    If baInput(lPos) = TLS_ALERT_LEVEL_FATAL Then
                        sError = "Fatal alert"
                        GoTo QH
                    End If
                End If
                lPos = lPos + lRecordSize
            Case TLS_CONTENT_TYPE_HANDSHAKE
                lPos = pvHandleHandshakeContent(uCtx, baInput, lPos, lPos + lRecordSize, sError)
                If LenB(sError) <> 0 Then
                    GoTo QH
                End If
            Case TLS_CONTENT_TYPE_APPDATA
                baServerIV = pvArrayXor(.ServerTrafficIV, .ServerTrafficSeqNo)
                If pvCryptoDecrypt(.AeadAlgo, baServerIV, .ServerTrafficKey, baInput, lRecordPos, 5, baInput, lPos, lRecordSize) Then
                    .ServerTrafficSeqNo = .ServerTrafficSeqNo + 1
                ElseIf pvArraySize(.PrevServerTrafficIV) <> 0 Then
                    baServerIV = pvArrayXor(.PrevServerTrafficIV, .PrevServerTrafficSeqNo)
                    If pvCryptoDecrypt(.AeadAlgo, baServerIV, .PrevServerTrafficKey, baInput, lRecordPos, 5, baInput, lPos, lRecordSize) Then
                        .PrevServerTrafficSeqNo = .PrevServerTrafficSeqNo + 1
                    Else
                        sError = "pvCryptoDecrypt w/ PrevServerTrafficIV failed"
                        GoTo QH
                    End If
                Else
                    sError = "pvCryptoDecrypt w/ ServerTrafficIV failed"
                    GoTo QH
                End If
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
                    If .MessSize > 0 Then
                        .MessSize = pvWriteBuffer(.MessBuffer, .MessSize, VarPtr(baInput(lPos)), lEnd - lPos)
                        .MessPos = pvHandleHandshakeContent(uCtx, .MessBuffer, .MessPos, .MessSize, sError)
                        If LenB(sError) <> 0 Then
                            GoTo QH
                        End If
                        If .MessPos >= .MessSize Then
                            Erase .MessBuffer
                            .MessSize = 0
                            .MessPos = 0
                        End If
                    Else
                        lPos = pvHandleHandshakeContent(uCtx, baInput, lPos, lEnd, sError)
                        If LenB(sError) <> 0 Then
                            GoTo QH
                        End If
                        If lPos < lEnd Then
                            .MessSize = pvWriteBuffer(.MessBuffer, .MessSize, VarPtr(baInput(lPos)), lEnd - lPos)
                            .MessPos = 0
                        End If
                    End If
                Case TLS_CONTENT_TYPE_APPDATA
                    Select Case .State
                    Case ucsTlsStatePostHandshake
                        .DecrPos = pvWriteBuffer(.DecrBuffer, .DecrPos, VarPtr(baInput(lPos)), lEnd - lPos)
                    Case Else
                        sError = "Invalid state for appdata content (" & .State & ")"
                        GoTo QH
                    End Select
                End Select
                '--- note: skip zero padding too
                lPos = lRecordPos + lRecordSize + 5
            Case Else
                sError = "Unexpected record type (" & lRecordType & ")"
                GoTo QH
            End Select
        lPos = pvReadEndOfBlock(baInput, lPos, cBlocks)
    Loop
    End With
    '--- success
    pvHandleRecord = lPos
QH:
End Function

Private Function pvHandleHandshakeContent(uCtx As UcsTlsContext, baInput() As Byte, ByVal lPos As Long, ByVal lEnd As Long, sError As String) As Long
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim lMessageType    As Long
    Dim cBlocks         As Collection
    Dim baMessage()     As Byte
    Dim baHandshakeHash() As Byte
    Dim baVerifyData()  As Byte
    Dim lVerifyPos      As Long
    Dim lRequestUpdate  As Long
    
    With uCtx
        Do While lPos < lEnd
            lMessagePos = lPos
            lPos = pvReadLong(baInput, lPos, lMessageType)
            lPos = pvReadBeginOfBlock(baInput, lPos, cBlocks, Size:=3, BlockSize:=lMessageSize)
                If lPos + lMessageSize > lEnd Then
                    '--- back off and bail out early
                    lPos = lMessagePos
                    Exit Do
                End If
                Select Case .State
                Case ucsTlsStateExpectServerHello
                    Select Case lMessageType
                    Case TLS_HANDSHAKE_TYPE_SERVER_HELLO
                        lPos = pvReadArray(baInput, lPos, baMessage, lMessageSize)
                        If Not pvHandleHandshakeServerHello(uCtx, baMessage, sError) Then
                            GoTo QH
                        End If
                        pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baInput(lMessagePos)), lMessageSize + 4
                        If Not pvDeriveHandshakeSecrets(uCtx, sError) Then
                            GoTo QH
                        End If
                        .State = ucsTlsStateExpectEncryptedExtensions
                    Case Else
                        sError = "Unexpected message type for ucsTlsStateExpectServerHello (lMessageType=" & lMessageType & ")"
                        GoTo QH
                    End Select
                Case ucsTlsStateExpectEncryptedExtensions
                    Select Case lMessageType
                    Case TLS_HANDSHAKE_TYPE_CERTIFICATE
                        pvReadArray baInput, lPos, .ServerCertificate, lMessageSize
                    Case TLS_HANDSHAKE_TYPE_CERTIFICATE_VERIFY
                        baHandshakeHash = pvCryptoHash(.DigestAlgo, .HandshakeMessages, 0)
                        lVerifyPos = pvWriteString(baVerifyData, 0, Space$(64) & "TLS 1.3, server CertificateVerify" & Chr$(0))
                        lVerifyPos = pvWriteArray(baVerifyData, lVerifyPos, baHandshakeHash)
                        '--- ToDo: verify .ServerCertificate signature
                        '--- ShellExecute("openssl x509 -pubkey -noout -in server.crt > server.pub")
                    Case TLS_HANDSHAKE_TYPE_FINISHED
                        pvReadArray baInput, lPos, baMessage, lMessageSize
                        baHandshakeHash = pvCryptoHash(.DigestAlgo, .HandshakeMessages, 0)
                        baVerifyData = pvHkdfExpand(.DigestAlgo, .ServerTrafficSecret, "finished", EmptyByteArray, .DigestSize)
                        baVerifyData = pvHkdfExtract(.DigestAlgo, baVerifyData, baHandshakeHash)
                        Debug.Assert StrConv(baVerifyData, vbUnicode) = StrConv(baMessage, vbUnicode)
                        If StrConv(baVerifyData, vbUnicode) <> StrConv(baMessage, vbUnicode) Then
                            sError = "Server Handshake verification failed"
                            GoTo QH
                        End If
                        .State = ucsTlsStatePostHandshake
                    Case TLS_HANDSHAKE_TYPE_ENCRYPTED_EXTENSIONS, TLS_HANDSHAKE_TYPE_COMPRESSED_CERTIFICATE
                        '--- do nothing
                    Case Else
                        sError = "Unexpected message type for ucsTlsStateExpectEncryptedExtensions (lMessageType=" & lMessageType & ")"
                        GoTo QH
                    End Select
                    lPos = lPos + lMessageSize
                    pvWriteBuffer .HandshakeMessages, pvArraySize(.HandshakeMessages), VarPtr(baInput(lMessagePos)), lMessageSize + 4
                    If .State = ucsTlsStatePostHandshake Then
                        .SendPos = pvSendClientHandshakeFinished(uCtx, .SendBuffer, .SendPos, sError)
                        If LenB(sError) <> 0 Then
                            GoTo QH
                        End If
                        If Not pvDeriveApplicationSecrets(uCtx, sError) Then
                            GoTo QH
                        End If
                    End If
                Case ucsTlsStatePostHandshake
                    Select Case lMessageType
                    Case TLS_HANDSHAKE_TYPE_NEW_SESSION_TICKET
                        '--- don't store tickets for now
                    Case TLS_HANDSHAKE_TYPE_KEY_UPDATE
                        Debug.Print "TLS_HANDSHAKE_TYPE_KEY_UPDATE"
                        If lMessageSize = 1 Then
                            lRequestUpdate = baInput(lPos)
                        Else
                            lRequestUpdate = -1
                        End If
                        If Not pvDeriveKeyUpdate(uCtx, lRequestUpdate <> 0, sError) Then
                            GoTo QH
                        End If
                        If lRequestUpdate <> 0 Then
                            '--- ack by TLS_HANDSHAKE_TYPE_KEY_UPDATE w/ update_not_requested(0)
                            If pvSendClientApplicationData(uCtx, baMessage, 0, FromHex("1800000100"), -1, sError) = 0 Then
                                GoTo QH
                            End If
                            .SendPos = pvWriteArray(.SendBuffer, .SendPos, baMessage)
                        End If
                    Case Else
                        sError = "Unexpected message type for ucsTlsStatePostHandshake (lMessageType=" & lMessageType & ")"
                        GoTo QH
                    End Select
                    lPos = lPos + lMessageSize
                Case Else
                    sError = "Invalid state for handshake content (" & .State & ")"
                    GoTo QH
                End Select
            lPos = pvReadEndOfBlock(baInput, lPos, cBlocks)
        Loop
    End With
    '--- success
    pvHandleHandshakeContent = lPos
QH:
End Function

Private Function pvHandleHandshakeServerHello(uCtx As UcsTlsContext, baMessage() As Byte, sError As String) As Boolean
    Dim lPos            As Long
    Dim lBlockSize      As Long
    Dim cBlocks         As Collection
    Dim lLegacyVersion  As Long
    Dim lLegacyCompress As Long
    Dim lExtType        As Long
    Dim lExchangeGroup  As Long
    
    With uCtx
        lPos = pvReadLong(baMessage, lPos, lLegacyVersion, Size:=2)
        lPos = pvReadArray(baMessage, lPos, .ServerRandom, .SecretSize)
        lPos = pvReadBeginOfBlock(baMessage, lPos, cBlocks, BlockSize:=lBlockSize)
            lPos = pvReadArray(baMessage, lPos, .LegacySessionID, lBlockSize)
        lPos = pvReadEndOfBlock(baMessage, lPos, cBlocks)
        lPos = pvReadLong(baMessage, lPos, .CipherSuite, Size:=2)
        Select Case .CipherSuite
        Case TLS_CIPHER_SUITE_CHACHA20_POLY1305_SHA256
            .AeadAlgo = ucsTlsAlgoAeadChacha20Poly1305
            .KeySize = TLS_CHACHA20_KEY_SIZE
            .IvSize = TLS_CHACHA20POLY1305_IV_SIZE
            .TagSize = TLS_CHACHA20POLY1305_TAG_SIZE
            .DigestAlgo = ucsTlsAlgoDigestSha256
            .DigestSize = TLS_SHA256_DIGEST_SIZE
        Case TLS_CIPHER_SUITE_AES_256_GCM_SHA384
            .AeadAlgo = ucsTlsAlgoAeadChacha20Poly1305
            .KeySize = TLS_AES256_KEY_SIZE
            .IvSize = TLS_AESGCM_IV_SIZE
            .TagSize = TLS_AESGCM_TAG_SIZE
            .DigestAlgo = ucsTlsAlgoDigestSha384
            .DigestSize = TLS_SHA384_DIGEST_SIZE
        Case Else
            sError = "Unsupported cipher suite (0x" & Hex$(.CipherSuite) & ")"
            GoTo QH
        End Select
        lPos = pvReadLong(baMessage, lPos, lLegacyCompress)
        Debug.Assert lLegacyCompress = 0
        lPos = pvReadBeginOfBlock(baMessage, lPos, cBlocks, Size:=2)
            Do While lPos < cBlocks.Item(cBlocks.Count)
                lPos = pvReadLong(baMessage, lPos, lExtType, Size:=2)
                lPos = pvReadBeginOfBlock(baMessage, lPos, cBlocks, Size:=2, BlockSize:=lBlockSize)
                    Select Case lExtType
                    Case TLS_EXTENSION_TYPE_KEY_SHARE
                        lPos = pvReadLong(baMessage, lPos, lExchangeGroup, Size:=2)
                        Debug.Assert lExchangeGroup = TLS_GROUP_X25519
                        lPos = pvReadBeginOfBlock(baMessage, lPos, cBlocks, Size:=2, BlockSize:=lBlockSize)
                            Debug.Assert lBlockSize = .KeySize
                            If lBlockSize <> .KeySize Then
                                sError = "Invalid server key size"
                                GoTo QH
                            End If
                            lPos = pvReadArray(baMessage, lPos, .ServerPublic, .KeySize)
                        lPos = pvReadEndOfBlock(baMessage, lPos, cBlocks)
                    Case TLS_EXTENSION_TYPE_SUPPORTED_VERSIONS
                        If lBlockSize >= 2 Then
                            Call pvReadLong(baMessage, lPos, .ServerSupportedVersion, Size:=2)
                        End If
                        lPos = lPos + lBlockSize
                    Case Else
                        lPos = lPos + lBlockSize
                    End Select
                lPos = pvReadEndOfBlock(baMessage, lPos, cBlocks)
            Loop
        lPos = pvReadEndOfBlock(baMessage, lPos, cBlocks)
    End With
    '--- success
    pvHandleHandshakeServerHello = True
QH:
End Function

Private Sub pvSetLastError(uCtx As UcsTlsContext, sError As String, Optional ByVal AlertDesc As UcsTlsAlertDescriptionsEnum = -1)
    uCtx.LastError = sError
    uCtx.LastAlertDesc = AlertDesc
End Sub

'= HMAC-based key derivation functions ===================================

Private Function pvDeriveHandshakeSecrets(uCtx As UcsTlsContext, sError As String) As Boolean
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
        '--- for ucsTlsAlgoDigestSha256 always 33AD0A1C607EC03B09E6CD9893680CE210ADF300AA1F2660E1B22E10F170F92A
        baEarlySecret = pvHkdfExtract(.DigestAlgo, EmptyByteArray(.DigestSize), EmptyByteArray(.KeySize))
        '--- for ucsTlsAlgoDigestSha256 always  E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855
        baEmptyHash = pvCryptoHash(.DigestAlgo, EmptyByteArray, 0)
        '--- for ucsTlsAlgoDigestSha256 always 6F2615A108C702C5678F54FC9DBAB69716C076189C48250CEBEAC3576C3611BA
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

Private Function pvDeriveApplicationSecrets(uCtx As UcsTlsContext, sError As String) As Boolean
    Dim baHandshakeHash() As Byte
    Dim baEmptyHash()   As Byte
    Dim baDerivedSecret() As Byte
    
    With uCtx
        If pvArraySize(.HandshakeMessages) = 0 Then
            sError = "Missing handshake records"
            GoTo QH
        End If
        baHandshakeHash = pvCryptoHash(.DigestAlgo, .HandshakeMessages, 0)
        '--- for ucsTlsAlgoDigestSha256 always E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855
        baEmptyHash = pvCryptoHash(.DigestAlgo, EmptyByteArray, 0)
        '--- for ucsTlsAlgoDigestSha256 always 6F2615A108C702C5678F54FC9DBAB69716C076189C48250CEBEAC3576C3611BA
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

Private Function pvDeriveKeyUpdate(uCtx As UcsTlsContext, ByVal bUpdateClient As Boolean, sError As String) As Boolean
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

Private Function pvHkdfExtract(ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baSalt() As Byte, baInput() As Byte) As Byte()
    pvHkdfExtract = pvCryptoHmac(eHash, baSalt, baInput, 0)
End Function

Private Function pvHkdfExpand(ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baSalt() As Byte, ByVal sLabel As String, baContext() As Byte, ByVal lSize As Long) As Byte()
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

Private Function pvCryptoDecrypt(eAead As UcsTlsCryptoAlgorithmsEnum, baServerIV() As Byte, baServerKey() As Byte, baAd() As Byte, ByVal lAdPos As Long, ByVal lAdSize As Long, baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    If eAead = ucsTlsAlgoAeadChacha20Poly1305 Then
        Debug.Assert pvArraySize(baServerIV) = TLS_CHACHA20POLY1305_IV_SIZE
        Debug.Assert pvArraySize(baServerKey) = TLS_CHACHA20_KEY_SIZE
        Debug.Assert pvArraySize(baBuffer) >= lPos + lSize
        If crypto_aead_chacha20poly1305_ietf_decrypt(baBuffer(lPos), ByVal 0, 0, baBuffer(lPos), lSize, 0, baAd(lAdPos), lAdSize, 0, baServerIV(0), baServerKey(0)) <> 0 Then
            GoTo QH
        End If
    Else
        Err.Raise vbObjectError, "pvCryptoDecrypt", "Unsupported aead type " & eAead
    End If
    '--- success
    pvCryptoDecrypt = True
QH:
End Function

Private Function pvCryptoEncrypt(eAead As UcsTlsCryptoAlgorithmsEnum, baClientIV() As Byte, baClientKey() As Byte, baAd() As Byte, ByVal lAdPos As Long, ByVal lAdSize As Long, baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    If eAead = ucsTlsAlgoAeadChacha20Poly1305 Then
        Debug.Assert pvArraySize(baClientIV) = TLS_CHACHA20POLY1305_IV_SIZE
        Debug.Assert pvArraySize(baClientKey) = TLS_CHACHA20_KEY_SIZE
        Debug.Assert pvArraySize(baBuffer) >= lPos + lSize + TLS_CHACHA20POLY1305_TAG_SIZE
        If crypto_aead_chacha20poly1305_ietf_encrypt(baBuffer(lPos), ByVal 0, baBuffer(lPos), lSize, 0, baAd(lAdPos), lAdSize, 0, 0, baClientIV(0), baClientKey(0)) <> 0 Then
            GoTo QH
        End If
    Else
        Err.Raise vbObjectError, "pvCryptoEncrypt", "Unsupported aead type " & eAead
    End If
    '--- success
    pvCryptoEncrypt = True
QH:
End Function

Private Function pvCryptoHash(eHash As UcsTlsCryptoAlgorithmsEnum, baInput() As Byte, ByVal lPos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim baRetVal()      As Byte
    Dim lPtr            As Long
    
    Select Case eHash
    Case ucsTlsAlgoDigestSha256
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
        Err.Raise vbObjectError, "pvCryptoHash", "Unsupported hash type " & eHash
    End Select
    pvCryptoHash = baRetVal
End Function

Private Function pvCryptoHmac(ByVal eHash As UcsTlsCryptoAlgorithmsEnum, baSalt() As Byte, baInput() As Byte, ByVal lPos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim baRetVal()      As Byte
    Dim lPtr            As Long
    
    Select Case eHash
    Case ucsTlsAlgoDigestSha256
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
        Err.Raise vbObjectError, "pvCryptoHmac", "Unsupported hash type " & eHash
    End Select
    pvCryptoHmac = baRetVal
End Function

Private Function pvCryptoDeriveSecret(ByVal eKeyX As UcsTlsCryptoAlgorithmsEnum, baPriv() As Byte, baPub() As Byte) As Byte()
    Dim baRetVal()      As Byte
    
    If eKeyX = ucsTlsAlgoKeyX25519 Then
        Debug.Assert pvArraySize(baPriv) = TLS_X25519_KEY_SIZE
        Debug.Assert pvArraySize(baPub) = TLS_X25519_KEY_SIZE
        ReDim baRetVal(0 To TLS_X25519_KEY_SIZE - 1) As Byte
        Call crypto_scalarmult_curve25519(baRetVal(0), baPriv(0), baPub(0))
    Else
        Err.Raise vbObjectError, "pvCryptoDeriveSecret", "Unsupported key-exchange type " & eKeyX
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

Private Function ToHex(baText() As Byte) As String
    ToHex = ToHexDump(StrConv(baText, vbUnicode))
End Function

Private Function FromHex(sText As String) As Byte()
    FromHex = StrConv(FromHexDump(sText), vbFromUnicode)
End Function

Private Function ToHexDump(sText As String) As String
    Dim lIdx            As Long
    
    For lIdx = 1 To Len(sText)
        ToHexDump = ToHexDump & Right$("0" & Hex$(Asc(Mid$(sText, lIdx, 1))), 2)
    Next
End Function

Private Function FromHexDump(sText As String) As String
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
