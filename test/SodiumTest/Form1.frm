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
   Begin VB.TextBox Text2 
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
      Height          =   348
      Left            =   3864
      TabIndex        =   2
      Top             =   168
      Width           =   1356
   End
   Begin VB.TextBox Text1 
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
'Private Const TLS_CIPHER_SUITE_AES_256_GCM_SHA384       As Long = &H1302
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
'Private Const TLS_CHACHA20POLY1305_TAG_SIZE             As Long = 16
'Private Const TLS_AES256_KEY_SIZE                       As Long = 32
'Private Const TLS_AESGCM_IV_SIZE                        As Long = 12
'Private Const TLS_AESGCM_TAG_SIZE                       As Long = 16
Private Const TLS_COMPRESS_NULL                         As Long = 0
Private Const TLS_SERVER_NAME_TYPE_HOSTNAME             As Long = 0
'--- libsodium
Private Const crypto_scalarmult_curve25519_BYTES        As Long = 32
Private Const crypto_hash_sha256_BYTES                  As Long = 32
Private Const crypto_auth_hmacsha256_BYTES              As Long = 32
Private Const crypto_auth_hmacsha256_KEYBYTES           As Long = 32

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (Source As Any, Destination As Any) As Long
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
'--- libsodium
Private Declare Function sodium_init Lib "libsodium" () As Long
'Private Declare Function sodium_runtime_has_aesni Lib "libsodium" () As Long
'Private Declare Function sodium_runtime_has_pclmul Lib "libsodium" () As Long
'Private Declare Function crypto_aead_aes256gcm_is_available Lib "libsodium" () As Long
Private Declare Function randombytes_buf Lib "libsodium" (lpOut As Any, ByVal lSize As Long) As Long
Private Declare Function crypto_scalarmult_curve25519 Lib "libsodium" (lpOut As Any, lpConstN As Any, lpConstP As Any) As Long
Private Declare Function crypto_scalarmult_curve25519_base Lib "libsodium" (lpOut As Any, lpN As Any) As Long
Private Declare Function crypto_hash_sha256 Lib "libsodium" (lpOut As Any, lpConstIn As Any, ByVal lSize As Long, ByVal lHighSize As Long) As Long
Private Declare Function crypto_auth_hmacsha256 Lib "libsodium" (lpOut As Any, lpConstIn As Any, ByVal lSize As Long, ByVal lHighSize As Long, lpConstKey As Any) As Long
Private Declare Function crypto_aead_chacha20poly1305_ietf_decrypt Lib "libsodium" (lpOut As Any, lOutSize As Any, ByVal nSec As Long, lConstIn As Any, ByVal lInSize As Long, ByVal lHighInSize As Long, lpConstAd As Any, ByVal lAdSize As Long, ByVal lHighAdSize As Long, lpConstNonce As Any, lpConstKey As Any) As Long
Private Declare Function crypto_aead_chacha20poly1305_ietf_encrypt Lib "libsodium" (lpOut As Any, lOutSize As Any, lConstIn As Any, ByVal lInSize As Long, ByVal lHighInSize As Long, lpConstAd As Any, ByVal lAdSize As Long, ByVal lHighAdSize As Long, ByVal nSec As Long, lpConstNonce As Any, lpConstKey As Any) As Long
'Private Declare Function crypto_aead_aes256gcm_decrypt Lib "libsodium" (lpOut As Any, lOutSize As Any, ByVal nSec As Long, lConstIn As Any, ByVal lInSize As Long, ByVal lHighInSize As Long, lpConstAd As Any, ByVal lAdSize As Long, ByVal lHighAdSize As Long, lpConstNonce As Any, lpConstKey As Any) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const LNG_KEY_SIZE              As Long = crypto_scalarmult_curve25519_BYTES
Private Const LNG_HASH_SIZE             As Long = crypto_hash_sha256_BYTES
Private Const LNG_SALT_SIZE             As Long = crypto_auth_hmacsha256_KEYBYTES
Private Const LNG_HELLO_RANDOM_SIZE     As Long = 32
Private Const LNG_IV_SIZE               As Long = 12
Private Const LNG_TAG_SIZE              As Long = 16

Private Enum UcsClientStateEnum
    ucsStateHandshakeStart
    ucsStateExpectServerHello
    ucsStateExpectEncryptedExtensions
    ucsStatePostHandshake
End Enum

Private Type UcsClientContextType
    State                       As UcsClientStateEnum
    ClientRandom()              As Byte
    ServerName                  As String
    LegacySessionID()           As Byte
    ClientPrivate()             As Byte
    ClientPublic()              As Byte
    ServerRandom()              As Byte '--- not used
    CipherSuite                 As Long
    ServerPublic()              As Byte
    ServerSupportedVersion      As Long '--- not used
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
    
    ServerApplicationSecret()   As Byte
    ServerApplicationKey()      As Byte
    ServerApplicationIV()       As Byte
    ServerApplicationSeqNo      As Long
    ClientApplicationSecret()   As Byte
    ClientApplicationKey()      As Byte
    ClientApplicationIV()       As Byte
    ClientApplicationSeqNo      As Long
    
    ServerCertificate()         As Byte

    SendBuffer()                As Byte
    SendPos                     As Long
    RecvBuffer()                As Byte
    RecvPos                     As Long
    
    DebugBox                    As TextBox
End Type

'=========================================================================
' Methods
'=========================================================================

Private Function pvInitClient(sServerName As String, oDebugBox As TextBox) As UcsClientContextType
    Dim uRetVal         As UcsClientContextType
    
    With uRetVal
        ReDim .ClientRandom(0 To LNG_HELLO_RANDOM_SIZE - 1) As Byte
        Call randombytes_buf(.ClientRandom(0), UBound(.ClientRandom) + 1)
'        ReDim .LegacySessionID(0 To 31) As Byte
'        Call randombytes_buf(.LegacySessionID(0), UBound(.LegacySessionID) + 1)
        ReDim .ClientPrivate(0 To LNG_KEY_SIZE - 1) As Byte
        Call randombytes_buf(.ClientPrivate(0), UBound(.ClientPrivate) + 1)
        .ClientPrivate(0) = .ClientPrivate(0) And 248
        .ClientPrivate(UBound(.ClientPrivate)) = (.ClientPrivate(UBound(.ClientPrivate)) And 127) Or 64
        ReDim .ClientPublic(0 To LNG_KEY_SIZE - 1) As Byte
        Call crypto_scalarmult_curve25519_base(.ClientPublic(0), .ClientPrivate(0))
        .ServerName = sServerName
        Set .DebugBox = oDebugBox
    End With
    pvInitClient = uRetVal
End Function

Private Function pvGetClientHello(uCtx As UcsClientContextType) As Byte()
    Dim baRetVal()      As Byte
    Dim lPos            As Long
    Dim cBlocks         As Collection
    
    '--- Record Header
    lPos = pvAppendLong(baRetVal, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
    lPos = pvAppendLong(baRetVal, lPos, TLS_RECORD_VERSION, Size:=2)
    lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks, Size:=2)
        '--- Handshake Header
        lPos = pvAppendLong(baRetVal, lPos, TLS_HANDSHAKE_TYPE_CLIENT_HELLO)
        lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks, Size:=3)
            lPos = pvAppendLong(baRetVal, lPos, TLS_CLIENT_LEGACY_VERSION, Size:=2)
            lPos = pvAppendArray(baRetVal, lPos, uCtx.ClientRandom)
            '--- Legacy Session ID
            lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks)
                lPos = pvAppendArray(baRetVal, lPos, uCtx.LegacySessionID)
            lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
            '--- Cipher Suites
            lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks, Size:=2)
'                lPos = pvAppendLong(baRetVal, lPos, TLS_CIPHER_SUITE_AES_256_GCM_SHA384, Size:=2)
                lPos = pvAppendLong(baRetVal, lPos, TLS_CIPHER_SUITE_CHACHA20_POLY1305_SHA256, Size:=2)
            lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
            '--- Legacy Compression Methods
            lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks)
                lPos = pvAppendLong(baRetVal, lPos, TLS_COMPRESS_NULL)
            lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
            '--- Extensions
            lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks, Size:=2)
                If LenB(uCtx.ServerName) <> 0 Then
                    '--- Extension - Server Name
                    lPos = pvAppendLong(baRetVal, lPos, TLS_EXTENSION_TYPE_SERVER_NAME, Size:=2)
                    lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks, Size:=2)
                        lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks, Size:=2)
                            lPos = pvAppendLong(baRetVal, lPos, TLS_SERVER_NAME_TYPE_HOSTNAME) '--- FQDN
                            lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks, Size:=2)
                                lPos = pvAppendString(baRetVal, lPos, uCtx.ServerName)
                            lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
                        lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
                    lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
                End If
                '--- Extension - Supported Groups
                lPos = pvAppendLong(baRetVal, lPos, TLS_EXTENSION_TYPE_SUPPORTED_GROUPS, Size:=2)
                lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks, Size:=2)
                    lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks, Size:=2)
                        lPos = pvAppendLong(baRetVal, lPos, TLS_GROUP_X25519, Size:=2)
                    lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
                lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
                '--- Extension - Signature Algorithms
                lPos = pvAppendLong(baRetVal, lPos, TLS_EXTENSION_TYPE_SIGNATURE_ALGORITHMS, Size:=2)
                lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks, Size:=2)
                    lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks, Size:=2)
                        lPos = pvAppendLong(baRetVal, lPos, TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, Size:=2)
                        lPos = pvAppendLong(baRetVal, lPos, TLS_SIGNATURE_ECDSA_SECP256R1_SHA256, Size:=2)
                        lPos = pvAppendLong(baRetVal, lPos, TLS_SIGNATURE_RSA_PKCS1_SHA256, Size:=2)
                        lPos = pvAppendLong(baRetVal, lPos, TLS_SIGNATURE_RSA_PKCS1_SHA1, Size:=2)
                    lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
                lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
                '--- Extension - Key Share
                lPos = pvAppendLong(baRetVal, lPos, TLS_EXTENSION_TYPE_KEY_SHARE, Size:=2)
                lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks, Size:=2)
                    lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks, Size:=2)
                        lPos = pvAppendLong(baRetVal, lPos, TLS_GROUP_X25519, Size:=2)
                        lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks, Size:=2)
                            lPos = pvAppendArray(baRetVal, lPos, uCtx.ClientPublic)
                        lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
                    lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
                lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
                '--- Extension - PSK Key Exchange Modes
                lPos = pvAppendLong(baRetVal, lPos, TLS_EXTENSION_TYPE_PSK_KEY_EXCHANGE_MODES, Size:=2)
                lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks, Size:=2)
                    lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks)
                        lPos = pvAppendLong(baRetVal, lPos, TLS_PSK_KE_MODE_PSK_DHE)
                    lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
                lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
                '--- Extension - Supported Versions
                lPos = pvAppendLong(baRetVal, lPos, TLS_EXTENSION_TYPE_SUPPORTED_VERSIONS, Size:=2)
                lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks, Size:=2)
                    lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks)
                        lPos = pvAppendLong(baRetVal, lPos, TLS_PROTOCOL_VERSION_TLS13_FINAL, Size:=2)
                    lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
                lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
            lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
        lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
    lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
    Debug.Assert cBlocks.Count = 0
    pvGetClientHello = baRetVal
End Function

Private Function pvGetClientHandshakeFinished(uCtx As UcsClientContextType, sError As String) As Byte()
    Dim baRetVal()      As Byte
    Dim lPos            As Long
    Dim cBlocks         As Collection
    Dim lRecordPos      As Long
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim baVerifyData()  As Byte
    Dim baClientIV()    As Byte
    Dim baClientKey()   As Byte
    Dim lResult         As Long
    
    '--- Legacy Change Cipher Spec
    baVerifyData = FromHex("140303000101")
    lPos = pvAppendArray(baRetVal, lPos, baVerifyData)
    '--- Record Header
    lRecordPos = lPos
    lPos = pvAppendLong(baRetVal, lPos, TLS_CONTENT_TYPE_APPDATA)
    lPos = pvAppendLong(baRetVal, lPos, TLS_RECORD_VERSION, Size:=2)
    lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks, Size:=2)
        lMessagePos = lPos
        '--- Handshake Finish
        lPos = pvAppendLong(baRetVal, lPos, TLS_HANDSHAKE_TYPE_FINISHED)
        lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks, Size:=3)
            ReDim baHandshakeHash(0 To LNG_HASH_SIZE - 1) As Byte
            Call crypto_hash_sha256(baHandshakeHash(0), uCtx.HandshakeMessages(0), pvArraySize(uCtx.HandshakeMessages), 0)
            baVerifyData = pvHkdfExpand(uCtx.ClientTrafficSecret, "finished", EmptyByteArray, LNG_HASH_SIZE)
            baVerifyData = pvHkdfExtract(baVerifyData, baHandshakeHash)
            lPos = pvAppendArray(baRetVal, lPos, baVerifyData)
        lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
        lPos = pvAppendLong(baRetVal, lPos, TLS_CONTENT_TYPE_HANDSHAKE)
        lMessageSize = lPos - lMessagePos
        lPos = pvAppendReserve(baRetVal, lPos, LNG_TAG_SIZE)
    lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
    Debug.Assert cBlocks.Count = 0
    baClientIV = pvArrayXor(uCtx.ClientTrafficIV, uCtx.ClientTrafficSeqNo)
    uCtx.ClientTrafficSeqNo = uCtx.ClientTrafficSeqNo + 1
    baClientKey = uCtx.ClientTrafficKey
    If uCtx.CipherSuite = TLS_CIPHER_SUITE_CHACHA20_POLY1305_SHA256 Then
        Debug.Assert UBound(baClientIV) + 1 = TLS_CHACHA20POLY1305_IV_SIZE
        Debug.Assert UBound(baClientKey) + 1 = TLS_CHACHA20_KEY_SIZE
        lResult = crypto_aead_chacha20poly1305_ietf_encrypt(baRetVal(lMessagePos), ByVal 0, baRetVal(lMessagePos), lMessageSize, 0, baRetVal(lRecordPos), 5, 0, 0, baClientIV(0), baClientKey(0))
        Debug.Assert lResult = 0
        If lResult <> 0 Then
            sError = "crypto_aead_chacha20poly1305_ietf_encrypt failed: " & lResult
            GoTo QH
        End If
    Else
        sError = "Invalid cipher suite (0x" & Hex$(uCtx.CipherSuite) & ")"
        GoTo QH
    End If
    pvGetClientHandshakeFinished = baRetVal
QH:
End Function

Private Function pvGetClientApplicationData(uCtx As UcsClientContextType, baData() As Byte, sError As String) As Byte()
    Dim baRetVal()      As Byte
    Dim lPos            As Long
    Dim cBlocks         As Collection
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim baClientIV()    As Byte
    Dim baClientKey()   As Byte
    Dim lResult         As Long
    
    '--- Record Header
    lPos = pvAppendLong(baRetVal, lPos, TLS_CONTENT_TYPE_APPDATA)
    lPos = pvAppendLong(baRetVal, lPos, TLS_RECORD_VERSION, Size:=2)
    lPos = pvAppendBeginBlock(baRetVal, lPos, cBlocks, Size:=2)
        lMessagePos = lPos
        lPos = pvAppendArray(baRetVal, lPos, baData)
        lPos = pvAppendLong(baRetVal, lPos, TLS_CONTENT_TYPE_APPDATA)
        lMessageSize = lPos - lMessagePos
        lPos = pvAppendReserve(baRetVal, lPos, LNG_TAG_SIZE)
    lPos = pvAppendEndBlock(baRetVal, lPos, cBlocks)
    Debug.Assert cBlocks.Count = 0
    baClientIV = pvArrayXor(uCtx.ClientApplicationIV, uCtx.ClientApplicationSeqNo)
    uCtx.ClientApplicationSeqNo = uCtx.ClientApplicationSeqNo + 1
    baClientKey = uCtx.ClientApplicationKey
    If uCtx.CipherSuite = TLS_CIPHER_SUITE_CHACHA20_POLY1305_SHA256 Then
        Debug.Assert UBound(baClientIV) + 1 = TLS_CHACHA20POLY1305_IV_SIZE
        Debug.Assert UBound(baClientKey) + 1 = TLS_CHACHA20_KEY_SIZE
        lResult = crypto_aead_chacha20poly1305_ietf_encrypt(baRetVal(lMessagePos), ByVal 0, baRetVal(lMessagePos), lMessageSize, 0, baRetVal(0), 5, 0, 0, baClientIV(0), baClientKey(0))
        Debug.Assert lResult = 0
        If lResult <> 0 Then
            sError = "crypto_aead_chacha20poly1305_ietf_encrypt failed: " & lResult
            GoTo QH
        End If
    Else
        sError = "Invalid cipher suite (0x" & Hex$(uCtx.CipherSuite) & ")"
        GoTo QH
    End If
    pvGetClientApplicationData = baRetVal
QH:
End Function

Private Function pvHandleInput(uCtx As UcsClientContextType, baInput() As Byte, sError As String) As Boolean
    Dim lRecordPos      As Long
    Dim lRecordSize     As Long
    Dim lPos            As Long
    Dim cBlocks         As Collection
    Dim lRecordType     As Long
    Dim lLegacyProtocol As Long
    Dim lHandshakeType  As Long
    Dim baServerIV()    As Byte
    Dim baServerKey()   As Byte
    Dim lResult         As Long
    Dim lEnd            As Long
    Dim baVerifyData()  As Byte
    Dim baMessage()     As Byte
    Dim lMessagePos     As Long
    Dim lMessageSize    As Long
    Dim lVerifyPos      As Long
    Dim baHandshakeHash() As Byte
    Dim lRequestUpdate  As Long
    
    Do While lPos < pvArraySize(baInput)
        lRecordPos = lPos
        lPos = pvDecodeLong(baInput, lPos, lRecordType)
        lPos = pvDecodeLong(baInput, lPos, lLegacyProtocol, Size:=2)
        lPos = pvDecodeBeginBlock(baInput, lPos, cBlocks, Size:=2, BlockSize:=lRecordSize)
            Select Case lRecordType
            Case TLS_CONTENT_TYPE_CHANGE_CIPHER_SPEC
                lPos = lPos + lRecordSize
            Case TLS_CONTENT_TYPE_HANDSHAKE
                lMessagePos = lPos
                lPos = pvDecodeLong(baInput, lPos, lHandshakeType)
                lPos = pvDecodeBeginBlock(baInput, lPos, cBlocks, Size:=3, BlockSize:=lMessageSize)
                    Select Case uCtx.State
                    Case ucsStateExpectServerHello
                        Select Case lHandshakeType
                        Case TLS_HANDSHAKE_TYPE_SERVER_HELLO
                            lPos = pvDecodeArray(baInput, lPos, baMessage, lMessageSize)
                            If Not pvDecodeServerHello(uCtx, baMessage, sError) Then
                                GoTo QH
                            End If
                            pvAppendBuffer uCtx.HandshakeMessages, pvArraySize(uCtx.HandshakeMessages), VarPtr(baInput(lMessagePos)), lMessageSize + 4
                            If Not pvDeriveHandshakeSecrets(uCtx, sError) Then
                                GoTo QH
                            End If
                            uCtx.State = ucsStateExpectEncryptedExtensions
                        Case Else
                            sError = "Unexpected message type for ucsStateExpectServerHello (lHandshakeType=" & lHandshakeType & ")"
                            GoTo QH
                        End Select
                    End Select
                lPos = pvDecodeEndBlock(baInput, lPos, cBlocks)
            Case TLS_CONTENT_TYPE_APPDATA
                If uCtx.State < ucsStatePostHandshake Then
                    baServerIV = pvArrayXor(uCtx.ServerTrafficIV, uCtx.ServerTrafficSeqNo)
                    uCtx.ServerTrafficSeqNo = uCtx.ServerTrafficSeqNo + 1
                    baServerKey = uCtx.ServerTrafficKey
                Else
                    baServerIV = pvArrayXor(uCtx.ServerApplicationIV, uCtx.ServerApplicationSeqNo)
                    uCtx.ServerApplicationSeqNo = uCtx.ServerApplicationSeqNo + 1
                    baServerKey = uCtx.ServerApplicationKey
                End If
                If uCtx.CipherSuite = TLS_CIPHER_SUITE_CHACHA20_POLY1305_SHA256 Then
                    Debug.Assert UBound(baServerIV) + 1 = TLS_CHACHA20POLY1305_IV_SIZE
                    Debug.Assert UBound(baServerKey) + 1 = TLS_CHACHA20_KEY_SIZE
                    lResult = crypto_aead_chacha20poly1305_ietf_decrypt(baInput(lPos), ByVal 0, 0, baInput(lPos), lRecordSize, 0, baInput(lRecordPos), 5, 0, baServerIV(0), baServerKey(0))
                    Debug.Assert lResult = 0
                    If lResult <> 0 Then
                        sError = "crypto_aead_chacha20poly1305_ietf_decrypt failed: " & lResult
                        GoTo QH
                    End If
                    If Not uCtx.DebugBox Is Nothing Then
                        uCtx.DebugBox.Text = uCtx.DebugBox.Text & vbCrLf & "Decrypted: " & lRecordSize - LNG_TAG_SIZE & vbCrLf & DesignDumpMemory(VarPtr(baInput(lPos)), lRecordSize - LNG_TAG_SIZE)
                    End If
                Else
                    sError = "Invalid cipher suite (0x" & Hex$(uCtx.CipherSuite) & ")"
                    GoTo QH
                End If
                lEnd = lPos + lRecordSize - LNG_TAG_SIZE - 1
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
                            lPos = pvDecodeLong(baInput, lPos, lHandshakeType)
                            lPos = pvDecodeBeginBlock(baInput, lPos, cBlocks, Size:=3, BlockSize:=lMessageSize)
                            If lMessageSize > 16 * 1024 Then
                                sError = "Unexpected message size (lMessageSize=" & lMessageSize & ")"
                                GoTo QH
                            End If
                            lPos = pvDecodeArray(baInput, lPos, baMessage, lMessageSize)
                            Select Case lHandshakeType
                            Case TLS_HANDSHAKE_TYPE_CERTIFICATE
                                uCtx.ServerCertificate = baMessage
                            Case TLS_HANDSHAKE_TYPE_CERTIFICATE_VERIFY
                                ReDim baHandshakeHash(0 To LNG_HASH_SIZE - 1) As Byte
                                Call crypto_hash_sha256(baHandshakeHash(0), uCtx.HandshakeMessages(0), pvArraySize(uCtx.HandshakeMessages), 0)
                                lVerifyPos = pvAppendString(baVerifyData, 0, Space$(64) & "TLS 1.3, server CertificateVerify" & Chr$(0))
                                lVerifyPos = pvAppendArray(baVerifyData, lVerifyPos, baHandshakeHash)
                                '--- ToDo: verify uCtx.ServerCertificate signature
                                '--- ShellExecute("openssl x509 -pubkey -noout -in server.crt > server.pub")
                            Case TLS_HANDSHAKE_TYPE_FINISHED
                                ReDim baHandshakeHash(0 To LNG_HASH_SIZE - 1) As Byte
                                Call crypto_hash_sha256(baHandshakeHash(0), uCtx.HandshakeMessages(0), pvArraySize(uCtx.HandshakeMessages), 0)
                                baVerifyData = pvHkdfExpand(uCtx.ServerTrafficSecret, "finished", EmptyByteArray, LNG_HASH_SIZE)
                                baVerifyData = pvHkdfExtract(baVerifyData, baHandshakeHash)
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
                            pvAppendBuffer uCtx.HandshakeMessages, pvArraySize(uCtx.HandshakeMessages), VarPtr(baInput(lMessagePos)), lMessageSize + 4
                            lPos = pvDecodeEndBlock(baInput, lPos, cBlocks)
                        Loop
                        '--- note: skip padding too
                        lPos = lRecordPos + lRecordSize + 5
                    Case ucsStatePostHandshake
                        Do While lPos < lEnd
                            lMessagePos = lPos
                            lPos = pvDecodeLong(baInput, lPos, lHandshakeType)
                            lPos = pvDecodeBeginBlock(baInput, lPos, cBlocks, Size:=3, BlockSize:=lMessageSize)
                            lPos = pvDecodeArray(baInput, lPos, baMessage, lMessageSize)
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
                                    If Not pvDeriveServerKeyUpdate(uCtx, sError) Then
                                        GoTo QH
                                    End If
                                    If lRequestUpdate = 1 Then
                                        If Not pvDeriveClientKeyUpdate(uCtx, sError) Then
                                            GoTo QH
                                        End If
                                        '--- ack by TLS_HANDSHAKE_TYPE_KEY_UPDATE w/ update_not_requested(0)
                                        baMessage = FromHex("1800000100")
                                        baMessage = pvGetClientApplicationData(uCtx, baMessage, sError)
                                        If pvArraySize(baMessage) = 0 Then
                                            GoTo QH
                                        End If
                                        uCtx.SendPos = pvAppendArray(uCtx.SendBuffer, uCtx.SendPos, baMessage)
                                    End If
                                Case Else
                                    sError = "Unexpected value in TLS_HANDSHAKE_TYPE_KEY_UPDATE (lRequestUpdate=" & lRequestUpdate & ")"
                                    GoTo QH
                                End Select
                            Case Else
                                sError = "Unexpected message type for ucsStatePostHandshake (lHandshakeType=" & lHandshakeType & ")"
                                GoTo QH
                            End Select
                            lPos = pvDecodeEndBlock(baInput, lPos, cBlocks)
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
                            Switch(baInput(lPos) = 1, "Warning alert: ", baInput(lPos) = 2, "Fatal alert: ", True, "Unknown alert: ") & baInput(lPos + 1)
                    End If
                    lPos = lPos + lRecordSize
                Case TLS_CONTENT_TYPE_APPDATA
                    Select Case uCtx.State
                    Case ucsStatePostHandshake
                        uCtx.RecvPos = pvAppendBuffer(uCtx.RecvBuffer, uCtx.RecvPos, VarPtr(baInput(lPos)), lEnd - lPos)
                    Case Else
                        sError = "Invalid state for TLS_CONTENT_TYPE_APPDATA (" & uCtx.State & ")"
                        GoTo QH
                    End Select
                    lPos = lPos + lRecordSize
                Case Else
                    lPos = lPos + lRecordSize
                End Select
            Case Else
                sError = "Unexpected record type (" & lRecordType & ")"
                GoTo QH
            End Select
        lPos = pvDecodeEndBlock(baInput, lPos, cBlocks)
    Loop
    '--- success
    pvHandleInput = True
QH:
End Function

Private Function pvDecodeServerHello(uCtx As UcsClientContextType, baInput() As Byte, sError As String) As Boolean
    Dim lPos            As Long
    Dim cBlocks         As Collection
    Dim lLegacyVersion  As Long
    Dim lBlockSize      As Long
    Dim lLegacyCompression As Long
    Dim lExtType        As Long
    Dim lExchangeGroup  As Long
    
    lPos = pvDecodeLong(baInput, lPos, lLegacyVersion, Size:=2)
    lPos = pvDecodeArray(baInput, lPos, uCtx.ServerRandom, LNG_HELLO_RANDOM_SIZE)
    lPos = pvDecodeBeginBlock(baInput, lPos, cBlocks, BlockSize:=lBlockSize)
        lPos = pvDecodeArray(baInput, lPos, uCtx.LegacySessionID, lBlockSize)
    lPos = pvDecodeEndBlock(baInput, lPos, cBlocks)
    lPos = pvDecodeLong(baInput, lPos, uCtx.CipherSuite, Size:=2)
    lPos = pvDecodeLong(baInput, lPos, lLegacyCompression)
    Debug.Assert lLegacyCompression = 0
    lPos = pvDecodeBeginBlock(baInput, lPos, cBlocks, Size:=2)
        Do While lPos < cBlocks.Item(cBlocks.Count)
            lPos = pvDecodeLong(baInput, lPos, lExtType, Size:=2)
            lPos = pvDecodeBeginBlock(baInput, lPos, cBlocks, Size:=2, BlockSize:=lBlockSize)
                Select Case lExtType
                Case TLS_EXTENSION_TYPE_KEY_SHARE
                    lPos = pvDecodeLong(baInput, lPos, lExchangeGroup, Size:=2)
                    Debug.Assert lExchangeGroup = TLS_GROUP_X25519
                    lPos = pvDecodeBeginBlock(baInput, lPos, cBlocks, Size:=2, BlockSize:=lBlockSize)
                        Debug.Assert lBlockSize = LNG_KEY_SIZE
                        If lBlockSize <> LNG_KEY_SIZE Then
                            sError = "Invalid server key size"
                            GoTo QH
                        End If
                        lPos = pvDecodeArray(baInput, lPos, uCtx.ServerPublic, LNG_KEY_SIZE)
                    lPos = pvDecodeEndBlock(baInput, lPos, cBlocks)
                Case TLS_EXTENSION_TYPE_SUPPORTED_VERSIONS
                    If lBlockSize >= 2 Then
                        Call pvDecodeLong(baInput, lPos, uCtx.ServerSupportedVersion, Size:=2)
                    End If
                    lPos = lPos + lBlockSize
                Case Else
                    lPos = lPos + lBlockSize
                End Select
            lPos = pvDecodeEndBlock(baInput, lPos, cBlocks)
        Loop
    lPos = pvDecodeEndBlock(baInput, lPos, cBlocks)
    '--- success
    pvDecodeServerHello = True
QH:
End Function

Private Function pvDeriveHandshakeSecrets(uCtx As UcsClientContextType, sError As String) As Boolean
    Dim baHelloHash()   As Byte
    Dim baZeroKey()     As Byte
    Dim baZeroSalt()    As Byte
    Dim baEarlySecret() As Byte
    Dim baEmptyHash()   As Byte
    Dim baDerivedSecret() As Byte
    Dim baSharedSecret() As Byte
    
    If pvArraySize(uCtx.HandshakeMessages) = 0 Then
        sError = "Missing handshake records"
        GoTo QH
    End If
    ReDim baHelloHash(0 To LNG_HASH_SIZE - 1) As Byte
    Call crypto_hash_sha256(baHelloHash(0), uCtx.HandshakeMessages(0), pvArraySize(uCtx.HandshakeMessages), 0)
    ReDim baZeroKey(0 To LNG_KEY_SIZE - 1) As Byte
    ReDim baZeroSalt(0 To LNG_SALT_SIZE - 1) As Byte
    baEarlySecret = pvHkdfExtract(baZeroSalt, baZeroKey)                                    ' 33AD0A1C607EC03B09E6CD9893680CE210ADF300AA1F2660E1B22E10F170F92A
    ReDim baEmptyHash(0 To LNG_HASH_SIZE - 1) As Byte
    Call crypto_hash_sha256(baEmptyHash(0), ByVal 0, 0, 0)                                  ' E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855
    baDerivedSecret = pvHkdfExpand(baEarlySecret, "derived", baEmptyHash, LNG_HASH_SIZE)    ' 6F2615A108C702C5678F54FC9DBAB69716C076189C48250CEBEAC3576C3611BA
    ReDim baSharedSecret(0 To LNG_KEY_SIZE - 1) As Byte
    Call crypto_scalarmult_curve25519(baSharedSecret(0), uCtx.ClientPrivate(0), uCtx.ServerPublic(0))
    uCtx.HandshakeSecret = pvHkdfExtract(baDerivedSecret, baSharedSecret)
    
    uCtx.ServerTrafficSecret = pvHkdfExpand(uCtx.HandshakeSecret, "s hs traffic", baHelloHash, LNG_HASH_SIZE)
    uCtx.ServerTrafficKey = pvHkdfExpand(uCtx.ServerTrafficSecret, "key", EmptyByteArray, LNG_KEY_SIZE)
    uCtx.ServerTrafficIV = pvHkdfExpand(uCtx.ServerTrafficSecret, "iv", EmptyByteArray, LNG_IV_SIZE)
    uCtx.ServerTrafficSeqNo = 0
    uCtx.ClientTrafficSecret = pvHkdfExpand(uCtx.HandshakeSecret, "c hs traffic", baHelloHash, LNG_HASH_SIZE)
    uCtx.ClientTrafficKey = pvHkdfExpand(uCtx.ClientTrafficSecret, "key", EmptyByteArray, LNG_KEY_SIZE)
    uCtx.ClientTrafficIV = pvHkdfExpand(uCtx.ClientTrafficSecret, "iv", EmptyByteArray, LNG_IV_SIZE)
    uCtx.ClientTrafficSeqNo = 0
    '--- success
    pvDeriveHandshakeSecrets = True
QH:
End Function

Private Function pvDeriveApplicationSecrets(uCtx As UcsClientContextType, sError As String) As Boolean
    Dim baHandshakeHash() As Byte
    Dim baZeroKey()     As Byte
    Dim baEmptyHash()   As Byte
    Dim baDerivedSecret() As Byte
    
    If pvArraySize(uCtx.HandshakeMessages) = 0 Then
        sError = "Missing handshake records"
        GoTo QH
    End If
    ReDim baHandshakeHash(0 To LNG_HASH_SIZE - 1) As Byte
    Call crypto_hash_sha256(baHandshakeHash(0), uCtx.HandshakeMessages(0), pvArraySize(uCtx.HandshakeMessages), 0)
    ReDim baZeroKey(0 To LNG_KEY_SIZE - 1) As Byte
    ReDim baEmptyHash(0 To LNG_HASH_SIZE - 1) As Byte
    Call crypto_hash_sha256(baEmptyHash(0), ByVal 0, 0, 0)                                      ' E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855
    baDerivedSecret = pvHkdfExpand(uCtx.HandshakeSecret, "derived", baEmptyHash, LNG_HASH_SIZE) ' 6F2615A108C702C5678F54FC9DBAB69716C076189C48250CEBEAC3576C3611BA
    uCtx.MasterSecret = pvHkdfExtract(baDerivedSecret, baZeroKey)
    
    uCtx.ServerApplicationSecret = pvHkdfExpand(uCtx.MasterSecret, "s ap traffic", baHandshakeHash, LNG_HASH_SIZE)
    uCtx.ServerApplicationKey = pvHkdfExpand(uCtx.ServerApplicationSecret, "key", EmptyByteArray, LNG_KEY_SIZE)
    uCtx.ServerApplicationIV = pvHkdfExpand(uCtx.ServerApplicationSecret, "iv", EmptyByteArray, LNG_IV_SIZE)
    uCtx.ServerApplicationSeqNo = 0
    uCtx.ClientApplicationSecret = pvHkdfExpand(uCtx.MasterSecret, "c ap traffic", baHandshakeHash, LNG_HASH_SIZE)
    uCtx.ClientApplicationKey = pvHkdfExpand(uCtx.ClientApplicationSecret, "key", EmptyByteArray, LNG_KEY_SIZE)
    uCtx.ClientApplicationIV = pvHkdfExpand(uCtx.ClientApplicationSecret, "iv", EmptyByteArray, LNG_IV_SIZE)
    uCtx.ClientApplicationSeqNo = 0
    pvCopySecrets uCtx
    '--- success
    pvDeriveApplicationSecrets = True
QH:
End Function

Private Sub pvCopySecrets(uCtx As UcsClientContextType)
    uCtx.ServerTrafficSecret = uCtx.ServerApplicationSecret
    uCtx.ServerTrafficKey = uCtx.ServerApplicationKey
    uCtx.ServerTrafficIV = uCtx.ServerApplicationIV
    uCtx.ServerTrafficSeqNo = uCtx.ServerApplicationSeqNo
    uCtx.ClientTrafficSecret = uCtx.ClientApplicationSecret
    uCtx.ClientTrafficKey = uCtx.ClientApplicationKey
    uCtx.ClientTrafficIV = uCtx.ClientApplicationIV
    uCtx.ClientTrafficSeqNo = uCtx.ClientApplicationSeqNo
End Sub

Private Function pvDeriveServerKeyUpdate(uCtx As UcsClientContextType, sError As String) As Boolean
    Dim baTemp()        As Byte
    
    baTemp = uCtx.ServerApplicationSecret
    If pvArraySize(baTemp) = 0 Then
        sError = "Missing previous server secret"
        GoTo QH
    End If
    uCtx.ServerApplicationSecret = pvHkdfExpand(baTemp, "traffic upd", EmptyByteArray, LNG_HASH_SIZE)
    uCtx.ServerApplicationKey = pvHkdfExpand(uCtx.ServerApplicationSecret, "key", EmptyByteArray, LNG_KEY_SIZE)
    uCtx.ServerApplicationIV = pvHkdfExpand(uCtx.ServerApplicationSecret, "iv", EmptyByteArray, LNG_IV_SIZE)
    uCtx.ServerApplicationSeqNo = 0
    pvCopySecrets uCtx
    '--- success
    pvDeriveServerKeyUpdate = True
QH:
End Function

Private Function pvDeriveClientKeyUpdate(uCtx As UcsClientContextType, sError As String) As Boolean
    Dim baTemp()        As Byte
    
    baTemp = uCtx.ClientApplicationSecret
    If pvArraySize(baTemp) = 0 Then
        sError = "Missing previous client secret"
        GoTo QH
    End If
    uCtx.ClientApplicationSecret = pvHkdfExpand(baTemp, "traffic upd", EmptyByteArray, LNG_HASH_SIZE)
    uCtx.ClientApplicationKey = pvHkdfExpand(uCtx.ClientApplicationSecret, "key", EmptyByteArray, LNG_KEY_SIZE)
    uCtx.ClientApplicationIV = pvHkdfExpand(uCtx.ClientApplicationSecret, "iv", EmptyByteArray, LNG_IV_SIZE)
    uCtx.ClientApplicationSeqNo = 0
    pvCopySecrets uCtx
    '--- success
    pvDeriveClientKeyUpdate = True
QH:
End Function

Private Function pvHkdfExtract(baSalt() As Byte, baInput() As Byte) As Byte()
    Dim baRetVal(0 To crypto_auth_hmacsha256_BYTES - 1) As Byte
    
    Debug.Assert pvArraySize(baSalt) = crypto_auth_hmacsha256_KEYBYTES
    Call crypto_auth_hmacsha256(baRetVal(0), baInput(0), UBound(baInput) + 1, 0, baSalt(0))
    pvHkdfExtract = baRetVal
End Function

Private Function pvHkdfExpand(baSalt() As Byte, ByVal sLabel As String, baContext() As Byte, ByVal lSize As Long) As Byte()
    Dim baRetVal()      As Byte
    Dim lRetValPos      As Long
    Dim baInfo()        As Byte
    Dim lInfoPos        As Long
    Dim baInput()       As Byte
    Dim lInputPos       As Long
    Dim lIdx            As Long
    Dim baLast()        As Byte
    
    Debug.Assert pvArraySize(baSalt) = crypto_auth_hmacsha256_KEYBYTES
    If LenB(sLabel) <> 0 Then
        sLabel = "tls13 " & sLabel
        ReDim baInfo(0 To 3 + Len(sLabel) + 1 + pvArraySize(baContext) - 1) As Byte
        lInfoPos = pvAppendLong(baInfo, lInfoPos, lSize, Size:=2)
        lInfoPos = pvAppendLong(baInfo, lInfoPos, Len(sLabel))
        lInfoPos = pvAppendString(baInfo, lInfoPos, sLabel)
        lInfoPos = pvAppendLong(baInfo, lInfoPos, pvArraySize(baContext))
        lInfoPos = pvAppendArray(baInfo, lInfoPos, baContext)
    Else
        baInfo = baContext
    End If
    lIdx = 1
    Do While lRetValPos < lSize
        lInputPos = pvAppendArray(baInput, 0, baLast)
        lInputPos = pvAppendArray(baInput, lInputPos, baInfo)
        lInputPos = pvAppendLong(baInput, lInputPos, lIdx)
        ReDim baLast(0 To crypto_auth_hmacsha256_BYTES - 1) As Byte
        Call crypto_auth_hmacsha256(baLast(0), baInput(0), lInputPos, 0, baSalt(0))
        lRetValPos = pvAppendArray(baRetVal, lRetValPos, baLast)
        lIdx = lIdx + 1
    Loop
    If UBound(baRetVal) <> lSize - 1 Then
        ReDim Preserve baRetVal(0 To lSize - 1) As Byte
    End If
    pvHkdfExpand = baRetVal
    Debug.Print "sLabel=" & sLabel & ", pvHkdfExpand=0x" & ToHex(baRetVal)
End Function

Private Function pvAppendBeginBlock(baDest() As Byte, ByVal lPos As Long, cBlocks As Collection, Optional ByVal Size As Long = 1) As Long
    If cBlocks Is Nothing Then
        Set cBlocks = New Collection
    End If
    cBlocks.Add lPos
    pvAppendBeginBlock = pvAppendReserve(baDest, lPos, Size)
    '--- note: keep Size in baDest
    baDest(lPos) = (Size And &HFF)
End Function

Private Function pvAppendEndBlock(baDest() As Byte, ByVal lPos As Long, cBlocks As Collection) As Long
    Dim lStart          As Long
    
    lStart = cBlocks.Item(cBlocks.Count)
    cBlocks.Remove cBlocks.Count
    pvAppendLong baDest, lStart, lPos - lStart - baDest(lStart), Size:=baDest(lStart)
    pvAppendEndBlock = lPos
End Function

Private Function pvAppendString(baDest() As Byte, ByVal lPos As Long, sValue As String) As Long
    pvAppendString = pvAppendArray(baDest, lPos, StrConv(sValue, vbFromUnicode))
End Function

Private Function pvAppendArray(baDest() As Byte, ByVal lPos As Long, baSrc() As Byte) As Long
    Dim lSize       As Long
    
    If pvArraySize(baSrc, RetVal:=lSize) > 0 Then
        lPos = pvAppendBuffer(baDest, lPos, VarPtr(baSrc(0)), lSize)
    End If
    pvAppendArray = lPos
End Function

Private Function pvAppendLong(baDest() As Byte, ByVal lPos As Long, ByVal lValue As Long, Optional ByVal Size As Long = 1) As Long
    Static baTemp(0 To 3) As Byte

    If Size <= 1 Then
        pvAppendLong = pvAppendBuffer(baDest, lPos, VarPtr(lValue), Size)
    Else
        baTemp(Size - 1) = (lValue And &HFF): lValue = lValue \ &H100
        baTemp(Size - 2) = (lValue And &HFF): lValue = lValue \ &H100
        If Size >= 3 Then
            baTemp(Size - 3) = (lValue And &HFF): lValue = lValue \ &H100
            If Size >= 4 Then
                baTemp(Size - 4) = (lValue And &HFF)
            End If
        End If
        pvAppendLong = pvAppendBuffer(baDest, lPos, VarPtr(baTemp(0)), Size)
    End If
End Function

Private Function pvAppendReserve(baDest() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Long
    pvAppendReserve = pvAppendBuffer(baDest, lPos, 0, lSize)
End Function

Private Function pvAppendBuffer(baDest() As Byte, ByVal lPos As Long, ByVal lPtr As Long, ByVal lSize As Long) As Long
    If Peek(ArrPtr(baDest)) = 0 Then
        ReDim baDest(0 To lPos + lSize - 1) As Byte
    ElseIf UBound(baDest) < lPos + lSize - 1 Then
        ReDim Preserve baDest(0 To lPos + lSize - 1) As Byte
    End If
    If lSize > 0 And lPtr <> 0 Then
        Debug.Assert IsBadReadPtr(lPtr, lSize) = 0
        Call CopyMemory(baDest(lPos), ByVal lPtr, lSize)
    End If
    pvAppendBuffer = lPos + lSize
End Function

Private Function pvDecodeBeginBlock(baInput() As Byte, ByVal lPos As Long, cBlocks As Collection, Optional ByVal Size As Long = 1, Optional BlockSize As Long) As Long
    If cBlocks Is Nothing Then
        Set cBlocks = New Collection
    End If
    pvDecodeBeginBlock = pvDecodeLong(baInput, lPos, BlockSize, Size)
    cBlocks.Add pvDecodeBeginBlock + BlockSize
End Function

Private Function pvDecodeEndBlock(baInput() As Byte, ByVal lPos As Long, cBlocks As Collection) As Long
    Dim lEnd          As Long
    
    #If baInput Then '--- touch args
    #End If
    lEnd = cBlocks.Item(cBlocks.Count)
    cBlocks.Remove cBlocks.Count
    Debug.Assert lPos = lEnd
    pvDecodeEndBlock = lEnd
End Function

Private Function pvDecodeLong(baInput() As Byte, ByVal lPos As Long, lValue As Long, Optional ByVal Size As Long = 1) As Long
    Static baTemp(0 To 3) As Byte
    
    If lPos + Size <= pvArraySize(baInput) Then
        If Size <= 1 Then
            lValue = baInput(lPos)
        Else
            baTemp(Size - 1) = baInput(lPos + 0)
            baTemp(Size - 2) = baInput(lPos + 1)
            If Size >= 3 Then baTemp(Size - 3) = baInput(lPos + 2)
            If Size >= 4 Then baTemp(Size - 4) = baInput(lPos + 3)
            Call CopyMemory(lValue, baTemp(0), Size)
        End If
    Else
        lValue = 0
    End If
    pvDecodeLong = lPos + Size
End Function

Private Function pvDecodeArray(baInput() As Byte, ByVal lPos As Long, baDest() As Byte, ByVal lSize As Long) As Long
    If lSize < 0 Then
        lSize = pvArraySize(baInput) - lPos
    End If
    If lSize > 0 Then
        ReDim baDest(0 To lSize - 1) As Byte
        If lPos + lSize <= pvArraySize(baInput) Then
            Call CopyMemory(baDest(0), baInput(lPos), lSize)
        ElseIf lPos < pvArraySize(baInput) Then
            Call CopyMemory(baDest(0), baInput(lPos), pvArraySize(baInput) - lPos)
        End If
    Else
        Erase baDest
    End If
    pvDecodeArray = lPos + lSize
End Function

'Private Function pvDecodeRecord(baInput() As Byte, ByVal lPos As Long) As Byte()
'    Dim baRetVal()      As Byte
'    Dim lSize           As Long
'
'    lSize = baInput(lPos + 3) * &H100& + baInput(lPos + 4)
'    pvDecodeArray baInput, lPos + 5, baRetVal, lSize
'    pvDecodeRecord = baRetVal
'End Function

Private Function pvArraySize(baSrc() As Byte, Optional RetVal As Long) As Long
    If Peek(ArrPtr(baSrc)) <> 0 Then
        RetVal = UBound(baSrc) + 1
    Else
        RetVal = 0
    End If
    pvArraySize = RetVal
End Function

Private Function pvArrayXor(baInput() As Byte, ByVal lSeqNo As Long) As Byte()
    Dim baRetVal()      As Byte
    Dim lIdx            As Long
    
    baRetVal = baInput
    lIdx = pvArraySize(baRetVal)
    Do While lSeqNo <> 0 And lIdx > 0
        lIdx = lIdx - 1
        baRetVal(lIdx) = baRetVal(lIdx) Xor (lSeqNo And &HFF)
        lSeqNo = lSeqNo \ &H100
    Loop
    pvArrayXor = baRetVal
End Function

Private Function Peek(ByVal lPtr As Long) As Long
    Call GetMem4(ByVal lPtr, Peek)
End Function

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

Property Get EmptyByteArray() As Byte()

End Property

'=========================================================================
' Events
'=========================================================================

Private Sub Form_Load()
    If GetModuleHandle("libsodium.dll") = 0 Then
        Call LoadLibrary(App.Path & "\libsodium.dll")
        Call sodium_init
    End If
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        Text2.Move 0, Text2.Top, ScaleWidth, ScaleHeight - Text2.Top
    End If
End Sub

Private Sub Command1_Click()
    Dim oSocket         As cAsyncSocket
    Dim vSplit          As Variant
    Dim uCtx            As UcsClientContextType
    Dim baSend()        As Byte
    Dim baRecv()        As Byte
    Dim sError          As String
    Dim sRequest        As String
    
    Set oSocket = New cAsyncSocket
    ' tls13.1d.pw, localhost:44330
    vSplit = Split(Text1.Text & ":443", ":")
    If Not oSocket.SyncConnect(CStr(vSplit(0)), Val(vSplit(1))) Then
        sError = oSocket.GetErrorDescription(oSocket.LastError)
        GoTo QH
    End If
    uCtx = pvInitClient(Text1.Text, Text2)
    baSend = pvGetClientHello(uCtx)
    If pvArraySize(baSend) = 0 Then
        GoTo QH
    End If
    pvAppendBuffer uCtx.HandshakeMessages, 0, VarPtr(baSend(5)), pvArraySize(baSend) - 5
    If Not oSocket.SyncSendArray(baSend) Then
        sError = oSocket.GetErrorDescription(oSocket.LastError)
        GoTo QH
    End If
    uCtx.State = ucsStateExpectServerHello
    If Not oSocket.SyncReceiveArray(baRecv, Timeout:=-1) Then
        sError = oSocket.GetErrorDescription(oSocket.LastError)
        GoTo QH
    End If
    Text2.Text = "pvArraySize(baRecv)=" & pvArraySize(baRecv) & vbCrLf & DesignDumpMemory(VarPtr(baRecv(0)), pvArraySize(baRecv))
    If Not pvHandleInput(uCtx, baRecv, sError) Then
        GoTo QH
    End If
    baSend = pvGetClientHandshakeFinished(uCtx, sError)
    If pvArraySize(baSend) = 0 Then
        GoTo QH
    End If
    If Not oSocket.SyncSendArray(baSend) Then
        sError = oSocket.GetErrorDescription(oSocket.LastError)
        GoTo QH
    End If
    If Not pvDeriveApplicationSecrets(uCtx, sError) Then
        GoTo QH
    End If
    uCtx.State = ucsStatePostHandshake
    sRequest = "GET / HTTP/1.0" & vbCrLf & _
               "Host: localhost" & vbCrLf & vbCrLf
    baSend = pvGetClientApplicationData(uCtx, StrConv(sRequest, vbFromUnicode), sError)
    If pvArraySize(baSend) = 0 Then
        GoTo QH
    End If
    If Not oSocket.SyncSendArray(baSend) Then
        sError = oSocket.GetErrorDescription(oSocket.LastError)
        GoTo QH
    End If
    If Not oSocket.SyncReceiveArray(baRecv, Timeout:=-1) Then
        sError = oSocket.GetErrorDescription(oSocket.LastError)
        GoTo QH
    End If
    If Not pvHandleInput(uCtx, baRecv, sError) Then
        GoTo QH
    End If
    If uCtx.RecvPos > 0 Then
        Text2.Text = Replace(Replace(StrConv(uCtx.RecvBuffer, vbUnicode), vbCr, vbNullString), vbLf, vbCrLf)
    End If
QH:
    If LenB(sError) <> 0 Then
        MsgBox sError, vbCritical
    End If
End Sub

'Private Sub Command2_Click()
'    Dim uCtx            As UcsClientContextType
'    Dim lPos            As Long
'    Dim baHelloHash(0 To LNG_HASH_SIZE - 1)  As Byte
'
'    lPos = pvAppendArray(uCtx.HandshakeMessages, 0, FromHex("010000c60303000102030405060708090a0b0c0d0e0f101112131415161718191a1b1c1d1e1f20e0e1e2e3e4e5e6e7e8e9eaebecedeeeff0f1f2f3f4f5f6f7f8f9fafbfcfdfeff0006130113021303010000770000001800160000136578616d706c652e756c666865696d2e6e6574000a00080006001d00170018000d00140012040308040401050308050501080606010201003300260024001d0020358072d6365880d1aeea329adf9121383851ed21a28e3b75e965d0d2cd166254002d00020101002b0003020304"))
'    lPos = pvAppendArray(uCtx.HandshakeMessages, lPos, FromHex("020000760303707172737475767778797a7b7c7d7e7f808182838485868788898a8b8c8d8e8f20e0e1e2e3e4e5e6e7e8e9eaebecedeeeff0f1f2f3f4f5f6f7f8f9fafbfcfdfeff130100002e00330024001d00209fd7ad6dcff4298dd3f96d5b1b2af910a0535b1488d7f8fabb349a982880b615002b00020304"))
'    Call crypto_hash_sha256(baHelloHash(0), uCtx.HandshakeMessages(0), pvArraySize(uCtx.HandshakeMessages), 0)
''    baHelloHash = FromHex("da75ce1139ac80dae4044da932350cf65c97ccc9e33f1e6f7d2d4b18b736ffd5")
'    Debug.Print "HelloHash=0x" & ToHex(baHelloHash)
'
'    ReDim uCtx.ClientPrivate(0 To LNG_KEY_SIZE - 1) As Byte
'    ReDim uCtx.ClientPublic(0 To LNG_KEY_SIZE - 1) As Byte
'    ReDim baServerPrivate(0 To LNG_KEY_SIZE - 1) As Byte
'    ReDim uCtx.ServerPublic(0 To LNG_KEY_SIZE - 1) As Byte
'    ReDim baSharedSecret(0 To LNG_KEY_SIZE - 1) As Byte
'
''    Print "sodium_runtime_has_aesni=" & sodium_runtime_has_aesni()
''    Print "sodium_runtime_has_pclmul=" & sodium_runtime_has_pclmul()
''    Print "crypto_aead_aes256gcm_is_available=" & crypto_aead_aes256gcm_is_available()
''    Call randombytes_buf(uCtx.ClientPrivate(0), UBound(uCtx.ClientPrivate) + 1)
''    Call crypto_scalarmult_curve25519_base(uCtx.ClientPublic(0), uCtx.ClientPrivate(0))
''    Debug.Print "uCtx.ClientPrivate=0x" & ToHex(uCtx.ClientPrivate)
''    Debug.Print "uCtx.ClientPublic=0x" & ToHex(uCtx.ClientPublic)
'    uCtx.ClientPrivate = FromHex("202122232425262728292a2b2c2d2e2f303132333435363738393a3b3c3d3e3f")
'    Call crypto_scalarmult_curve25519_base(uCtx.ClientPublic(0), uCtx.ClientPrivate(0))
'    Debug.Print "ClientPrivate=0x" & ToHex(uCtx.ClientPrivate)
'    Debug.Print "ClientPublic=0x" & ToHex(uCtx.ClientPublic)
'    baServerPrivate = FromHex("909192939495969798999a9b9c9d9e9fa0a1a2a3a4a5a6a7a8a9aaabacadaeaf")
'    Call crypto_scalarmult_curve25519_base(uCtx.ServerPublic(0), baServerPrivate(0))
'    Debug.Print "ServerPrivate=0x" & ToHex(baServerPrivate)
'    Debug.Print "ServerPublic=0x" & ToHex(uCtx.ServerPublic)
'    Call crypto_scalarmult_curve25519(baSharedSecret(0), uCtx.ClientPrivate(0), uCtx.ServerPublic(0))
'    Debug.Print "SharedSecret=0x" & ToHex(baSharedSecret)
'
'
'    Dim baZeroKey(0 To LNG_KEY_SIZE - 1) As Byte
'    Dim baZeroSalt(0 To LNG_SALT_SIZE - 1) As Byte
'    Dim baEarlySecret() As Byte
'    Dim baEmptyHash(0 To LNG_HASH_SIZE - 1) As Byte
'    Dim baDerivedSecret() As Byte
'    Dim baHandshakeSecret() As Byte
'
'    baEarlySecret = pvHkdfExtract(baZeroSalt, baZeroKey)
'    Debug.Print "EarlySecret=0x" & ToHex(baEarlySecret)           ' 33AD0A1C607EC03B09E6CD9893680CE210ADF300AA1F2660E1B22E10F170F92A
'    Call crypto_hash_sha256(baEmptyHash(0), ByVal 0, 0, 0)
'    Debug.Print "EmptyHash=0x" & ToHex(baEmptyHash)               ' E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855
'    baDerivedSecret = pvHkdfExpand(baEarlySecret, "derived", baEmptyHash, 32)
'    Debug.Print "DerivedSecret=0x" & ToHex(baDerivedSecret)       ' 6F2615A108C702C5678F54FC9DBAB69716C076189C48250CEBEAC3576C3611BA
'    baHandshakeSecret = pvHkdfExtract(baDerivedSecret, baSharedSecret)
'    Debug.Print "HandshakeSecret=0x" & ToHex(baHandshakeSecret)   ' FB9FC80689B3A5D02C33243BF69A1B1B20705588A794304A6E7120155EDF149A
'
'
'    uCtx.ClientTrafficSecret = pvHkdfExpand(baHandshakeSecret, "c hs traffic", baHelloHash, LNG_HASH_SIZE)
'    Debug.Print "ClientTrafficSecret=0x" & ToHex(uCtx.ClientTrafficSecret)
'    uCtx.ServerTrafficSecret = pvHkdfExpand(baHandshakeSecret, "s hs traffic", baHelloHash, LNG_HASH_SIZE)
'    Debug.Print "ServerTrafficSecret=0x" & ToHex(uCtx.ServerTrafficSecret)
'    uCtx.ClientTrafficKey = pvHkdfExpand(uCtx.ClientTrafficSecret, "key", EmptyByteArray, 16)
'    Debug.Print "ClientTrafficKey=0x" & ToHex(uCtx.ClientTrafficKey)
'    uCtx.ServerTrafficKey = pvHkdfExpand(uCtx.ServerTrafficSecret, "key", EmptyByteArray, 16)
'    Debug.Print "ServerTrafficKey=0x" & ToHex(uCtx.ServerTrafficKey)
'    uCtx.ClientTrafficIV = pvHkdfExpand(uCtx.ClientTrafficSecret, "iv", EmptyByteArray, LNG_IV_SIZE)
'    Debug.Print "ClientTrafficIV=0x" & ToHex(uCtx.ClientTrafficIV)
'    uCtx.ServerTrafficIV = pvHkdfExpand(uCtx.ServerTrafficSecret, "iv", EmptyByteArray, LNG_IV_SIZE)
'    Debug.Print "ServerTrafficIV=0x" & ToHex(uCtx.ServerTrafficIV)
'End Sub

