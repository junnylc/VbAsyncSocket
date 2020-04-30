Attribute VB_Name = "mdTlsCrypto"
'=========================================================================
'
' Some of the cryptographic thunks are based on the following sources
'
'  1. https://github.com/esxgx/easy-ecc by Kenneth MacKay
'     which is distributed under the BSD 2-clause license
'
'  2. https://github.com/ctz/cifra by Joseph Birr-Pixton
'     CC0 1.0 Universal license (Public Domain Dedication)
'
'  3. https://github.com/github/putty by Simon Tatham
'     which is distributed under the MIT licence
'
'=========================================================================
Option Explicit
DefObj A-Z

#Const ImplUseLibSodium = (ASYNCSOCKET_USE_LIBSODIUM <> 0)
#Const ImplUseBCrypt = False

'=========================================================================
' API
'=========================================================================

Private Const TLS_SIGNATURE_RSA_PSS_RSAE_SHA256         As Long = &H804
Private Const TLS_SIGNATURE_RSA_PSS_RSAE_SHA384         As Long = &H805
Private Const TLS_SIGNATURE_RSA_PSS_RSAE_SHA512         As Long = &H806
Private Const TLS_SIGNATURE_RSA_PSS_PSS_SHA256          As Long = &H809
Private Const TLS_SIGNATURE_RSA_PSS_PSS_SHA384          As Long = &H80A
Private Const TLS_SIGNATURE_RSA_PSS_PSS_SHA512          As Long = &H80B
'--- for CryptAcquireContext
Private Const PROV_RSA_FULL                             As Long = 1
Private Const PROV_RSA_AES                              As Long = 24
Private Const CRYPT_NEWKEYSET                           As Long = &H8
Private Const CRYPT_DELETEKEYSET                        As Long = &H10
Private Const CRYPT_VERIFYCONTEXT                       As Long = &HF0000000
'--- for CryptDecodeObjectEx
Private Const X509_ASN_ENCODING                         As Long = 1
Private Const PKCS_7_ASN_ENCODING                       As Long = &H10000
Private Const X509_PUBLIC_KEY_INFO                      As Long = 8
Private Const PKCS_RSA_PRIVATE_KEY                      As Long = 43
Private Const PKCS_PRIVATE_KEY_INFO                     As Long = 44
Private Const X509_ECC_PRIVATE_KEY                      As Long = 82
Private Const CNG_RSA_PRIVATE_KEY_BLOB                  As Long = 83
Private Const CRYPT_DECODE_NOCOPY_FLAG                  As Long = &H1
Private Const CRYPT_DECODE_ALLOC_FLAG                   As Long = &H8000
'--- for CryptCreateHash
Private Const CALG_MD5                                  As Long = &H8003&
Private Const CALG_SHA1                                 As Long = &H8004&
Private Const CALG_SHA_256                              As Long = &H800C&
Private Const CALG_SHA_384                              As Long = &H800D&
Private Const CALG_SHA_512                              As Long = &H800E&
'--- for CryptSignHash
Private Const AT_KEYEXCHANGE                            As Long = 1
Private Const AT_SIGNATURE                              As Long = 2
Private Const RSA1024BIT_KEY                            As Long = &H4000000
Private Const MAX_RSA_KEY                               As Long = 8192     '--- in bits
'--- for CryptVerifySignature
Private Const NTE_BAD_SIGNATURE                         As Long = &H80090006
Private Const NTE_BAD_ALGID                             As Long = &H80090008
Private Const NTE_PROV_TYPE_NOT_DEF                     As Long = &H80090017
Private Const ERROR_FILE_NOT_FOUND                      As Long = 2
'--- for CertGetCertificateContextProperty
Private Const CERT_KEY_PROV_INFO_PROP_ID                As Long = 2
'--- for PFXImportCertStore
Private Const CRYPT_EXPORTABLE                          As Long = &H1
'--- for CryptExportKey
Private Const PUBLICKEYBLOB                             As Long = 6
Private Const PRIVATEKEYBLOB                            As Long = 7
'--- for CryptAcquireCertificatePrivateKey
Private Const CRYPT_ACQUIRE_CACHE_FLAG                  As Long = &H1
Private Const CRYPT_ACQUIRE_SILENT_FLAG                 As Long = &H40
Private Const CRYPT_ACQUIRE_ALLOW_NCRYPT_KEY_FLAG       As Long = &H10000
'Private Const CRYPT_ACQUIRE_PREFER_NCRYPT_KEY_FLAG      As Long = &H20000
'--- for NCryptImportKey
Private Const NCRYPT_OVERWRITE_KEY_FLAG                 As Long = &H80
Private Const NCRYPT_DO_NOT_FINALIZE_FLAG               As Long = &H400
'--- for NCryptSetProperty
Private Const NCRYPT_PERSIST_FLAG                       As Long = &H80000000
'--- for CertStrToName
Private Const CERT_OID_NAME_STR                         As Long = 2
'--- for CryptGetKeyParam
Private Const KP_KEYLEN                                 As Long = 9
'--- for CertOpenStore
Private Const CERT_STORE_PROV_MEMORY                    As Long = 2
Private Const CERT_STORE_CREATE_NEW_FLAG                As Long = &H2000
'--- for CertAddEncodedCertificateToStore
Private Const CERT_STORE_ADD_USE_EXISTING               As Long = 2
'--- for CertGetCertificateChain
Private Const CERT_TRUST_IS_NOT_TIME_VALID              As Long = &H1
Private Const CERT_TRUST_IS_NOT_TIME_NESTED             As Long = &H2
Private Const CERT_TRUST_IS_REVOKED                     As Long = &H4
Private Const CERT_TRUST_IS_NOT_SIGNATURE_VALID         As Long = &H8
Private Const CERT_TRUST_IS_UNTRUSTED_ROOT              As Long = &H20
Private Const CERT_TRUST_REVOCATION_STATUS_UNKNOWN      As Long = &H40
Private Const CERT_TRUST_IS_PARTIAL_CHAIN               As Long = &H10000
'--- for CertFindCertificateInStore
Private Const CERT_FIND_EXISTING                        As Long = &HD0000
'--- for CERT_ALT_NAME_ENTRY
Private Const CERT_ALT_NAME_DNS_NAME                    As Long = 3
'--- OIDs
Private Const szOID_RSA_RSA                             As String = "1.2.840.113549.1.1.1"
Private Const szOID_ECC_CURVE_P256                      As String = "1.2.840.10045.3.1.7"
Private Const szOID_ECC_CURVE_P384                      As String = "1.3.132.0.34"
Private Const szOID_ECC_CURVE_P521                      As String = "1.3.132.0.35"
Private Const szOID_PKCS_12_pbeWithSHA1And3KeyTripleDES As String = "1.2.840.113549.1.12.1.3"
Private Const szOID_SUBJECT_ALT_NAME2                   As String = "2.5.29.17"
'--- BLOBs magic
Private Const BCRYPT_RSAPRIVATE_MAGIC                   As Long = &H32415352
Private Const BCRYPT_ECDH_PRIVATE_P256_MAGIC            As Long = &H324B4345
Private Const BCRYPT_ECDH_PRIVATE_P384_MAGIC            As Long = &H344B4345
Private Const BCRYPT_ECDH_PRIVATE_P521_MAGIC            As Long = &H364B4345
'--- buffer types
Private Const NCRYPTBUFFER_PKCS_ALG_OID                 As Long = 41
Private Const NCRYPTBUFFER_PKCS_ALG_PARAM               As Long = 42
Private Const NCRYPTBUFFER_PKCS_KEY_NAME                As Long = 45
Private Const NCRYPTBUFFER_PKCS_SECRET                  As Long = 46
'--- export policy flags
Private Const NCRYPT_ALLOW_EXPORT_FLAG                  As Long = &H1
Private Const NCRYPT_ALLOW_PLAINTEXT_EXPORT_FLAG        As Long = &H2
'--- for thunks
Private Const MEM_COMMIT                                As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE                    As Long = &H40
#If ImplUseBCrypt Then
    '--- for BCryptSignHash
    Private Const BCRYPT_PAD_PSS                        As Long = 8
    '--- for BCryptVerifySignature
    Private Const STATUS_INVALID_SIGNATURE              As Long = &HC000A000
    Private Const ERROR_INVALID_DATA                    As Long = &HC000000D
#End If

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
'--- advapi32
Private Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextW" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGenRandom Lib "advapi32" (ByVal hProv As Long, ByVal dwLen As Long, ByVal pbBuffer As Long) As Long
Private Declare Function CryptImportKey Lib "advapi32" (ByVal hProv As Long, pbData As Any, ByVal dwDataLen As Long, ByVal hPubKey As Long, ByVal dwFlags As Long, phKey As Long) As Long
Private Declare Function CryptGenKey Lib "advapi32" (ByVal hProv As Long, ByVal AlgId As Long, ByVal dwFlags As Long, phKey As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32" (ByVal hProv As Long, ByVal AlgId As Long, ByVal hKey As Long, ByVal dwFlags As Long, phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32" (ByVal hHash As Long, pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32" (ByVal hHash As Long) As Long
Private Declare Function CryptSignHash Lib "advapi32" Alias "CryptSignHashA" (ByVal hHash As Long, ByVal dwKeySpec As Long, ByVal szDescription As Long, ByVal dwFlags As Long, pbSignature As Any, pdwSigLen As Long) As Long
Private Declare Function CryptVerifySignature Lib "advapi32" Alias "CryptVerifySignatureA" (ByVal hHash As Long, pbSignature As Any, ByVal dwSigLen As Long, ByVal hPubKey As Long, ByVal szDescription As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, pbData As Any, pdwDataLen As Long, dwBufLen As Long) As Long
'Private Declare Function CryptDecrypt Lib "advapi32" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, pbData As Any, pdwDataLen As Long) As Long
Private Declare Function CryptGetUserKey Lib "advapi32" (ByVal hProv As Long, ByVal dwKeySpec As Long, phUserKey As Long) As Long
Private Declare Function CryptExportKey Lib "advapi32" (ByVal hKey As Long, ByVal hExpKey As Long, ByVal dwBlobType As Long, ByVal dwFlags As Long, pbData As Any, pdwDataLen As Long) As Long
Private Declare Function CryptGetKeyParam Lib "advapi32" (ByVal hKey As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long
'--- Crypt32
Private Declare Function CryptImportPublicKeyInfo Lib "crypt32" (ByVal hCryptProv As Long, ByVal dwCertEncodingType As Long, pInfo As Any, phKey As Long) As Long
Private Declare Function CryptImportPublicKeyInfoEx2 Lib "crypt32" (ByVal dwCertEncodingType As Long, ByVal pInfo As Long, ByVal dwFlags As Long, ByVal pvAuxInfo As Long, phKey As Long) As Long
Private Declare Function CryptDecodeObjectEx Lib "crypt32" (ByVal dwCertEncodingType As Long, ByVal lpszStructType As Any, pbEncoded As Any, ByVal cbEncoded As Long, ByVal dwFlags As Long, ByVal pDecodePara As Long, pvStructInfo As Any, pcbStructInfo As Long) As Long
Private Declare Function CryptEncodeObjectEx Lib "crypt32" (ByVal dwCertEncodingType As Long, ByVal lpszStructType As Any, pvStructInfo As Any, ByVal dwFlags As Long, ByVal pEncodePara As Long, pvEncoded As Any, pcbEncoded As Long) As Long
Private Declare Function CryptAcquireCertificatePrivateKey Lib "crypt32" (ByVal pCert As Long, ByVal dwFlags As Long, ByVal pvParameters As Long, phCryptProvOrNCryptKey As Long, pdwKeySpec As Long, pfCallerFreeProvOrNCryptKey As Long) As Long
Private Declare Function PFXImportCertStore Lib "crypt32" (pPFX As Any, ByVal szPassword As Long, ByVal dwFlags As Long) As Long
Private Declare Function CertCreateCertificateContext Lib "crypt32" (ByVal dwCertEncodingType As Long, pbCertEncoded As Any, ByVal cbCertEncoded As Long) As Long
Private Declare Function CertFreeCertificateContext Lib "crypt32" (ByVal pCertContext As Long) As Long
Private Declare Function CertEnumCertificatesInStore Lib "crypt32" (ByVal hCertStore As Long, ByVal pPrevCertContext As Long) As Long
Private Declare Function CertGetCertificateContextProperty Lib "crypt32" (ByVal pCertContext As Long, ByVal dwPropId As Long, pvData As Any, pcbData As Long) As Long
Private Declare Function CertStrToName Lib "crypt32" Alias "CertStrToNameW" (ByVal dwCertEncodingType As Long, ByVal pszX500 As Long, ByVal dwStrType As Long, ByVal pvReserved As Long, pbEncoded As Any, pcbEncoded As Long, ByVal ppszError As Long) As Long
Private Declare Function CertCreateSelfSignCertificate Lib "crypt32" (ByVal hCryptProvOrNCryptKey As Long, pSubjectIssuerBlob As Any, ByVal dwFlags As Long, pKeyProvInfo As Any, ByVal pSignatureAlgorithm As Long, pStartTime As Any, pEndTime As Any, ByVal pExtensions As Long) As Long
Private Declare Function CertOpenStore Lib "crypt32" (ByVal lpszStoreProvider As Long, ByVal dwEncodingType As Long, ByVal hCryptProv As Long, ByVal dwFlags As Long, ByVal pvPara As Long) As Long
Private Declare Function CertCloseStore Lib "crypt32" (ByVal hCertStore As Long, ByVal dwFlags As Long) As Long
Private Declare Function CertAddEncodedCertificateToStore Lib "crypt32" (ByVal hCertStore As Long, ByVal dwCertEncodingType As Long, pbCertEncoded As Any, ByVal cbCertEncoded As Long, ByVal dwAddDisposition As Long, ByVal ppCertContext As Long) As Long
Private Declare Function CertCreateCertificateChainEngine Lib "crypt32" (pConfig As Any, phChainEngine As Long) As Long
Private Declare Function CertFreeCertificateChainEngine Lib "crypt32" (ByVal hChainEngine As Long) As Long
Private Declare Function CertGetCertificateChain Lib "crypt32" (ByVal hChainEngine As Long, ByVal pCertContext As Long, ByVal pTime As Long, ByVal hAdditionalStore As Long, pChainPara As Any, ByVal dwFlags As Long, ByVal pvReserved As Long, ppChainContext As Long) As Long
Private Declare Function CertFreeCertificateChain Lib "crypt32" (ByVal pChainContext As Long) As Long
Private Declare Function CertFindExtension Lib "crypt32" (ByVal pszObjId As String, ByVal cExtensions As Long, ByVal rgExtensions As Long) As Long
Private Declare Function CertFindCertificateInStore Lib "crypt32" (ByVal hCertStore As Long, ByVal dwCertEncodingType As Long, ByVal dwFindFlags As Long, ByVal dwFindType As Long, pvFindPara As Any, ByVal pPrevCertContext As Long) As Long
'--- NCrypt
Private Declare Function NCryptImportKey Lib "ncrypt" (ByVal hProvider As Long, ByVal hImportKey As Long, ByVal pszBlobType As Long, pParameterList As Any, phKey As Long, pbData As Any, ByVal cbData As Long, ByVal dwFlags As Long) As Long
Private Declare Function NCryptExportKey Lib "ncrypt" (ByVal hKey As Long, ByVal hExportKey As Long, ByVal pszBlobType As Long, pParameterList As Any, pbOutput As Any, ByVal cbOutput As Long, pcbResult As Any, ByVal dwFlags As Long) As Long
Private Declare Function NCryptFreeObject Lib "ncrypt" (ByVal hKey As Long) As Long
Private Declare Function NCryptGetProperty Lib "ncrypt" (ByVal hObject As Long, ByVal pszProperty As Long, pbOutput As Any, ByVal cbOutput As Long, pcbResult As Long, ByVal dwFlags As Long) As Long
Private Declare Function NCryptSetProperty Lib "ncrypt" (ByVal hObject As Long, ByVal pszProperty As Long, pbInput As Any, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function NCryptFinalizeKey Lib "ncrypt" (ByVal hKey As Long, ByVal dwFlags As Long) As Long
#If ImplUseBCrypt Then
    '--- BCrypt
    Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt" (ByRef hAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptImportKeyPair Lib "bcrypt" (ByVal hAlgorithm As Long, ByVal hImportKey As Long, ByVal pszBlobType As Long, ByRef hKey As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptDestroyKey Lib "bcrypt" (ByVal hKey As Long) As Long
    Private Declare Function BCryptSignHash Lib "bcrypt" (ByVal hKey As Long, pPaddingInfo As Any, pbInput As Any, ByVal cbInput As Long, pbOutput As Any, ByVal cbOutput As Long, pcbResult As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptVerifySignature Lib "bcrypt" (ByVal hKey As Long, pPaddingInfo As Any, pbHash As Any, ByVal cbHash As Long, pbSignature As Any, ByVal cbSignature As Long, ByVal dwFlags As Long) As Long
#End If
#If ImplUseLibSodium Then
    '--- libsodium
    Private Declare Function sodium_init Lib "libsodium" () As Long
    Private Declare Function randombytes_buf Lib "libsodium" (ByVal lpOut As Long, ByVal lSize As Long) As Long
    Private Declare Function crypto_scalarmult_curve25519 Lib "libsodium" (lpOut As Any, lpConstN As Any, lpConstP As Any) As Long
    Private Declare Function crypto_scalarmult_curve25519_base Lib "libsodium" (lpOut As Any, lpConstN As Any) As Long
    Private Declare Function crypto_hash_sha256 Lib "libsodium" (lpOut As Any, lpConstIn As Any, ByVal lSize As Long, Optional ByVal lHighSize As Long) As Long
    Private Declare Function crypto_hash_sha256_init Lib "libsodium" (lpState As Any) As Long
    Private Declare Function crypto_hash_sha256_update Lib "libsodium" (lpState As Any, lpConstIn As Any, ByVal lSize As Long, Optional ByVal lHighSize As Long) As Long
    Private Declare Function crypto_hash_sha256_final Lib "libsodium" (lpState As Any, lpOut As Any) As Long
    Private Declare Function crypto_hash_sha512_init Lib "libsodium" (lpState As Any) As Long
    Private Declare Function crypto_hash_sha512_update Lib "libsodium" (lpState As Any, lpConstIn As Any, ByVal lSize As Long, Optional ByVal lHighSize As Long) As Long
    Private Declare Function crypto_hash_sha512_final Lib "libsodium" (lpState As Any, lpOut As Any) As Long
    Private Declare Function crypto_aead_chacha20poly1305_ietf_decrypt Lib "libsodium" (lpOut As Any, lOutSize As Any, ByVal nSec As Long, lConstIn As Any, ByVal lInSize As Long, ByVal lHighInSize As Long, lpConstAd As Any, ByVal lAdSize As Long, ByVal lHighAdSize As Long, lpConstNonce As Any, lpConstKey As Any) As Long
    Private Declare Function crypto_aead_chacha20poly1305_ietf_encrypt Lib "libsodium" (lpOut As Any, lOutSize As Any, lConstIn As Any, ByVal lInSize As Long, ByVal lHighInSize As Long, lpConstAd As Any, ByVal lAdSize As Long, ByVal lHighAdSize As Long, ByVal nSec As Long, lpConstNonce As Any, lpConstKey As Any) As Long
    Private Declare Function crypto_aead_aes256gcm_is_available Lib "libsodium" () As Long
    Private Declare Function crypto_aead_aes256gcm_decrypt Lib "libsodium" (lpOut As Any, lOutSize As Any, ByVal nSec As Long, lConstIn As Any, ByVal lInSize As Long, ByVal lHighInSize As Long, lpConstAd As Any, ByVal lAdSize As Long, ByVal lHighAdSize As Long, lpConstNonce As Any, lpConstKey As Any) As Long
    Private Declare Function crypto_aead_aes256gcm_encrypt Lib "libsodium" (lpOut As Any, lOutSize As Any, lConstIn As Any, ByVal lInSize As Long, ByVal lHighInSize As Long, lpConstAd As Any, ByVal lAdSize As Long, ByVal lHighAdSize As Long, ByVal nSec As Long, lpConstNonce As Any, lpConstKey As Any) As Long
    Private Declare Function crypto_hash_sha512_statebytes Lib "libsodium" () As Long
#End If

Private Type CRYPT_BLOB_DATA
    cbData              As Long
    pbData              As Long
End Type

Private Type CRYPT_BIT_BLOB
    cbData              As Long
    pbData              As Long
    cUnusedBits         As Long
End Type

Private Type CRYPT_ALGORITHM_IDENTIFIER
    pszObjId            As Long
    Parameters          As CRYPT_BLOB_DATA
End Type

Private Type CERT_PUBLIC_KEY_INFO
    Algorithm           As CRYPT_ALGORITHM_IDENTIFIER
    PublicKey           As CRYPT_BIT_BLOB
End Type

Private Type CRYPT_ECC_PRIVATE_KEY_INFO
    dwVersion           As Long
    PrivateKey          As CRYPT_BLOB_DATA
    szCurveOid          As Long
    PublicKey           As CRYPT_BLOB_DATA
End Type

Private Type CRYPT_KEY_PROV_INFO
    pwszContainerName   As Long
    pwszProvName        As Long
    dwProvType          As Long
    dwFlags             As Long
    cProvParam          As Long
    rgProvParam         As Long
    dwKeySpec           As Long
End Type

Private Type CERT_CONTEXT
    dwCertEncodingType  As Long
    pbCertEncoded       As Long
    cbCertEncoded       As Long
    pCertInfo           As Long
    hCertStore          As Long
End Type

Private Type CRYPT_PRIVATE_KEY_INFO
    Version             As Long
    Algorithm           As CRYPT_ALGORITHM_IDENTIFIER
    PrivateKey          As CRYPT_BLOB_DATA
    pAttributes         As Long
End Type

Private Type CRYPT_PKCS12_PBE_PARAMS
    iIterations         As Long
    cbSalt              As Long
    SaltBuffer(0 To 31) As Byte
End Type

Private Type CERT_ALT_NAME_ENTRY
    dwAltNameChoice     As Long
    pwszDNSName         As Long
    Padding             As Long
End Type

Private Type CERT_ALT_NAME_INFO
    cAltEntry           As Long
    rgAltEntry          As Long
End Type

Private Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type

Private Type CERT_INFO
    dwVersion           As Long
    SerialNumber        As CRYPT_BLOB_DATA
    SignatureAlgorithm  As CRYPT_ALGORITHM_IDENTIFIER
    Issuer              As CRYPT_BLOB_DATA
    NotBefore           As FILETIME
    NotAfter            As FILETIME
    Subject             As CRYPT_BLOB_DATA
    SubjectPublicKeyInfo As CERT_PUBLIC_KEY_INFO
    IssuerUniqueId      As CRYPT_BIT_BLOB
    SubjectUniqueId     As CRYPT_BIT_BLOB
    cExtension          As Long
    rgExtension         As Long
End Type

Private Type CERT_TRUST_STATUS
    dwErrorStatus       As Long
    dwInfoStatus        As Long
End Type

Private Type CERT_CHAIN_CONTEXT
    cbSize              As Long
    TrustStatus         As CERT_TRUST_STATUS
    cElems              As Long
    rgElem              As Long
    '--- more here
End Type

Private Type CTL_USAGE
    cUsageIdentifier    As Long
    rgpszUsageIdentifier As Long
End Type

Private Type CERT_USAGE_MATCH
    dwType              As Long
    Usage               As CTL_USAGE
End Type

Private Type CERT_CHAIN_PARA
    cbSize              As Long
    RequestedUsage      As CERT_USAGE_MATCH
End Type

Private Type CERT_EXTENSION
    pszObjId            As Long
    fCritical           As Long
    Value               As CRYPT_BLOB_DATA
End Type

Private Type NCryptBuffer
    cbBuffer            As Long
    BufferType          As Long
    pvBuffer            As Long
End Type

Private Type NCryptBufferDesc
    ulVersion           As Long
    cBuffers            As Long
    pBuffers            As Long
    Buffers()           As NCryptBuffer
End Type

Private Type SYSTEMTIME
    wYear               As Integer
    wMonth              As Integer
    wDayOfWeek          As Integer
    wDay                As Integer
    wHour               As Integer
    wMinute             As Integer
    wSecond             As Integer
    wMilliseconds       As Integer
End Type

Private Type CERT_CHAIN_ENGINE_CONFIG
    cbSize              As Long
    hRestrictedRoot     As Long
    hRestrictedTrust    As Long
    hRestrictedOther    As Long
    cAdditionalStore    As Long
    rghAdditionalStore  As Long
    dwFlags             As Long
    dwUrlRetrievalTimeout As Long
    MaximumCachedCertificates As Long
    CycleDetectionModulus As Long
    '--- Win7+
    hExclusiveRoot      As Long
    hExclusiveTrustedPeople As Long
    dwExclusiveFlags    As Long
End Type

Private Type CERT_CHAIN_ELEMENT
    cbSize              As Long
    pCertContext        As Long
    TrustStatus         As CERT_TRUST_STATUS
    pRevocationInfo     As Long
    pIssuanceUsage      As Long
    pApplicationUsage   As Long
    pwszExtendedErrorInfo As Long
End Type

#If ImplUseBCrypt Then
    Private Type BCRYPT_PSS_PADDING_INFO
        pszAlgId        As Long
        cbSalt          As Long
    End Type
#End If

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_GLOB                  As String = "YHE5dsAbQHbAczl2AAAAAP///////////////wAAAAAAAAAAAAAAAAEAAAD/////S2DSJz48zjv2sFPMsAYdZbyGmHZVveuz55M6qtg1xlqWwpjYRTmh9KAz6y2BfQN38kCkY+XmvPhHQizh8tEXa/VRvzdoQLbLzl4xa1czzisWng98Suvnjpt/Gv7iQuNPUSVj/MLKufOEnhenrfrmvP//////////AAAAAP//////////AAAAAAAAAAD//////v/////////////////////////////////////////vKuzT7ciFKp3RLoqNOVbGWocTUI8IFAMSQYH+bpwdGBkt+ONrBY6Y5Oc+4qcvMbO3CnZyOF5UOmwpVb9d8gJVOCpUguBB91mYm6eLYjsdbnStIPMex7GONwWLviLKh6pfDuqQfB1Dep2Bfh3OsWAKwLjwtRMx2ul8FJoovR30+CnckpK/mJ5dbywmlkreFzZzKcXMahns7HqnsEiyDRpY3y039IFNY8f///////////////////////////////+YL4pCkUQ3cc/7wLWl27XpW8JWOfER8Vmkgj+S1V4cq5iqB9gBW4MSvoUxJMN9DFV0Xb5y/rHegKcG3Jt08ZvBwWmb5IZHvu/GncEPzKEMJG8s6S2qhHRK3KmwXNqI+XZSUT6YbcYxqMgnA7DHf1m/8wvgxkeRp9VRY8oGZykpFIUKtyc4IRsu/G0sTRMNOFNUcwpluwpqdi7JwoGFLHKSoei/oktmGqhwi0vCo1FsxxnoktEkBpnWhTUO9HCgahAWwaQZCGw3Hkx3SCe1vLA0swwcOUqq2E5Pypxb828uaO6Cj3RvY6V4FHjIhAgCx4z6/76Q62xQpPej+b7yeHHGIq4o" & _
                                                    "15gvikLNZe8jkUQ3cS87TezP+8C1vNuJgaXbtek4tUjzW8JWORnQBbbxEfFZm08Zr6SCP5IYgW3a1V4cq0ICA6OYqgfYvm9wRQFbgxKMsuROvoUxJOK0/9XDfQxVb4l78nRdvnKxlhY7/rHegDUSxyWnBtyblCZpz3Txm8HSSvGewWmb5OMlTziGR77vtdWMi8adwQ9lnKx3zKEMJHUCK1lvLOktg+SmbqqEdErU+0G93KmwXLVTEYPaiPl2q99m7lJRPpgQMrQtbcYxqD8h+5jIJwOw5A7vvsd/Wb/Cj6g98wvgxiWnCpNHkafVb4ID4FFjygZwbg4KZykpFPwv0kaFCrcnJskmXDghGy7tKsRa/G0sTd+zlZ0TDThT3mOvi1RzCmWosnc8uwpqduau7UcuycKBOzWCFIUscpJkA/FMoei/ogEwQrxLZhqokZf40HCLS8IwvlQGo1FsxxhS79YZ6JLREKllVSQGmdYqIHFXhTUO9LjRuzJwoGoQyNDSuBbBpBlTq0FRCGw3Hpnrjt9Md0gnqEib4bW8sDRjWsnFswwcOcuKQeNKqthOc+Njd0/KnFujuLLW828uaPyy713ugo90YC8XQ29jpXhyq/ChFHjIhOw5ZBoIAseMKB5jI/r/vpDpvYLe62xQpBV5xrL3o/m+K1Ny4/J4ccacYSbqzj4nygfCwCHHuIbRHuvgzdZ92up40W7uf0999bpvF3KqZ/AGppjIosV9YwquDfm+BJg/ERtHHBM1C3EbhH0EI/V32yiTJMdAe6vKMry+yRUKvp48TA0QnMRnHUO2Qj7LvtTFTCp+ZfycKX9Z7PrWOqtvy18XWEdKjBlEbGV4cGFuZCAxNi1ieXRlIGsAZXhwYW5k" & _
                                                    "IDMyLWJ5dGUgawAAAAUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD8AAAAY3x3e/Jrb8UwAWcr/terdsqCyX36WUfwrdSir5ykcsC3/ZMmNj/3zDSl5fFx2DEVBMcjwxiWBZoHEoDi6yeydQmDLBobblqgUjvWsynjL4RT0QDtIPyxW2rLvjlKTFjP0O+q+0NNM4VF+QJ/UDyfqFGjQI+SnTj1vLbaIRD/89LNDBPsX5dEF8Snfj1kXRlzYIFP3CIqkIhG7rgU3l4L2+AyOgpJBiRcwtOsYpGV5HnnyDdtjdVOqWxW9Opleq4IunglLhymtMbo3XQfS72LinA+tWZIA/YOYTVXuYbBHZ7h+JgRadmOlJseh+nOVSjfjKGJDb/mQmhBmS0PsFS7Fo0BAgQIECBAgBs2Uglq1TA2pTi/QKOegfPX+3zjOYKbL/+HNI5DRMTe6ctUe5QypsIjPe5MlQtC+sNOCC6hZijZJLJ2W6JJbYvRJXL49mSGaJgW1KRczF1ltpJscEhQ/e252l4VRlenjZ2EkNirAIy80wr35FgFuLNFBtAsHo/KPw8Cwa+9AwETims6kRFBT2fc6pfyz87wtOZzlqx0IuetNYXi+TfoHHXfbkfxGnEdKcWJb7diDqoYvhv8Vj5LxtJ5IJrbwP54zVr0H92oM4gHxzGxEhBZJ4DsX2BRf6kZtUoNLeV6n5PJnO+g4DtNrir1sMjruzyDU5lhFysEfrp31ibhaRRjVSEMfQAAAAAAAQAAAAEAAAA=" ' 1952, 30.4.2020 15:27:45
Private Const STR_THUNK1                As String = "OCK6AIAnAACgKgAA4D8AAKBGAABARwAA0EgAABBOAAAwPwAAMEYAAABHAACARwAAEEoAACA0AABwNAAAUDIAAOA0AABwNQAAsDQAAAA5AACQOQAAgDUAAKAmAABgJgAAsBsAADAbAACwfQAAzMzMzOgAAAAAWC11QLkABQBAuQCLAMPMzMzMzMzMzMzMzMzM6AAAAABYLZVAuQAFAEC5AMPMzMzMzMzMzMzMzMzMzMxVi+yD7GhTi10QU+hgjQAAhcAPhVsBAABWi3UMjUXIV1ZQ6KmdAACLfQiNRchQV41FmFDoOJ0AAI1FyFBQ6I6dAABTVlboJp0AAFNT6H+dAADoav///wWwAAAAUFNXV+gMlQAA6Ff///8FsAAAAFBTU1Po+ZQAAOhE////BbAAAABQU1dT6KadAABTV1fo3pwAAOgp////BbAAAABQV1dT6MuUAADoFv///wWwAAAAUFNXV+i4lAAAagBX6KCpAAALwnQl6Pf+//8FsAAAAFBXV+iKhwAAV4vw6HKhAADB5h8JdyyLdQzrBlfoYaEAAFdT6NqcAADoxf7//wWwAAAAUI1FmFBTU+gknQAA6K/+//8FsAAAAFCNRZhQU1PoDp0AAOiZ/v//BbAAAABQU41FmFBQ6PicAACNRZhQV1foLZwAAOh4/v//BbAAAABQjUXIUFdQ6NecAABTV+iwoQAAVlPoqaEAAI1FyFBW6J+hAABfXluL5V3CDADMzMzMzMxVi+yD7EhTi10QU+gQjAAAhcAPhUcBAABWi3UMjUXYV1ZQ6FmcAACLfQiNRdhQV41FuFDo6JsAAI1F2FBQ6D6cAABTVlbo1psAAFNT6C+cAADo6v3//4PAEFBTV1fozpMAAOjZ/f//g8AQUFNTU+i9kwAA" & _
                                                    "6Mj9//+DwBBQU1dT6FycAABTV1folJsAAOiv/f//g8AQUFdXU+iTkwAA6J79//+DwBBQU1dX6IKTAABqAFfoKqgAAAvCdCPogf3//4PAEFBXV+hWiAAAV4vw6F6gAADB5h8JdxyLdQzrBlfoTaAAAFdT6JabAADoUf3//4PAEFCNRbhQU1Po4psAAOg9/f//g8AQUI1FuFBTU+jOmwAA6Cn9//+DwBBQU41FuFBQ6LqbAACNRbhQV1fo75oAAOgK/f//g8AQUI1F2FBXUOibmwAAU1fopKAAAFZT6J2gAACNRdhQVuiToAAAX15bi+VdwgwAzMzMzMzMzMzMzFWL7FaLdQhW6HOKAACFwHQXjUYwUOhmigAAhcB0CrgBAAAAXl3CBAAzwF5dwgQAzFWL7FaLdQhW6HOKAACFwHQXjUYgUOhmigAAhcB0CrgBAAAAXl3CBAAzwF5dwgQAzFWL7IHs+AAAAFOLXQyNRZhWV1NQ6KefAACNQzBQiUX4jYU4////UOiUnwAA/3UUjYUI////UI2FaP///1CNhTj///9QjUWYUOgDCAAAi10QU+iqnQAAjXD+hfZ+YA8fAFZT6KmmAAALwnUHuAEAAADrAjPAjQRAweAEjY0I////A8iNlWj///8D0IlNFFH32IlV/I29OP///1ID+I1dmAPYV1PoqAQAAFdT/3UU/3X86JsCAACLXRBOhfZ/o2oAU+hLpgAAC8J1B7gBAAAA6wIzwI0EQMHgBI2dCP///wPYjY1o////UwPIjb04////USv4iU0QjXWYK/BXVuhPBAAA6Gr7//8FsAAAAFCNhWj///9QjUWYUI1FyFDowJkAAFeNRchQUOj1mAAA/3UMjUXIUFDo6JgA" & _
                                                    "AOgz+///BbAAAABQjUXIUFDoU5EAAP91+I1FyFBQ6MaYAABWjUXIUFDou5gAAFdWU/91EOjgAQAAjUXIUI2FCP///1CNhWj///9Q6IkKAACLdQiNhWj///9QVugpngAAjYUI////UI1GMFDoGZ4AAF9eW4vlXcIQAFWL7IHsqAAAAFOLXQyNRbhWV1NQ6FeeAACNQyBQiUX4jYV4////UOhEngAA/3UUjYVY////UI1FmFCNhXj///9QjUW4UOjWBgAAi10QU+hNnAAAg+gCiUUUhcB+Ww8fAFBT6PmkAAALwnUHuAEAAADrAjPAweAFjZ1Y////A9iNTZgDyI21eP///1P32IlN/FED8I19uAP4VlfokQQAAFZXU/91/Oj2AQAAi0UUi10QSIlFFIXAf6hqAFPooKQAAAvCdQWNSAHrAjPJweEFjZ1Y////A9mJTRBTjUWYA8GNvXj///9QK/mNdbgr8VdW6DwEAADox/n//4PAEFCNRZhQjUW4UI1F2FDoUpgAAFeNRdhQUOiHlwAA/3UMjUXYUFDoepcAAOiV+f//g8AQUI1F2FBQ6BeSAAD/dfiNRdhQUOhalwAAVo1F2FBQ6E+XAABXVo1FmANFEFNQ6EABAACNRdhQjYVY////UI1FmFDoPAkAAIt1CI1FmFBW6O+cAACNhVj///9QjUYgUOjfnAAAX15bi+VdwhAAzMzMzMzMVYvsg+wwU1ZX6BL5//+LXQgFsAAAAIt1EFBTVo1F0FDoa5cAAI1F0FBQ6AGXAACNRdBQU1PolpYAAI1F0FBWVuiLlgAA6Nb4//+LdQwFsAAAAIt9FFBWV1foMpcAAFeNRdBQ6MiWAADos/j//wWwAAAAUFONRdBQUOgS" & _
                                                    "lwAA6J34//8FsAAAAFCLRRBQjUXQUFDo+ZYAAOiE+P//BbAAAABQi0UQU1BQ6OOWAACLRRBQVlboGJYAAOhj+P//BbAAAABQjUXQUFOLXRBT6L+WAABTV1fo95UAAOhC+P//BbAAAABQVldX6KSWAACNRdBQU+h6mwAAX15bi+VdwhAAzFWL7IPsIFNWV+gS+P//i10Ig8AQi3UQUFNWjUXgUOidlgAAjUXgUFDoM5YAAI1F4FBTU+jIlQAAjUXgUFZW6L2VAADo2Pf//4t1DIPAEIt9FFBWV1foZpYAAFeNReBQ6PyVAADot/f//4PAEFBTjUXgUFDoSJYAAOij9///g8AQUItFEFCNReBQUOgxlgAA6Iz3//+DwBBQi0UQU1BQ6B2WAACLRRBQVlboUpUAAOht9///g8AQUI1F4FBTi10QU+j7lQAAU1dX6DOVAADoTvf//4PAEFBWV1fo4pUAAI1F4FBT6OiaAABfXluL5V3CEADMzMzMzMzMzMzMzMzMzMxVi+yB7JAAAABTVlfoD/f//4tdCAWwAAAAi30QUFNXjUWgUOholQAAjUWgUFDo/pQAAI1FoFBTU+iTlAAAjUWgUFdX6IiUAADo0/b//4tdDAWwAAAAi3UUUFNWjUWgUOhsjAAA6Lf2//8FsAAAAFBTVlboGZUAAOik9v//BbAAAABQ/3UIjUXQV1DoAZUAAI1F0FBTU+g2lAAA6IH2//8FsAAAAFBX/3UIjUXQUOgejAAAVlfod5QAAOhi9v//BbAAAABQjUXQUFdX6MGUAADoTPb//wWwAAAAUFeLfQiNhXD///9XUOillAAAjYVw////UFZW6NeTAADoIvb//wWwAAAAUFNWVuiElAAAjUWg" & _
                                                    "UI2FcP///1DoFJQAAOj/9f//BbAAAABQjUXQUI2FcP///1BQ6FiUAADo4/X//wWwAAAAUFeNhXD///9QjUXQUOg8lAAAjUWgUI1F0FBQ6G6TAADoufX//wWwAAAAUFONRdBQU+gYlAAAjYVw////UFfo65gAAF9eW4vlXcIQAMzMVYvsg+xgU1ZX6IL1//+LXQiDwBCLfRBQU1eNRcBQ6A2UAACNRcBQUOijkwAAjUXAUFNT6DiTAACNRcBQV1foLZMAAOhI9f//i10Mg8AQi3UUUFNWjUXAUOgjiwAA6C71//+DwBBQU1ZW6MKTAADoHfX//4PAEFD/dQiNReBXUOiskwAAjUXgUFNT6OGSAADo/PT//4PAEFBX/3UIjUXgUOjbigAAVlfoJJMAAOjf9P//g8AQUI1F4FBXV+hwkwAA6Mv0//+DwBBQV4t9CI1FoFdQ6FmTAACNRaBQVlbojpIAAOip9P//g8AQUFNWVug9kwAAjUXAUI1FoFDo0JIAAOiL9P//g8AQUI1F4FCNRaBQUOgZkwAA6HT0//+DwBBQV41FoFCNReBQ6AKTAACNRcBQjUXgUFDoNJIAAOhP9P//g8AQUFONReBQU+jgkgAAjUWgUFfo5pcAAF9eW4vlXcIQAMzMzMzMzMzMzMzMzMxVi+yD7DBWi3UIV1b/dRDoXJcAAIt9DFf/dRToUJcAAI1F0FDoF4AAAItFGMdF0AEAAADHRdQAAAAAhcB0ClCNRdBQ6CiXAACNRdBQV1bobQMAAI1F0FBXVugC9P//jUXQUP91FP91EOhTAwAAX16L5V3CFADMzMzMzMzMzMzMzFWL7IPsIFaLdQhXVv91EOg8lwAAi30MV/91FOgwlwAAjUXg" & _
                                                    "UOj3fwAAi0UYx0XgAQAAAMdF5AAAAACFwHQKUI1F4FDoCJcAAI1F4FBXVug9AwAAjUXgUFdW6AL1//+NReBQ/3UU/3UQ6CMDAABfXovlXcIUAMzMzMzMzMzMzMzMU4tEJAyLTCQQ9+GL2ItEJAj3ZCQUA9iLRCQI9+ED01vCEADMzMzMzMzMzMzMzMzMgPlAcxWA+SBzBg+lwtPgw4vQM8CA4R/T4sMzwDPSw8yA+UBzFYD5IHMGD63Q0+rDi8Iz0oDhH9PowzPAM9LDzFWL7ItFEFNWi3UIjUh4V4t9DI1WeDvxdwQ70HMLjU94O/F3MDvXciwr+LsQAAAAK/CLFDgDEItMOAQTSASNQAiJVDD4iUww/IPrAXXkX15bXcIMAIvXjUgQi94r0CvYK/64BAAAAI12II1JIA8QQdAPEEw34GYP1MgPEU7gDxBMCuAPEEHgZg/UyA8RTAvgg+gBddJfXltdwgwAzMzMzMxVi+yLVRyD7AiLRSBWi3UIV4t9DAPXE0UQiRaJRgQ7RRB3D3IEO9dzCbgBAAAAM8nrDg9XwGYPE0X4i038i0X4A0UkXxNNKANFFIlGCIvGE00YiU4MXovlXcIkAMzMzMxVi+yLVQyLTQiLAjEBi0IEMUEEi0IIMUEIi0IMMUEMXcIIAMzMzMzMzMzMzMzMzMxVi+yD7AiLTQiLVRBTVosBjVkEweoCM/aJVRCJXfiNBIUEAAAAiUX8V4XSdEKLVQyLfRCDwgJmZg8fhAAAAAAAD7ZK/o1SBA+2QvvB4QgLyA+2QvzB4QgLyA+2Qv3B4QgLyIkMs0Y793LWi0X8i9e5AQAAADP/iU0MO/APg40AAACLxivCjQSDiUUIDx9EAACLXLP8O/p1" & _
                                                    "CEEz/4lNDOsEhf91LejX8P//BYgFAADBwwhQU+hIeAAAi9jowfD//4tNDA+2hAiIBgAAweAYM9jrHYP6BnYeg/8EdRnooPD//wWIBQAAUFPoFHgAAIvYi0UIi1UQiwhHM8uDwASLXfiJRQiJDLNGi00MO3X8coJfXluL5V3CDADMzMzMzMzMzMxVi+yD7DCNRdD/dRBQ6F6OAACNRdBQi0UIUFDo8I0AAP91EI1F0FBQ6OONAACNRdBQi0UMUFDo1Y0AAIvlXcIMAMzMzMzMzMzMzMzMzMzMzFWL7IPsII1F4P91EFDoPo4AAI1F4FCLRQhQUOjQjQAA/3UQjUXgUFDow40AAI1F4FCLRQxQUOi1jQAAi+VdwgwAzMzMzMzMzMzMzMzMzMzMVYvsg+wMU1aLdQxXiz6JffiNHL0AAAAAiV306KDv//9TiwD/0DPSiUX8hf9+EWaQiw4ryosMjokMkEI713zxi10IiwOJRQg7x38GjUcBiUUIjTSFAAAAAOhj7///VosA/9CLVQiJRQyF0n4Si30Mi87B6QIzwPOri334i0UMuQEAAAA5C3wXjVD8A9YPH0AAiwSLjVL8iUIEQTsLfvKLXfyLG1PoBwUAAIvQhdJ0GovK0+OD/wF+EYtF/LkgAAAAK8qLQATT6AvYUlPogGsAAFD/dRRXi338V/91CP91DOisQwAAi10Qhdt0LrgBAAAAOQN8JYt9DItVCIPH/EoD/oXSeASLD+sCM8mJDINKQIPvBDsDfuqLffyLTfSFyXQNi8fGAACNQAGD6QF19eiO7v//V4tACP/Qi10Mhdt0EYX2dA2Lw8YAAI1AAYPuAXX16Gvu//9Ti0AI/9BfXluL5V3CEADMzMzMzMzM" & _
                                                    "zMzMzMxVi+xTi10MVleLG408nQQAAADoOe7//1eLAP/QV4vwagBW6OlLAACDxAyJHmoAVv91DP91COhW/v//gz4BdhGQiwaDPIYAdQhIiQaD+AF38F+Lxl5bXcIIAMzMzMzMzMxVi+yD7ByLRQhTVleLCItFDIlN9IsAO8iL2IlF7A9P2VOJXfCNPJ0AAAAA6JBgAAADx400hQAAAACJdeTor+3//1aLAP/Qi9Az9olV+IXbflWLRQwD+ovLjRSYi0UIiUX8i0UMKUX8i0UIOwh/CItF/IsEEOsCM8CLXfiJBLOLRQw7CH8EiwLrAjPAi13wRokHg+oEg8cESTvzfMuLVfiNPJ0AAAAAi8PB4AQDwlCJRQiNBNpTUI0EF1BS6MBEAACLTfSLRRBBA03siU0MhcB0DIsAO8h/Bo1IAYlNDI00jQQAAADoB+3//1aLAP/QVov4agBX6LdKAACLRQwzyYPEDIkHjVEBO8J8JYt1CAPbg8b8O9N/BIsG6wIzwIXAiQSXD0XKQoPuBDsXfuaLRQyLdRCJD4X2dHu6AQAAAA9XwGYPE0XoO8J8aovGjV8Ei3XsK8eJdfSLdeiJdfyLdRCJRfA7F38KiwOJRQiLRfDrB8dFCAAAAAA7Fn8FizQY6wIz9jPAA3UIE8ADdfyJMxNF9IlF/MdF9AAAAACF9nQFO9EPT8qLdRBCi0Xwg8MEO1UMfq+LXfiJD4XbdBSLTeSFyXQNi8PGAACNQAGD6QF19egb7P//U4tACP/Qi8dfXluL5V3CDADMzMzMzMzMzMzMVYvsi00MVoXJeDWLdQiLBsHgAjvIfSmLwZmD4gMDwsH4AoHhAwAAgHkFSYPJ/EGLRIYEweED0+gPtsBeXcII" & _
                                                    "ADPAXl3CCADMzMzMzMzMzMxVi+yLVQhTVot1DFeLOoseg/8BdQgzwDlCBA9E+IP7AXUIM8A5RgQPRNg7+4vDD0/HhcB0MyvWjQyGiVUIDx8AO8d+BDPS6wOLFAo7w34EM/brBosxO9ZyJDvWdxSLVQiD6QSD6AF12F9eM8BbXcIIAF9euAEAAABbXcIIAF9eg8j/W13CCADMzMzMzMzMzMxVi+yLRQyDwAOZg+IDU1ZXjRwCwfsCjTydBAAAAOj+6v//V4sI/9FXi/BqAFborkgAAIPEDIkeg/sBfBGNDJ0AAAAAwekCjX4EM8Dzq4tdDIXbdEeNDN0AAAAADx9EAACLRQiD6QhLiU0MihBAiUUIi8OFwHkDg8ADwfgCjTyGgeEfAACAeQVJg8ngQQ+2wtPgCUcEi00Mhdt1xYM+AXYQiwaDPIYAdQhIiQaD+AF38F+Lxl5bXcIIAMzMzMzMzMxVi+yLTQyFyXhCVot1CIsGweAFO8h9NIvBmYPiHwPCwfgFgeEfAACAeQVJg8ngQboBAAAA0+KDfRAAdAkJVIYEXl3CDAD30iFUhgReXcIMAFWL7FaLdQi6EAAAAFcz/5C5IAAAAIvGK8rT6IXAdQaLytPmA/rR+nXni8dfXl3CBADMzMzMzMzMzMzMzMzMzMxVi+yLRQiZU4PiH1ZXjRwCwfsFQ400nQQAAADosOn//1aLCP/RVov4agBX6GBHAACDxAyJH2oB/3UIV+gw////i8dfXltdwgQAzMzMzMzMzFWL7IPsIFNWi3UIM8lXiU3sgQTOAAABAIsEzoNUzgQAi1zOBA+s2BDB+xCJReiD+Q91FcdF/AEAAACL0MdF8AAAAACJXfjrIg9XwGYPE0X0i0X4" & _
                                                    "iUXwi0X0Zg8TReCLVeCJRfyLReSJRfiD+Q+NeQFqABvA99gPr8crVfxqJY00xotF+BtF8FBS6NL1//+LTegDwRPTg+gBg9oAAQaLRewRVgSLdQgPpMsQweEQKQzGi8+JTewZXMYEg/kQD4JP////X15bi+VdwgQAzMzMzMxVi+yD7BCLVQxWVw+2Cg+2QgHB4QgLyA+2QgLB4QgLyA+2QgPB4QgLyA+2QgWJTfAPtkoEweEIC8gPtkIGweEIC8gPtkIHweEIC8gPtkIJiU30D7ZKCMHhCAvID7ZCCsHhCAvID7ZCC8HhCAvID7ZCDIlN+A+2Sg3B4AgLyA+2Qg7B4QgLyA+2Qg/B4QgLyIlN/ItNCIs5jXEEi8fB4AQD8I1F8FZQ6FX2//+D7hCDx/90LY1F8FDodEQAAI1F8FDoC0UAAFaNRfBQ6DH2//+NRfBQ6BhEAACD7hCD7wF1041F8FDoR0QAAI1F8FDo3kQAAFaNRfBQ6AT2//+LdRCLVfCLwotN9MHoGIgGi8LB6BCIRgGLwsHoCIhGAovBwegYiFYDiEYEi8HB6BCIRgWLwcHoCIhGBohOB4tN+IvBwegYiEYIi8HB6BCIRgmLwcHoCIhGCohOC4tN/IvBwegYiEYMi8HB6BCIRg2LwcHoCIhGDl+ITg9ei+VdwgwAzMxVi+yD7BBTVleLVQyLXQgPtgoPtkIBweEIjXMEC8gPtkICweEIC8gPtkIDweEIC8gPtkIFiU3wD7ZKBMHhCAvID7ZCBsHhCAvID7ZCB8HhCAvID7ZCCYlN9A+2SgjB4QgLyA+2QgrB4QgLyA+2QgvB4QgLyA+2QgyJTfgPtkoNweAIC8gPtkIOweEIC8gPtkIPweEIC8iN" & _
                                                    "RfBWUIlN/Ojd9P//vwEAAACDxhA5O3YukI1F8FDoh20AAI1F8FDoHmwAAI1F8FDohUQAAFaNRfBQ6Kv0//9Hg8YQOzty041F8FDoWm0AAI1F8FDo8WsAAFaNRfBQ6If0//+LdRCLVfCLwotN9MHoGIgGi8LB6BCIRgGLwsHoCIhGAovBwegYiFYDiEYEi8HB6BCIRgWLwcHoCIhGBohOB4tN+IvBwegYiEYIi8HB6BCIRgmLwcHoCIhGCohOC4tN/IvBwegYiEYMi8HB6BCIRg2LwcHoCIhGDl+ITg9eW4vlXcIMAMzMzMxVi+xWi3UIaPQAAABqAFboXEMAAItFEIPEDIP4EHQ1g/gYdBqD+CB1PFD/dQzHBg4AAABW6Pfz//9eXcIMAGoY/3UMxwYMAAAAVujh8///Xl3CDABqEP91DMcGCgAAAFboy/P//15dwgwAzMzMzMzMVYvsgewAAQAAVuhR5f//vkBZuQCB7gBAuQAD8Og/5f///3UoucBXuQDHRfQQAAAA/3UkgekAQLkAiXX4A8GJRfyNhQD///9Q6EP/////dQiNhQD///9qEP91FGoM/3Ug/3Uc/3UY/3UQ/3UMUI1F9FDoOg8AAF6L5V3CJADMzMxVi+yB7AABAABW6NHk//++QFm5AIHuAEC5AAPw6L/k////dSi5wFe5AMdF9BAAAAD/dSSB6QBAuQCJdfgDwYlF/I2FAP///1Dow/7//2oQ/3UMjYUA/////3UIagz/dSD/dRz/dRj/dRT/dRBQjUX0UOh6EAAAXovlXcIkAMzMzFWL7FFTi10YM8CJRfyF23Rxi1UQi00MVsdFGAEAAABXizmL8iv3O94PQvOFwHUdD7ZFFFZQi0UIA8dQ" & _
                                                    "6MBBAACLTQyDxAyLRfyLVRCF/3UJO/IPREUYiUX8jQQ+O8J1F/91CP91IP9VHItNDItVEMcBAAAAAOsCATGLRfwr3nWgX15bi+VdwhwAzMzMzMzMzFWL7FaLdSCLxoPoAHRgg+gBD4SsAAAAU4PoAVeNRRR0bYt9KItdJFdTagFQ/3UQ/3UM/3UI6LYAAACLTRhXUzhNHHQvjUb+i3UQUFFW/3UM/3UI6Bj///9XU2oBjUUcUFb/dQz/dQjohAAAAF9bXl3CJACNRv+LdRBQUVb/dQz/dQjo6f7//19bXl3CJAD/dSiLXRD/dSSLfQyLdQhqAVBTV1boSAAAAP91KI1FHP91JGoBUFNXVug0AAAAX1teXcIkAP91KIpFHP91JDBFFI1FFGoBUP91EP91DP91COgNAAAAXl3CJADMzMzMzMzMzFWL7P91IItFHFBQ/3UY/3UU/3UQ/3UM/3UI6BEAAABdwhwAzMzMzMzMzMzMzMzMzFWL7ItNDItFJFOLXRSLEVaLdRhXhdJ0WYX2dFWLRRCL/ivCO8YPQviLwgNFCFdTUOjrPwAAi0UMA98r94PEDAE4i30QOTiLRSR1Kf91CFCF9nUN/1Ugi00Mi0UkiTHrFP9VHItNDItFJMcBAAAAAOsDi30QO/dyGVNQO/d1Bf9VIOsD/1Uci0UkK/cD3zv3c+eF9nQui0UMiwiLxyvBi/47xg9C+ItFCFcDwVNQ6G4/AACLRQwD34PEDAE4K/eLfRB11V9eW13CIADMzMzMzMxVi+yLTRyD7AhXi30Yhcl0dlOLXQxWgzsAdRH/dQj/dST/VSCLRRCLTRyJA4sDi/GLVRAr0DvBiVUYD0LwM8CJdfyF9nQvi10UK9+JXfhm" & _
                                                    "kIt1/I0UOIoME4tVGANVCItd+DIMAo0UOECICjvGcuGLXQyLTRwpMyvOAXUUA/6JTRyFyXWRXltfi+VdwiAAzMxVi+zoSOH//7kgZrkAgekAQLkAA8GLTQhRUP91FI1BdP91EP91DGpAUI1BNFDoPv///13CEADMzMzMzMzMzMzMVYvsg+xsi00UU1ZXD7ZZAw+2QQIPtlEHweIIweMIC9gPtkEBweMIC9gPtgHB4wgL2A+2QQYL0Ild2MHiCA+2QQUL0A+2QQTB4ggL0A+2QQqJVfSJVdQPtlELweIIC9APtkEJweIIC9APtkEIweIIC9APtkEOiVXwiVXQD7ZRD8HiCAvQD7ZBDcHiCAvQD7ZBDItNCMHiCAvQiVX4D7ZBAolVzA+2UQPB4ggL0A+2QQHB4ggL0A+2AcHiCAvQD7ZBBolV7IlVyA+2UQfB4ggL0A+2QQXB4ggL0A+2QQTB4ggL0A+2QQqJVeiJVcQPtlELweIIC9DB4ggPtkEJC9APtkEIweIIC9APtkEOiVXkiVXAD7ZRD8HiCAvQD7ZBDcHiCAvQD7ZBDItNDMHiCAvQiVXgD7ZBAolVvA+2UQPB4ggL0A+2QQHB4ggL0A+2AcHiCAvQD7ZBBolVCIlVuA+2UQfB4ggL0A+2QQXB4ggL0A+2QQTB4ggL0A+2QQqJVRSJVbQPtlELweIIC9APtkEJweIIC9APtkEIweIIC9APtkEOiVUMiVWwD7ZRD8HiCAvQD7ZBDcHiCAvQD7ZBDMHiCAvQiVX8iVWsi1UQD7ZKAw+2QgLB4QgLyA+2QgHB4QgLyA+2AsHhCAvIiU3ciU2oD7ZyBw+2QgYPtnoLD7ZKDsHmCAvwwecID7ZCBcHmCAvwx0WY" & _
                                                    "CgAAAA+2QgTB5ggL8A+2QgoL+Il1pA+2QgnB5wgL+A+2QgjB5wgL+A+2Qg/B4AgLwYl9oA+2Sg3B4AgLwQ+2SgyLVdzB4AgLwYtN7IlFnOsDi10QA9mLTQgz04ldEMHCEAPKiU0IM03swcEMA9kz04ldEItdCMHCCAPaiVXci1X0A1XoM/KJXQgz2cHGEItNFAPOwcMHiU0UM03owcEMA9Ez8olV9ItVFMHGCAPWiXXsi3XwA3XkM/6JVRQz0cHHEItNDAPPwcIHiU0MM03kwcEMA/Ez/ol18It1DMHHCAP3iX2Ui334A33gM8eJdQwz8cHAEItN/APIwcYHiU38M03gwcEMA/kzx4l9+It9/MHACAP4iX38M/mLTRADysHHBzPBiU0Qi00MwcAQA8iJTQwzyotVEMHBDAPRM8KJVRCLVQzBwAgD0IlVDDPRi030A87BwgeJTfSJVeiLVdwz0YtN/MHCEAPKiU38M86LdfTBwQwD8TPWiXX0i3X8wcIIA/KJdfwz8YtN8APPwcYHiU3wiXXki3XsM/GLTQjBxhADzolNCDPPi33wwcEMA/kz94l98It9CMHGCAP+iX0IM/mLTfgDy8HHB4l94It9lDP5iU34i00UwccQA8+JTRQzy4td+MHBDAPZM/uJXfjBxwgBfRSLXRQz2YvLiV3swcEHg22YAYtd+IlN7A+FQP7//wFFnAFdzItN2ANNEAFVqItVGIlN2Itd2IvDi03UA030iBqJTdSLTdADTfDB6AiIQgGLw4lN0ItN7AFNyItNxANN6MHoEIhCAsHrGIhaA4td1IvDiFoEwegIiEIFi8OJTcSLTcADTeTB6BCIQgaJTcCLTbwDTeABdaQBfaDB6xiIWgeL"
Private Const STR_THUNK2                As String = "XdCLw4haCIlNvItNuANNCMHoCIhCCYvDiU24i020A00UwegQiEIKwesYiFoLi13Mi8OJTbSLTbADTQyIWgzB6AiIQg2Lw4lNsItNrANN/MHoEIhCDsHrGIhaD4tdyIvDiU2siFoQwegIiEIRi8PB6BCIQhLB6xiIWhOLXcSLw4haFMHoCIhCFYvDwegQiEIWwesYiFoXi13Ai8OIWhjB6AiIQhmLw8HoEIhCGsHrGIhaG4tdvIvDiFocwegIiEIdi8PB6BCIQh7B6xiIWh+LXbiLw4haIMHoCIhCIYvDwegQiEIiwesYiFoji120i8OIWiTB6AiIQiWLw8HoEIhCJsHrGIhaJ4tdsIvDiFoowegIiEIpi8PB6BCIQirB6xiIWiuL2YhaLIvDwegIiEIti8PB6BCIQi7B6xiIWi+LXaiLw4haMMHoCIhCMY1KPIvDwesYwegQiEIyiFozi12ki8OIWjTB6AiIQjWLw8HoEIhCNsHrGIhaN4tdoIvDiFo4wegIiEI5i8PB6BCIQjrB6xiIWjuLVZyLwsHoCIgRiEEBi8JfwegQweoYXohBAohRA1uL5V3CFADMVYvsVv91EIt1CP91DFbo7VkAAGoQ/3UUjUYgUOj/NwAAi0UYg8QMx0Z0AAAAAIlGeF5dwhQAzMzMzMzMzMzMzFWL7FaLdQhX/3UM/3YwjX4gV41GEFBW6ET5//+LVngzwIAHAXULQDvCdAaABDgBdPVfXl3CCADMzMzMzMzMzMxVi+yD7BCNRfBqEP91IFDojDcAAIPEDI1F8FBqAP91JP91HP91GP91FP91EP91DP91COjZUgAAi+VdwiAAzMzMVYvs/3UkagH/dSD/dRz/dRj/dRT/dRD/dQz/dQjorlIAAF3CIADMzMzM" & _
                                                    "zMzMzMzMVYvs6LjZ//+50Hm5AIHpAEC5AAPBi00IUVD/dRSLAf91EP91DP8wjUEoUI1BGFDorPf//13CEADMzMzMzMzMzFWL7ItNCItFDIlBLItFEIlBMF3CDADMzMzMzMzMzMzMVYvsVot1CGo0agBW6O82AACLTQzHRiwAAAAAiwGJRjCLRRCJRgSNRgiJDsdGKAAAAAD/Mf91FFDokzYAAIPEGF5dwhAAzMzMzMzMzMzMzMxVi+yB7CAEAABTVldqcI2FcP3//8eFYP3//0HbAABqAFDHhWT9//8AAAAAx4Vo/f//AQAAAMeFbP3//wAAAADobDYAAIt1DI2FYP///2ofVlDoKjYAAIpGH4PEGIClYP////gkPwxAiIV/////jYXg+////3UQUOhEYAAAD1fAjbVg/v//Zg8ThWD+//+NvWj+//+5HgAAAGYPE0WA86W5HgAAAGYPE4Xg/v//jXWAx4Vg/v//AQAAAI19iMeFZP7//wAAAADzpbkeAAAAx0WAAQAAAI214P7//8dFhAAAAACNvej+//+7/gAAAPOluSAAAACNteD7//+NveD9///zpYvDD7bLwfgDg+EHD7a0BWD///+NheD9///T7oPmAVZQjUWAUOgWVgAAVo2FYP7//1CNheD+//9Q6AJWAACNheD+//9QjUWAUI2F4Pz//1Do6+T//42F4P7//1CNRYBQUOj6XQAAjYVg/v//UI2F4P3//1CNheD+//9Q6MDk//+NhWD+//9QjYXg/f//UFDozF0AAI2F4Pz//1CNhWD+//9Q6JldAACNRYBQjYVg/P//UOiJXQAAjUWAUI2F4P7//1CNRYBQ6JVHAACNheD8//9QjYXg/f//UI2F4P7/" & _
                                                    "/1Doe0cAAI2F4P7//1CNRYBQjYXg/P//UOhE5P//jYXg/v//UI1FgFBQ6FNdAACNRYBQjYXg/f//UOgjXQAAjYVg/P//UI2FYP7//1CNheD+//9Q6CldAACNhWD9//9QjYXg/v//UI1FgFDoEkcAAI2FYP7//1CNRYBQUOjh4///jUWAUI2F4P7//1BQ6PBGAACNhWD8//9QjYVg/v//UI1FgFDo2UYAAI2F4Pv//1CNheD9//9QjYVg/v//UOi/RgAAjYXg/P//UI2F4P3//1DojFwAAFaNheD9//9QjUWAUOh7VAAAVo2FYP7//1CNheD+//9Q6GdUAACD6wEPiR/+//+NheD+//9QUOiRMQAAjYXg/v//UI1FgFBQ6GBGAACNRYBQ/3UI6KRJAABfXluL5V3CDADMzMzMzMzMzMzMzFWL7IPsII1F4MZF4AlQ/3UMD1fAx0X5AAAAAP91CA8RReFmx0X9AABmD9ZF8cZF/wDoqvz//4vlXcIIAMzMzMxVi+yB7BQBAABTi10IjUXwVleLfQwPV8BQUItDBFfGRfAAZg/WRfHHRfkAAAAAZsdF/QAAxkX/AP/Qi3Ukg/4MdSBW/3UgjUXQUOjRMgAAg8QMZsdF3QAAxkXcAMZF3wHrMI1F8FCNhez+//9Q6O4oAABW/3UgjYXs/v//UOgOJwAAjUXQUI2F7P7//1DovicAAI1F8FCNhTz///9Q6L4oAAD/dRyNhTz/////dRhQ6LwmAACNRdDGReAAUFdTjUWMx0XpAAAAAA9XwGbHRe0AAFBmD9ZF4cZF7wDocPv//2oEagyNRYxQ6EP7//9qEI1F4FBQjUWMUOjz+v///3UUjYU8/////3UQUOiBJgAAjUXA" & _
                                                    "UI2FPP///1DoMScAAIt1LI1F4FZQjUXAUFDoP38AADLSjUXAuwEAAACF9nQai30oi8gr+YoMB41AATJI/wrRK/N18YTSdRT/dRSNRYz/dTD/dRBQ6IX6//8z2w9XwA8RRfCKRfAPEUXQikXQDxFF4IpF4A8RRcCKRcBqUI2FPP///2oAUOi0MQAAio08////jUWMajRqAFDooTEAAIpNjIPEGIvDX15bi+VdwiwAVYvsgewUAQAAU4tdCI1F8FZXi30MD1fAUFCLQwRXxkXwAGYP1kXxx0X5AAAAAGbHRf0AAMZF/wD/0It1JIP+DHUgVv91II1F0FDoETEAAIPEDGbHRd0AAMZF3ADGRd8B6zCNRfBQjYXs/v//UOguJwAAVv91II2F7P7//1DoTiUAAI1F0FCNhez+//9Q6P4lAACNRfBQjYU8////UOj+JgAA/3UcjYU8/////3UYUOj8JAAAjUXQxkXgAFBXU41FjMdF6QAAAAAPV8Bmx0XtAABQZg/WReHGRe8A6LD5//9qBGoMjUWMUOiD+f//ahCNReBQUI1FjFDoM/n//4t9FI1FjIt1KFdW/3UQUOgf+f//V1aNhTz///9Q6LEkAACNRcDGRcAAUI2FPP///8dFyQAAAAAPV8Bmx0XNAABQZg/WRcHGRc8A6EQlAAD/dTCNReBQjUXAUP91LOhRfQAAD1fADxFF8IpF8A8RRdCKRdAPEUXgikXgDxFFwIpFwGpQjYU8////agBQ6AIwAACKhTz///9qNI1FjGoAUOjvLwAAikWMg8QYX15bi+VdwiwAVYvsi1UMi00QVot1CIsGMwKJAYtGBDNCBIlBBItGCDNCCIlBCItGDDNCDIlBDF5dwgwAzMzM" & _
                                                    "zMzMzMzMzMzMzFWL7FFTi10MVleLfQhmx0X8AOGLD4vB0eiD4QGJA4tXBIvC0eiD4gHB4R8LyMHiH4lLBIt3CIvG0eiD5gEL0MHmH4lTCItPDIvB0eiD4QEL8F+JcwwPtkQN/MHgGDEDXluL5V3CCADMzMzMzMzMzMxVi+yLVQxWi3UID7YOD7ZGAcHhCAvID7ZGAsHhCAvID7ZGA8HhCAvIiQoPtk4ED7ZGBcHhCAvID7ZGBsHhCAvID7ZGB8HhCAvIiUoED7ZOCA+2RgnB4QgLyA+2RgrB4QgLyA+2RgvB4QgLyIlKCA+2TgwPtkYNweEIC8gPtkYOweEIC8gPtkYPweEIC8iJSgxeXcIIAMzMzMzMzMzMzMzMVYvsg+wgVldqEI1F4GoAUOh7LgAAahD/dQyNRfBQ6D0uAACLfQiDxBgPEE3gM/aQi8a5HwAAAIPgHyvIi8bB+AWLBIfT6KgBdAwPEEXwZg/vyA8RTeCNRfBQUOiQ/v//RoH+gAAAAHzHahCNReBQ/3UQ6OktAACDxAxfXovlXcIMAMzMzMzMzMzMzMzMzMzMVYvsVot1DFeLfQiLF4vCwegYiAaLwsHoEIhGAYvCwegIiEYCiFYDi08Ei8HB6BiIRgSLwcHoEIhGBYvBwegIiEYGiE4Hi08Ii8HB6BiIRgiLwcHoEIhGCYvBwegIiEYKiE4Li08Mi8HB6BiIRgyLwcHoEIhGDYvBwegIiEYOX4hOD15dwggAzMzMzMzMzMzMVYvsg+xEVot1CIO+qAAAAAB0BlboR0YAADPJDx9EAAAPtoQOiAAAAIlEjbxBg/kQcu5Wx0X8AAAAAOghRQAAjUW8UFbot0QAAItVDDPJZpCKBI6IBBFBg/kQ" & _
                                                    "cvRorAAAAGoAVugHLQAAigaDxAxei+VdwggAzMzMzMzMzMzMzMxVi+xWi3UIaKwAAABqAFbo3CwAAItNDGoQ/3UQD7YBiUZED7ZBAYlGSA+2QQKJRkwPtkEDg+APiUZQD7ZBBCX8AAAAiUZUD7ZBBYlGWA+2QQaJRlwPtkEHg+APiUZgD7ZBCCX8AAAAiUZkD7ZBCYlGaA+2QQqJRmwPtkELg+APiUZwD7ZBDCX8AAAAiUZ0D7ZBDYlGeA+2QQ6JRnwPtkEPg+APx4aEAAAAAAAAAImGgAAAAI2GiAAAAFDoASwAAIPEGF5dwgwAzMzMzMzMzMzMVYvs6HjO//+5MLm5AIHpAEC5AAPBi00IUVD/dRCNgagAAAD/dQxqEFCNgZgAAABQ6Gvr//9dwgwAzMzMzMzMzFWL7IPsGFNWV+gyzv///3UIvmDAuQC5QAAAAIHuAEC5AAPwi0UIVo14ZItAYPfhAweL2IPSAIPACIPgPyvIUWoAagBogAAAAGpAV4t9CA+k2gOJVfyNRyDB4wNQiVX46Azq//+LVfyLy4vCiF3vwegYiEXoi8LB6BCIRemLwsHoCIhF6opF+IhF64vCD6zBGGoIwegYiE3si8KLyw+swRDB6BCLw4hN7Q+s0AiIRe6NRehQweoIV+hkAQAAixeLwot1DMHoGIgGi8LB6BCIRgGLwsHoCIhGAohWA4tPBIvBwegYiEYEi8HB6BCIRgWLwcHoCIhGBohOB4tPCIvBwegYiEYIi8HB6BCIRgmLwcHoCIhGCohOC4tPDIvBwegYiEYMi8HB6BCIRg2LwcHoCIhGDohOD4tPEIvBwegYiEYQi8HB6BCIRhGLwcHoCIhGEohOE4tPFIvBwegYiEYU" & _
                                                    "i8HB6BCIRhWLwcHoCIhGFohOF4tPGIvBwegYiEYYi8HB6BCIRhmLwcHoCIhGGohOG4tPHIvBwegYiEYci8HB6BCIRh2LwWpowegIagCIRh5XiE4f6CkqAACDxAxfXluL5V3CCADMzMzMzMzMzMzMzMzMVYvsVot1CGpoagBW6P8pAACDxAzHBmfmCWrHRgSFrme7x0YIcvNuPMdGDDr1T6XHRhB/Ug5Rx0YUjGgFm8dGGKvZgx/HRhwZzeBbXl3CBABVi+zoGMz//7lgwLkAgekAQLkAA8GLTQhRUP91EI1BZP91DGpAUI1BIFDoEen//13CDADMzMzMzMzMzMzMzMzMVYvsg+xAjUXAUP91COi+AAAAajCNRcBQ/3UM6DApAACDxAyL5V3CCADMzMzMzMzMVYvsVot1CGjIAAAAagBW6DwpAACDxAzHBtieBcHHRgRdnbvLx0YIB9V8NsdGDCopmmLHRhAX3XAwx0YUWgFZkcdGGDlZDvfHRhzY7C8Vx0YgMQvA/8dGJGcmM2fHRigRFVhox0Ysh0q0jsdGMKeP+WTHRjQNLgzbx0Y4pE/6vsdGPB1ItUdeXcIEAMzMzMzM6RsEAADMzMzMzMzMzMzMzFWL7IPsHItFCFONmMQAAABWi4DAAAAAV7+AAAAA9+eL8AMzi8aD0gAPpMIDweADiVX8iUX4iVX06NPK////dQi5MMK5AIHpAEC5AAPBUI1GEIt1CIPgfyv4V2oAagBogAAAAGiAAAAAU41GQFDozub//2oIjUXkx0XkAAAAAFBWx0XoAAAAAOiEAwAAi138i8OLVfiLysHoGIhF5IvDwegQiEXli8PB6AiIReaKRfSIReeLww+swRhqCMHoGIhN6IvD" & _
                                                    "i8qIVesPrMEQwegQi8KITekPrNgIiEXqjUXkUFbB6wjoKQMAAIteBIvDiw6JTfzB6BiLfQyIB4vDwegQiEcBi8PB6AiIRwKLww+swRiIXwPB6BiITwSLw4tN/A+swRDB6BCITwWLTfyLwQ+s2AiIRwaLxohPB8HrCItYCIvLi1AMi8LB6BiIRwiLwsHoEIhHCYvCwegIiEcKi8IPrMEYiFcLwegYiE8Mi8KLyw+swRDB6BCITw2Lww+s0AiIRw6LxohfD8HqCItYEIvLi1AUi8LB6BiIRxCLwsHoEIhHEYvCwegIiEcSi8IPrMEYiFcTwegYiE8Ui8KLyw+swRDB6BCLw4hPFQ+s0AiIRxaLxsHqCIhfF4tYGIvLi1Aci8LB6BiIRxiLwsHoEIhHGYvCwegIiEcai8IPrMEYiFcbwegYiE8ci8KLyw+swRDB6BCITx2Lww+s0AiIRx6LxohfH8HqCItYIIvLi1Aki8LB6BiIRyCLwsHoEIhHIYvCwegIiEcii8IPrMEYiFcjwegYiE8ki8KLyw+swRDB6BCITyWLww+s0AiIRyaLxohfJ8HqCItYKIvLi1Asi8LB6BiIRyiLwsHoEIhHKYvCwegIiEcqi8IPrMEYiFcrwegYiE8si8KLyw+swRDB6BCLw4hPLQ+s0AjB6giIRy6LxohfL413OGjIAAAAagCLWDCLy4tQNIvCwegYiEcwi8LB6BCIRzGLwsHoCIhHMovCD6zBGIhXM8HoGIhPNIvCi8sPrMEQwegQiE81i8MPrNAIiEc2iF83i30IweoIV4tXPIvCi184i8vB6BiIBovCwegQiEYBi8LB6AiIRgKLwg+swRiIVgPB6BiITgSLwovLD6zBEMHoEIvD" & _
                                                    "iE4FD6zQCIhGBsHqCIheB+hFJQAAg8QMX15bi+VdwggAzMzMzMzMzMzMVYvsVot1CGjIAAAAagBW6BwlAACDxAzHBgjJvPPHRgRn5glqx0YIO6fKhMdGDIWuZ7vHRhAr+JT+x0YUcvNuPMdGGPE2HV/HRhw69U+lx0Yg0YLmrcdGJH9SDlHHRigfbD4rx0YsjGgFm8dGMGu9QfvHRjSr2YMfx0Y4eSF+E8dGPBnN4FteXcIEAMzMzMzMVYvs6PjG//+5MMK5AIHpAEC5AAPBi00IUVD/dRCNgcQAAAD/dQxogAAAAFCNQUBQ6Ovj//9dwgwAzMzMzMzMzFWL7FaLdQj/dQyLDo1GCFD/dgSLQQT/0ItWLItGMAPWSF6ARAIIAXUTDx+AAAAAAIXAdAhIgEQCCAF09F3CCABVi+xTi10MVleLfQgPtkMomYvIi/IPpM4ID7ZDKcHhCJkLyAvyD6TOCA+2QyrB4QiZC8gL8g+kzggPtkMrweEImQvIC/IPtkMsD6TOCJnB4QgL8gvID7ZDLQ+kzgiZweEIC/ILyA+2Qy4PpM4ImcHhCAvyC8gPtkMvD6TOCJnB4QgL8gvIiXcEiQ8PtkMgmYvIi/IPtkMhD6TOCJnB4QgL8gvID7ZDIg+kzgiZweEIC/ILyA+2QyMPpM4ImcHhCAvyC8gPtkMkD6TOCJnB4QgLyAvyD6TOCA+2QyXB4QiZC8gL8g+kzggPtkMmweEImQvIC/IPpM4ID7ZDJ8HhCJkLyAvyiU8IiXcMD7ZDGJmLyIvyD6TOCA+2QxnB4QiZC8gL8g+2QxoPpM4ImcHhCAvyC8gPtkMbD6TOCJnB4QgL8gvID7ZDHA+kzgiZweEIC/ILyA+2Qx0PpM4I" & _
                                                    "mcHhCAvyC8gPtkMeD6TOCJnB4QgL8gvID7ZDHw+kzgiZweEIC/ILyIl3FIlPEA+2QxCZi8iL8g+2QxEPpM4ImcHhCAvyC8gPtkMSD6TOCMHhCJkLyAvyD6TOCA+2QxPB4QiZC8gL8g+kzggPtkMUweEImQvIC/IPpM4ID7ZDFcHhCJkLyAvyD6TOCA+2QxbB4QiZC8gL8g+2QxcPpM4ImcHhCAvyC8iJdxyJTxgPtkMImYvIi/IPtkMJD6TOCJnB4QgL8gvID7ZDCg+kzgiZweEIC/ILyA+2QwsPpM4ImcHhCAvyC8gPtkMMD6TOCJnB4QgL8gvID7ZDDQ+kzgiZweEIC/ILyA+2Qw4PpM4ImcHhCAvyC8gPtkMPD6TOCJnB4QgLyAvyiU8giXckD7YDmYvIi/IPtkMBD6TOCJnB4QgL8gvID7ZDAg+kzgiZweEIC/ILyA+2QwMPpM4ImcHhCAvyC8gPtkMED6TOCJnB4QgL8gvID7ZDBQ+kzgiZweEIC/ILyA+2QwYPpM4ImcHhCAvyC8gPtkMHD6TOCJnB4QgLyAvyiXcsiU8oX15bXcIIAMzMzMzMVYvsU4tdDFZXi30ID7ZDGJmLyIvyD6TOCA+2QxnB4QiZC8gL8g+kzggPtkMaweEImQvIC/IPpM4ID7ZDG8HhCJkLyAvyD7ZDHA+kzgiZweEIC/ILyA+2Qx0PpM4ImcHhCAvyC8gPtkMeD6TOCJnB4QgL8gvID7ZDHw+kzgiZweEIC/ILyIl3BIkPD7ZDEJmLyIvyD7ZDEQ+kzgiZweEIC/ILyA+2QxIPpM4ImcHhCAvyC8gPtkMTD6TOCJnB4QgL8gvID7ZDFA+kzgiZweEIC8gL8g+kzggPtkMVweEI" & _
                                                    "mQvIC/IPpM4ID7ZDFsHhCJkLyAvyD6TOCA+2QxfB4QiZC8gL8olPCIl3DA+2QwiZi8iL8g+kzggPtkMJweEImQvIC/IPtkMKD6TOCJnB4QgL8gvID7ZDCw+kzgiZweEIC/ILyA+2QwwPpM4ImcHhCAvyC8gPtkMND6TOCJnB4QgL8gvID7ZDDg+kzgiZweEIC/ILyA+2Qw8PpM4ImcHhCAvyC8iJdxSJTxAPtgOZi8iL8g+2QwEPpM4ImcHhCAvyC8gPtkMCD6TOCMHhCJkLyAvyD7ZDAw+kzgiZweEIC/ILyA+2QwQPpM4ImcHhCAvyC8gPtkMFD6TOCJnB4QgL8gvID7ZDBg+kzgiZweEIC/ILyA+2QwcPpM4ImcHhCAvIC/KJdxyJTxhfXltdwggAzMzMVYvsgeyQAAAAjUXQ/3UMUOjL+v//jUXQUOjSTgAAhcB0CDPAi+VdwggAjUXQUOgNwf//BXABAABQ6NJNAACD+AF0Fej4wP//BXABAABQjUXQUFDoGGgAAGoAjUXQUOjdwP//BRABAABQjYVw////UOhbxP//jYVw////UOjvw///hcB1nYpFoItNCCQBBAKIAY2FcP///1CNQQFQ6K8AAAC4AQAAAIvlXcIIAMzMzMxVi+yD7GCNReD/dQxQ6C79//+NReBQ6FVOAACFwHQIM8CL5V3CCACNReBQ6GDA//8FkAAAAFDolU0AAIP4AXQV6EvA//8FkAAAAFCNReBQUOh7aQAAagCNReBQ6DDA//+DwFBQjUWgUOhjxf//jUWgUOh6w///hcB1pYpFwItNCCQBBAKIAY1FoFCNQQFQ6H0CAAC4AQAAAIvlXcIIAMzMVYvsVot1CLEoV4t9DA+2RweI" & _
                                                    "RigPtkcGiEYpiweLVwTo+8z//4hGKrEgiweLVwTo7Mz//4hGK4sPi0cED6zBGIhOLIsPwegYi0cED6zBEIhOLYsPwegQi0cED6zBCIhOLrEowegID7YHiEYvD7ZHD4hGIA+2Rw6IRiGLRwiLVwzom8z//4hGIrEgi0cIi1cM6IvM//+IRiOLTwiLRwwPrMEYiE4ki08IwegYi0cMD6zBEIhOJYtPCMHoEItHDA+swQiITiaxKMHoCA+2RwiIRicPtkcXiEYYD7ZHFohGGYtHEItXFOg2zP//iEYasSCLRxCLVxToJsz//4hGG4tPEItHFA+swRiIThyLTxDB6BiLRxQPrMEQiE4di08QwegQi0cUD6zBCIhOHrEowegID7ZHEIhGHw+2Rx+IRhAPtkceiEYRi0cYi1cc6NHL//+IRhKxIItHGItXHOjBy///iEYTi08Yi0ccD6zBGIhOFItPGMHoGItHHA+swRCIThWLTxjB6BCLRxwPrMEIiE4WsSjB6AgPtkcYiEYXD7ZHJ4hGCA+2RyaIRgmLRyCLVyTobMv//4hGCrEgi0cgi1ck6FzL//+IRguLTyCLRyQPrMEYwegYiE4Mi08gi0ckD6zBEMHoEIhODYtPIItHJA+swQjB6AiITg4PtkcgiEYPD7ZHL4gGD7ZHLohGAbEoi0coi1cs6AjL//+IRgKxIItHKItXLOj4yv//iEYDi08oi0csD6zBGMHoGIhOBItPKItHLA+swRDB6BCITgWLTyiLRywPrMEIwegIiE4GD7ZHKF+IRgdeXcIIAMzMzMzMzMzMVYvsVot1CLEoV4t9DA+2RweIRhgPtkcGiEYZiweLVwToi8r//4hGGrEgiweLVwTofMr//4hG" & _
                                                    "G4sPi0cED6zBGIhOHIsPwegYi0cED6zBEIhOHYsPwegQi0cED6zBCIhOHrEowegID7YHiEYfD7ZHD4hGEA+2Rw6IRhGLRwiLVwzoK8r//4hGErEgi0cIi1cM6BvK//+IRhOLTwiLRwwPrMEYiE4Ui08IwegYi0cMD6zBEIhOFYtPCMHoEItHDA+swQiIThaxKMHoCA+2RwiIRhcPtkcXiEYID7ZHFohGCYtHEItXFOjGyf//iEYKsSCLRxCLVxTotsn//4hGC4tPEItHFA+swRiITgyLTxDB6BiLRxQPrMEQiE4Ni08QwegQi0cUD6zBCIhODrEowegID7ZHEIhGDw+2Rx+IBg+2Rx6IRgGLRxiLVxzoYsn//4hGArEgi0cYi1cc6FLJ//+IRgOLTxiLRxwPrMEYwegYiE4Ei08Yi0ccD6zBEMHoEIhOBYtPGItHHA+swQjB6AiITgYPtkcYX4hGB15dwggAzMxVi+yD7DBTi10ID1fAVot1DMdF0AMAAADHRdQAAAAADxFF2I1GAWYP1kX4UFMPEUXo6Er1//+APgR1FY1GMVCNQzBQ6Dj1//9eW4vlXcIIAFdTjXswV+iVWQAA6IC7//8FsAAAAFCNRdBQV1fo31kAAFNXV+gXWQAA6GK7//8FsAAAAFDoV7v//wXgAAAAUFdX6PpQAABX6KQZAACKBjP2iw8kAQ+2wIPhAZk7yHUEO/J0ElfoJ7v//wWwAAAAUFfoS2IAAF9eW4vlXcIIAMzMVYvsg+wgU4tdCA9XwFaLdQzHReADAAAAx0XkAAAAAA8RReiNRgFmD9ZF+FBT6I73//+APgR1FY1GIVCNQyBQ6Hz3//9eW4vlXcIIAFdTjXsgV+j5WAAA6LS6" & _
                                                    "//+DwBBQjUXgUFdX6EVZAABTV1fofVgAAOiYuv//g8AQUOiPuv//g8AwUFdX6HRQAABX6I4ZAACKBjP2iw8kAQ+2wIPhAZk7yHUEO/J0EFfoYbr//4PAEFBX6JdjAABfXluL5V3CCADMzMzMzMzMzMzMzMzMzFWL7IHs8AAAAI2FEP////91CFDoSP7///91DI1F0FDovPP//2oAjUXQUI2FEP///1CNhXD///9Q6JO9//+NhXD///9Q/3UQ6AT6//+NhXD///9Q6Bi9///32BvAQIvlXcIMAMzMzMzMzMzMzMzMzMxVi+yB7KAAAACNhWD/////dQhQ6Kj+////dQyNReBQ6Fz2//9qAI1F4FCNhWD///9QjUWgUOjWvv//jUWgUP91EOgK/P//jUWgUOjhvP//99gbwECL5V3CDADMzMzMzMxVi+yD7GCNRaBW/3UIUOh9/f//i3UMjUWgUI1GAcYGBFDoWvn//41F0FCNRjFQ6E35//+4AQAAAF6L5V3CCADMVYvsg+xAjUXAVv91CFDoDf7//4t1DI1FwFCNRgHGBgRQ6Ir7//+NReBQjUYhUOh9+///uAEAAABei+VdwggAzFWL7IHswAAAAFeLfRBX6I1GAACFwHQJM8Bfi+VdwhAAV+jKuP//BXABAABQ6I9FAACD+AF0Eui1uP//BXABAABQV1fo2F8AAGoAV+iguP//BRABAABQjYVA////UOgevP//jYVA////UOiCuP//BXABAABQ6EdFAACD+AF0GOhtuP//BXABAABQjYVA////UFDoil8AAI2FQP///1Do/kUAAIXAD4Vt////Vot1FI2FQP///1BW6EX4////dQiNRaBQ6Mnx///oJLj//wVw"
Private Const STR_THUNK3                As String = "AQAAUI1FoFCNhUD///9QjUXQUOjaUgAA/3UMjUWgUOie8f//6Pm3//8FcAEAAFCNRdBQjUWgUI1F0FDokk0AAOjdt///BXABAABQV1foAE4AAOjLt///BXABAABQV41F0FBQ6IpSAACNRdBQjUYwUOi99///XrgBAAAAX4vlXcIQAFWL7IHsgAAAAFeLfRBX6G1FAACFwHQJM8Bfi+VdwhAAV+h6t///BZAAAABQ6K9EAACD+AF0Euhlt///BZAAAABQV1fomGAAAGoAV+hQt///g8BQUI1FgFDog7z//41FgFDoOrf//wWQAAAAUOhvRAAAg/gBdBXoJbf//wWQAAAAUI1FgFBQ6FVgAACNRYBQ6OxEAACFwA+Fe////1aLdRSNRYBQVuh2+f///3UIjUXAUOia8///6OW2//8FkAAAAFCNRcBQjUWAUI1F4FDoHlMAAP91DI1FwFDocvP//+i9tv//BZAAAABQjUXgUI1FwFCNReBQ6JZMAADoobb//wWQAAAAUFdX6CRPAADoj7b//wWQAAAAUFeNReBQUOjOUgAAjUXgUI1GIFDo8fj//164AQAAAF+L5V3CEADMzMzMVYvsgeyAAgAAjYWA/f//Vv91CFDoZ/r//4t1EI2F0P7//1ZQ6Nfv//+NRjBQjYVg////UOjH7///jYXQ/v//UOjLQwAAhcAPhZwDAACNhWD///9Q6LdDAACFwA+FiAMAAI2F0P7//1Do87X//wVwAQAAUOi4QgAAg/gBD4VoAwAAjYVg////UOjTtf//BXABAABQ6JhCAACD+AEPhUgDAABTV+i4tf//BXABAABQjYVg////UI1FwFDo0ksAAP91DI2FAP///1DoM+///+iOtf//BXABAABQjUXAUI2FAP//" & _
                                                    "/1BQ6EdQAADocrX//wVwAQAAUI1FwFCNhdD+//9QjYWg/v//UOglUAAAjYWA/f//UI2FEP7//1DoklgAAI2FsP3//1CNhUD+//9Q6H9YAADoKrX//wUQAQAAUI2FMP///1DoaFgAAOgTtf//BUABAABQjYVg////UOhRWAAA6Py0//8FsAAAAFCNhTD///9QjYUQ/v//UI1FwFDoT1MAAI2FQP7//1CNhRD+//9QjYVg////UI2FMP///1Donrv//+i5tP//BbAAAABQjUXAUFDo2UoAAI1FwFCNhUD+//9QjYUQ/v//UOgyxP//x0XwAAAAAOiGtP//BRABAACJRfSNhYD9//+JRfiNhRD+//+JRfyNhaD+//9Q6PBVAACL2I2FAP///1Do4lUAADvDD0fYjYUA////jXP/VlDo3V4AAAvCdAe/AQAAAOsCM/9WjYWg/v//UOjDXgAAC8J0B74CAAAA6wIz9gv3jUWQi3S18FZQ6FZXAACNRjBQjYVw/v//UOhGVwAAjUXAUOgNQAAAjXP+x0XAAQAAAMdFxAAAAACF9g+I6AAAAA8fQACNRcBQjYVw/v//UI1FkFDo/LP//1aNhQD///9Q6E9eAAALwnQHvwEAAADrAjP/Vo2FoP7//1DoNV4AAAvCdAe4AgAAAOsCM8ALx4t8hfCF/w+EhQAAAFeNhTD///9Q6L1WAACNRzBQjYVg////UOitVgAAjUXAUI2FYP///1CNhTD///9Q6ObC///oQbP//wWwAAAAUI2FMP///1CNRZBQjYXg/f//UOiUUQAAjYVw/v//UI1FkFCNhWD///9QjYUw////UOjmuf//jYXg/f//UI1FwFBQ6KVQAACD7gEPiRz////o" & _
                                                    "57L//wWwAAAAUI1FwFBQ6AdJAACNRcBQjYVw/v//UI1FkFDoY8L//41FkFDourL//wVwAQAAUOh/PwAAX1uD+AF0Feijsv//BXABAABQjUWQUFDow1kAAI2F0P7//1CNRZBQ6FM/AAD32F4bwECL5V3CDAAzwF6L5V3CDADMzMzMzMzMzMzMzMzMzFWL7IHssAEAAI2FUP7//1b/dQhQ6Df3//+LdRCNhTD///9WUOjn7v//jUYgUI1FkFDo2u7//42FMP///1Do/j8AAIXAD4VuAwAAjUWQUOjtPwAAhcAPhV0DAACNhTD///9Q6Pmx//8FkAAAAFDoLj8AAIP4AQ+FPQMAAI1FkFDo3LH//wWQAAAAUOgRPwAAg/gBD4UgAwAAU1fowbH//wWQAAAAUI1FkFCNReBQ6D5KAAD/dQyNhVD///9Q6E/u///omrH//wWQAAAAUI1F4FCNhVD///9QUOjTTQAA6H6x//8FkAAAAFCNReBQjYUw////UI2FEP///1DosU0AAI2FUP7//1CNhbD+//9Q6P5UAACNhXD+//9QjYXQ/v//UOjrVAAA6Dax//+DwFBQjYVw////UOjWVAAA6CGx//+DwHBQjUWQUOjEVAAA6A+x//+DwBBQjYVw////UI2FsP7//1CNReBQ6JRPAACNhdD+//9QjYWw/v//UI1FkFCNhXD///9Q6La4///o0bD//4PAEFCNReBQUOhTSQAAjUXgUI2F0P7//1CNhbD+//9Q6JzA///HRdAAAAAA6KCw//+DwFCJRdSNhVD+//+JRdiNhbD+//+JRdyNhRD///9Q6FxSAACL2I2FUP///1DoTlIAADvDD0fYjYVQ////jXP/VlDo+VoAAAvC" & _
                                                    "dAe/AQAAAOsCM/9WjYUQ////UOjfWgAAC8J0B74CAAAA6wIz9gv3jUWwi3S10FZQ6NJTAACNRiBQjYXw/v//UOjCUwAAjUXgUOiJPAAAjXP+x0XgAQAAAMdF5AAAAACF9g+I1QAAAI1F4FCNhfD+//9QjUWwUOicsf//Vo2FUP///1Dob1oAAAvCdAe/AQAAAOsCM/9WjYUQ////UOhVWgAAC8J0B7gCAAAA6wIzwAvHi3yF0IX/dHpXjYVw////UOhBUwAAjUcgUI1FkFDoNFMAAI1F4FCNRZBQjYVw////UOhgv///6Guv//+DwBBQjYVw////UI1FsFCNhZD+//9Q6PBNAACNhfD+//9QjUWwUI1FkFCNhXD///9Q6BW3//+NhZD+//9QjUXgUFDoBE0AAIPuAQ+JK////+gWr///g8AQUI1F4FBQ6JhHAACNReBQjYXw/v//UI1FsFDo5L7//41FsFDo667//wWQAAAAUOggPAAAX1uD+AF0FejUrv//BZAAAABQjUWwUFDoBFgAAI2FMP///1CNRbBQ6PQ7AAD32F4bwECL5V3CDAAzwF6L5V3CDADMzMzMzMzMzMzMzMzMzMxVi+yLTQiLwcHoB4Hhf39//yUBAQEBA8lrwBszwV3CBADMzMzMzMzMzMzMzMzMzMxVi+zoeK7//7nwkrkAgekAQLkAA8GLTQhRUP91EI1BMP91DGoQUI1BIFDoccv//13CDADMzMzMzMzMzMzMzMzMVYvsi00Ii0UQAUE4g1E8AIlFEIlNCF3ppP///8zMzMxVi+xWi3UIg35IAXUNVugtAAAAx0ZIAgAAAItFEAFGQFD/dQyDVkQAVuhy////Xl3CDADMzMzMzMzMzMzM" & _
                                                    "zMzMVYvsVot1CItOMIXJdCm4EAAAACvBUI1GIAPBagBQ6F0LAACDxAyNRiBQVugQAAAAx0YwAAAAAF5dwgQAzMzMzFWL7IPsEI1F8FZXUP91DOj82///i3UIjUXwjX4QV1dQ6Dvb//9XVlfog9z//19ei+VdwggAzMzMzMzMzMzMzMxVi+yD7BRTVot1CItGSIP4AXQFg/gCdQ1W6GL////HRkgAAAAAi144i1Y8D6TaA2oIi8LB4wPB6BiLy4hF7IvCwegQiEXti8LB6AiIRe4PtsKIRe+Lwg+swRiJVfzB6BiITfCLwovLiF3zD6zBEMHoEIvDiE3xD6zQCIhF8o1F7FDB6ghW6Fb+//+LXkCLVkQPpNoDagiLwsHjA8HoGIvLiEXsi8LB6BCIRe2LwsHoCIhF7g+2wohF74vCD6zBGIlV/MHoGIhN8IvCi8uIXfMPrMEQwegQi8OITfEPrNAIiEXyjUXsUMHqCFbo8f3///91DI1GEFDoBdz//15bi+VdwggAzMzMzMzMzMzMzMzMzFWL7FaLdQhqUGoAVujfCQAAg8QMVv91DOij2v//x0ZIAQAAAF5dwggAzMzMzMzMzFWL7FaLdRRXM/+D7gF4MotFDFOLXQgr2ClFEI0UsGaQiwwTjVL8M8ADzxPAA0oEg9AAg+4Bi/iLRRCJTBAEeeBbi8dfXl3CEADMzMzMzMzMVYvsVleLfRCLx5mD4h8D0MH6BYHnHwAAgHkFT4PP4EeLdQyLz4vG0+CF/3UEM/brCbkgAAAAK8/T7ot9CDPJAUSXBBPJhfZ1BIXJdCczwAPOE8ABTJcIg9AAg8IDhcB0E40Ul4vIjVIEM8ABSvwTwIXAdfBfXl3CDADMzMzMzMxV" & _
                                                    "i+yD7DiLRQyLVRRTVovwM9sr8leJdewPiP0BAACLRQiLzsHhBYvWiU30Dx9EAACLNJiF9nUMQ4PpIIlN9OnKAQAAVujXwP//i/iLz9Pmhf9+Go1DATtFDH0Si0UIuSAAAAArz4tUmATT6gvyi00gi8b3ZRyLRfSDweErx4vyA8iJdfyJTfh5HYP54A+OgwEAAPfZ0+4zyYl1/IlN+IX2D4RvAQAAi/mB5x8AAIB5BU+Dz+BHi8GJfdCZg+IfA8KLVQzB+AUr0ItFFCvQhf91cY1I/8dF6AEAAACNNBEPV8BmDxNFyDvzD4wBAQAAi0UIjTywi0XMiUXki0XIiUXwhcl5BDPA6waLRRCLBIj3ZfwDRfD30BNV5IlV8DPSAwfHReQAAAAAE9IDReiJB4PSAE5JiVXog+8EO/N9w+mtAAAAjXj/x0XcAAAAAI0EF4lF2DvDD4ybAAAAi03QuiAAAAAr0cdF6AEAAACJVcwPV8CLVQhmDxNF4I0EgolF8ItF5IlF1ItF4IlF5A8fRAAAhf95BDPA6waLRRCLBLj35ovwA3XkE1XUM8CJVeSL1tPiC1Xci03w99LHRdQAAAAAAxETwANV6IkRi03Mg9AAg23wBE+JReiLxot1/NPoi03QiUXci0XYSIlF2DvDfaGLTfiLdfyLfRiF/3QIUVZX6Hz9//+LVeyLTfSLRQg72g+OG/7//4tFDItVFItdEDPJhcB+Lov6K/iNFLsPH0QAAI0ED4XAeQQzwOsCiwKLdQiLNI478HJidwlBg8IEO00MfN6LTQzHRQwBAAAAjVH/hdJ4NYvyK/EDdRSNPLOLXQiF9nkEM8nrAosPM8D30QMMkxPAA00MiQyTg9AAToPvBIlFDIPq" & _
                                                    "AXnYi0UYhcB0CmoAagFQ6Nf8//9fXluL5V3CHADMzMzMzMzMzMzMzMzMzFWL7IPsHFOLXRRWV4P7Mg+OtAEAAIt9GIvDi3UMi8uZVyvC0fhQ/3UQK8iJRfhW/3UIiU386Mb///+LTfiLRRBX/3X8jRSOjQTIiVXkiUX0i0UI/3X0Uo0EiFCJRezonf///4tF/ECNDIeNUQTHAQAAAADHAgAAAACJTQyNTwSJVeiLVfiJTfDHAQAAAADHBwAAAACF0n47i134A8Arwo0Mh4tF/EArwo0Uh4t9CCv+iwQ3jVIEiUL8jXYEi0b8jUkEiUH8g+sBdeaLXRSLfRiLTfCLdfxWUf917FHoj/v//1aJB4tF6FD/deRQ6H/7//+LTQyJAY1OAYvBjRTPweAEA8eJVQhQUVL/dQxX6O7+//+LVfgzycdHDAAAAADHRwgAAAAAx0cEAAAAAI00EscHAAAAAIX2fh2LRfyLXRBAK8KNFMeLBIuNUgRBiUL8O8588otdFIt1/I0ENlCNRwhQ/3X0UOgG+///iUcEjUYBjTQAi0UIVlBXUOgxAgAAi038i30QVo1BASvYA9sr2Y0En1D/dQhQ6NT6//+L8IX2D4ScAAAAjVf8jRSaDx9AADPAjVL8AXIEE8CL8IX2dfBfXluL5V3CFACLVRCNDBuFyX4GM8CL+vOri0UIjQydAAAAAI0U2o0cATvYdlWLRQyNNAGNSvyJdRSLVQiJTRCQg+sEM/878HYnDx+AAAAAAIsDg+4E9yYDx4PSAAEBi0UMg9IAg+kEi/o78Hfji1UIi3UUiTmLTRCD6QSJTRA72ne+X15bi+VdwhQAzMzMzMxVi+yD7BBTVot1FFeD/jIPjrgAAACLTRiL" & _
                                                    "xot9DJkrwovQjRzx0fqLxivCiVX8U1CJRfiNNJUAAAAAjQTRiXX0UI0EPot1CFCNBJZQ6F/9//+LRfQDRRhT/3X8iUUIUItF+I0Eh1BW6JT///9T/3X8i134/3UYV40EnlDogP///4t9EIt1/IXbfh2LVRiNBHOLTfQDz40UgosCjVIEiQGNSQSD6wF18YtdGFZT/3UIU+hr+f//i0UUVleNBINQU+hc+f//X15bi+VdwhQAi10QhfZ+CIvOM8CL+/Ori0UIjRSziVUUjTywO/h2aotFDIPoBI00sItFCIl1DA8fRAAAg+8ED1fAZg8TRfCLyjvTdjiLRfCLXfSJRRhmDx9EAACLBo12/Pcng+kEAwGD0gADRRiJARPTM9uJVRg7TRB34ItVFItdEItFCIt1DIPqBIlVFDv4d6pfXluL5V3CFADMVYvsVot1FIPuAcdFFAEAAAB4NYtFCFOLXRBXi30MjRSwK/gr2IsMF41S/DPA99EDSgQTwANNFIlMEwSD0ACD7gGJRRR5319bXl3CEADMzMxVi+yB7IAAAAC5IAAAAFOLXQxWV4vzjX2A86W+/QAAAI1FgFBQ6JYqAACD/gJ0EIP+BHQLU41FgFBQ6KEUAACD7gF53It9CI11gLkgAAAA86VfXluL5V3CCADMzMzMzMxVi+xTVot1CFdW6HH1//+L2FPoafX//4vQUuhh9f//i/gz/ov3i8czw8HPCDPywcAIi87ByRAzwTPHM8ZfM8MzRQheW13CBADMzMzMzMzMzFWL7FaLdQj/Nuii/////3YEiQbomP////92CIlGBOiN/////3YMiUYI6IL///+JRgxeXcIEAMzMzMzMzMzMzMxVi+xTi10IVlcPtnsH" & _
                                                    "D7ZDAg+2cwsPtlMPwecIC/gPtksDD7ZDDcHnCAv4weYID7ZDCMHnCAv4weIID7ZDBgvwweEID7ZDAcHmCAvwD7ZDDMHmCAvwD7ZDCgvQD7ZDBcHiCAvQD7YDweIIC9APtkMOC8iJUwwPtkMJweEIC8iJcwgPtkMEiXsEweEIXwvIXokLW13CBADMzMzMzMzMzMzMVYvsVujXov//i3UIBZMGAABQ/zboRyoAAIkG6MCi//8FkwYAAFD/dgToMioAAIlGBOiqov//BZMGAABQ/3YI6BwqAACJRgjolKL//wWTBgAAUP92DOgGKgAAiUYMXl3CBADMzMzMzMzMzMzMzMzMzFWL7ItFCIvQVot1EIX2dBVXi30MK/iKDBeNUgGISv+D7gF18l9eXcPMzMzMzMzMzFWL7ItNEIXJdB8PtkUMVovxacABAQEBV4t9CMHpAvOri86D4QPzql9ei0UIXcPMzFWL7FaLdQhW6HPz//+L0IvOM9bByRDBwgjBzggz0TPWM8JeXcIEAMzMzMzMzMzMzFWL7FaLdQj/NujC/////3YEiQbouP////92CIlGBOit/////3YMiUYI6KL///+JRgxeXcIEAMzMzMzMzMzMzMxVi+yD7GAPV8DHRaABAAAAVleNRaDHRaQAAAAAUA8RRajHRdABAAAADxFFuMdF1AAAAABmD9ZFyA8RRdgPEUXoZg/WRfjoVqH//wWwAAAAUI1FoFDo5ykAAI1FoFDozkIAAIt9CI1w/4P+AXYsDx8AjUXQUFDoNj8AAFaNRaBQ6LxLAAALwnQLV41F0FBQ6L0+AABOg/4Bd9eNRdBQV+hNRAAAX16L5V3CBADMzMzMzFWL7IPsQFYPV8DHRcABAAAA" & _
                                                    "V41FwMdFxAAAAABQDxFFyMdF4AEAAABmD9ZF2MdF5AAAAAAPEUXoZg/WRfjorqD//4PAEFCNRcBQ6IErAACNRcBQ6HhCAACLfQiNcP+D/gF2KY1F4FBQ6MM+AABWjUXAUOgZSwAAC8J0C1eNReBQUOhKPgAAToP+AXfXjUXgUFfoCkQAAF9ei+VdwgQAzMxVi+yD7BRTVlfoQqD//4tdDIsAiwuNDI0EAAAAUf/QiwuL+Il99I0MjQQAAABRU1foq/3//4PEDOgToP//i3UIiwCLDo0MjQQAAABR/9CLDovYiUX8jQyNBAAAAFFWU+h8/f//g8QM6OSf//+NsJQHAADo2Z///4sOjQyNBAAAAFGLCP/Riw6JRfiNDI0EAAAAUVZQ6Ef9//+DxAzor5///42wmAcAAOikn///iw6NDI0EAAAAUYsI/9GJRQiLBo0EhQQAAABQVot1CFboD/3//4PEDMdF8AEAAADocJ///wWYBwAAUFPotLP//4XAD4RhAQAA6Fef//8FlAcAAFBT6Juz//+FwA+EKgIAAIs7jTS9BAAAAOg1n///VosA/9BWi9hqAFOJXezo4vz//4k7g8QMi130ixuNPJ0EAAAA6Ayf//9XiwD/0FeL8GoAVui8/P//i330g8QMiR6LXexWU/91/FfoJq///4M7AXYRkIsDgzyDAHUISIkDg/gBd/CDPgF2EIsGgzyGAHUISIkGg/gBd/CLB40MhQQAAACFyXQNi8fGAACNQAGD6QF19eicnv//V4tACP/Qi0X8i/iLTfhRiUX0i0UIUFaJXfyJTeyJRfjoh7D//4lFCItF8PfYiUXwi0XsiwiNDI0EAAAAhcl0C8YAAI1AAYPpAXX16Eye////" & _
                                                    "deyLQAj/0IsGjQyFBAAAAIXJdBKLxg8fRAAAxgAAjUABg+kBdfXoIJ7//1aLQAj/0OgVnv//BZgHAABQU+hZsv//i138hcAPhaL+//+LdQiLA40MhQQAAACFyXQNi8PGAACNQAGD6QF19ejcnf//U4tACP/QiweNDIUEAAAAhcl0FIvHDx+AAAAAAMYAAI1AAYPpAXX16LCd//9Xi0AI/9CLffiLB40MhQQAAACFyXQVi8cPH4QAAAAAAMYAAI1AAYPpAXX16ICd//9Xi0AI/9CDffAAD42JAQAAi0UMiziNNL0EAAAA6F+d//9WiwD/0FaL8IlF8GoAVugM+///uAEAAACJPoPEDIlF7DPJi9g7+It9CA+MIgEAAItFDI1WBCvHiUX0i8crxolFCOm9AAAAiwONDIUEAAAAhcl0DYvDxgAAjUABg+kBdfXo+pz//1OLQAj/0IsHjQyFBAAAAIXJdBKLxw8fRAAAxgAAjUABg+kBdfXo0Jz//1eLQAj/0It9+IsHjQyFBAAAAIXJdBWLxw8fhAAAAAAAxgAAjUABg+kBdfXooJz//1eLQAj/0It1CIsGjQyFBAAAAIXJdBWLxg8fhAAAAAAAxgAAjUABg+kBdfXocJz//1aLQAj/0F9eM8Bbi+VdwggAi0UIi3UMOx5/CANF9Is0EOsCM/Y7H38Ii0UIiwQQ6wIzwCvw99Ar8Ykyhcl0BzvwG8lB6wY7xhvJ99mLReyF9ot18A9Fw0ODwgSJRew7Hn6viQaLB40MhQQAAACFyXQNi8fGAACNQAGD6QF19ejtm///V4tACP/QX4vGXluL5V3CCADMzMzMzMzMzMzMzMxVi+yD7CRTVot1EFeLHold+I08nQAAAACJ" & _
                                                    "fdzosJv//1eLAP/QM9KJRRCF234RZpCLDivKiwyOiQyQQjvTfPGLdQiLfQyLz4sGOwcPR86LCYlNCI0ECTvDfw2Lw5krwovI0flBiU0IjRyNAAAAAIld9Oham///U4sA/9CLVQgrFolF/IXSfgyLffyLyjPA86uLfQyLBjPJhcB+G4td/I0Uk4td9CvBjVIEQYsEholC/IsGO8h87ugVm///U4sA/9CLVQiL8CsXiXXshdJ+C4vKM8CL/vOri30MiwczyYXAfhWNFJYrwY1SBEGLBIeJQvyLBzvIfO6LfQiNNP0AAAAAiXXg6Mia//9WiwD/0FeJRQjoig0AAI00hQAAAACJdeToq5r//1aLCP/Ri00QiUXwizFWiXUM6IWw//+JReiFwHQji8jT5ol1DIt1+IP+AX4XuSAAAAAryItFEItABNPoCUUM6wOLdfj/dfBX/3UI/3Xs/3X86Obx////degD//91DIl9+OjWFgAAUGoAVv91EFf/dQjoBu///zv3i8cPTMaJRQyNNIUEAAAA6CCa//9WiwD/0FaL+GoAV+jQ9///i0UMg8QMM9KJB4XAfieLTfiLXQwryItFCI00iA8fQACLD412BItG/CvKQokEjzvTfO6LXfSDPwF2FmYPH0QAAIsHgzyHAHUISIkHg/gBd/CLdfCF9nQUi03khcl0DYvGxgAAjUABg+kBdfXooJn//1aLQAj/0It1CIX2dBSLTeCFyXQNi8bGAACNQAGD6QF19eh6mf//VotACP/Qi03chcl0DotFEMYAAI1AAYPpAXX16FqZ//+LTRBRi0AI/9CLdfyF9nQVhdt0EYvLi8ZmkMYAAI1AAYPpAXX16DCZ//9Wi0AI/9CLdeyF9nQR" & _
                                                    "hdt0DYvGxgAAjUABg+sBdfXoDZn//1aLQAj/0IvHX15bi+VdwgwAzMzMzMzMzMzMzMzMVYvsg+woU4tdEPZDBAFXjXsEdRRT/3UM/3UI6NAEAABfW4vlXcIMAFZT/3UI6G6q//+LG4vLweEFUYlF6Ojerv//i/BW/3UQ6FP4////dRCJRQhW/3Xo6MT8//+JRfCLReiLCI0UjQQAAACF0nQNi8jGAQCNSQGD6gF19ehvmP///3Xoi0AI/9D/dRBW6A6q//+LDolF7I0UjQQAAACF0nQNi87GAQCNSQGD6gF19eg8mP//VotACP/QjTSdAAAAAIl13OgnmP//VosA/9CJReiF234ajU78i9MDyA8fQACLB41/BIkBjUn8g+oBdfHo/Jf//1aLAP/Qi/gzyYl99IXbfimLRQiNVvyDwAQD15CLfQg7D30EizjrAjP/iTpBg8AEg+oEO8t85ot99ItVCIsCjQyFBAAAAIXJdA6LwpDGAACNQAGD6QF19eigl////3UIi0AI/9Dok5f//1aLAP/Qi9CJVeSF234Oi86L+sHpAjPA86uLffRTV1dS6K7y//+LffAzwIXbfiKLVeSNTwSDwvwD1jsHfQSLMesCM/aJMkCDwQSD6gQ7w3zpiweNDIUEAAAAhcl0EYvHDx9AAMYAAI1AAYPpAXX16CCX//9Xi0AI/9CNNN0AAAAAiXXg6AuX//9WiwD/0Iv4iX0I6PyW//9Wiwj/0YlF+DPAhdt+JYtN7IPG/IPBBAP3i33sOwd9BIsR6wIz0okWQIPBBIPuBDvDfOmLdeyLBo0MhQQAAACFyXQUi8YPH4AAAAAAxgAAjUABg+kBdfXooJb//1aLQAj/0FPoZAkAAI0MWwPB" & _
                                                    "jTSFAAAAAIl12OiAlv//VosA/9CLfQiL0ItFDDP2uR8AAACJVfyJdeyJTfA5MH4xixCNPJC4AQAAANPghQd1EoPpAXkJRrkfAAAAg+8EO/J844t9CItV/ItFDIl17IlN8DswD42fAAAAZg8fRAAAhckPiH0AAAC+AQAAANPGkFJT/3X4jQSfUFDoke3//1P/dfz/dfT/dej/dfjojwUAAItNDIsBK0XshTSBdCb/dfyLRfhTV/915I0EmFDoXu3//1P/dfz/dfT/dehX6F4FAADrCIvHi334iUX4i0Xwi1X8SNHOiUXwhcB5lIt17ItFDIl9CEa5HwAAAIl17IlN8DswD4xn////U1L/dfT/dehX6BkFAACLRRCLOI00vQQAAADoaJX//1aLAP/QVmoAUIlFEOgX8///i1UQg8QMiTqF234qi0UIjTydAAAAAAPHM/aJRQyLCosAK85GiQSKi0UMg8AEiUUMO/N86esDi33cgzoBi3UIdhZmDx9EAACLAoM8ggB1CEiJAoP4AXfwi138hdt0FItN2IXJdA2Lw8YAAI1AAYPpAXX16OCU//9Ti0AI/9CF9nQUi03ghcl0DYvGxgAAjUABg+kBdfXovZT//1aLQAj/0Itd+F6F23Qai03ghcl0E4vDZg8fRAAAxgAAjUABg+kBdfXokJT//1OLQAj/0Itd9IXbdBOF/3QPi8+Lw8YAAI1AAYPpAXX16GuU//9Ti0AI/9CLXeiF23QZhf90FYvPi8NmDx9EAADGAACNQAGD6QF19ehAlP//U4tACP/Qi13khdt0EYX/dA2Lw8YAAI1AAYPvAXX16B2U//9Ti0AI/9CLRRBfW4vlXcIMAMzMzMzMzMzMzMzMzFWL7IPs"
Private Const STR_THUNK4                As String = "MFNWV4t9EFf/dQjom6X//4sfiUXsiV3gjTSdAAAAAIl11OjUk///VosI/9Ez0olF+IXbfhVmDx9EAACLDyvKiwyPiQyQQjvTfPHorJP//1aLAP/Qi33si9OL8Il12CsXhdJ+C4vKM8CL/vOri33siwczyYXAfhaNFJaQK8GNUgRBiwSHiUL8iwc7yHzujTzdAAAAAIl93Ohfk///V4sA/9CL8Il1COhQk///V4sI/9GJRfyNBBuFwH4Pi8iL/jPA86uNPN0AAAAAU8dEN/wBAAAA6PQFAACNNIUAAAAAiXXQ6BWT//9WiwD/0IlF8DP2i0UMiXX0jX4fiwiJTeSFyX4ujRSIi9kPH0AAuAEAAACLz9PghQJ1EoPvAXkJRr8fAAAAg+oEO/N84Ytd4Il19ItF+IsAUIlF6OitqP//i1XoiUXghcB0GovI0+KD+wF+EbkgAAAAK8iLRfiLQATT6AvQUugjDwAAiUXoO3XkD42uAAAAi0UMDx9AAIX/D4iOAAAAi8++AQAAANPGi00I/3XwjQSZU/91/FBQ6Ovp////deCNBBv/dehqAFP/dfhQ/3X86BPn//+LTQyLAStF9IU0gXQz/3Xwi0X8U/91CP912I0EmFDosOn///914I0EG/916GoAU/91+FD/dQjo2Ob//4tNCOsMi0UIi038iU0IiUX80c6D7wF5hIt19ItFDEa/HwAAAIl19DswD4xZ////i30Qiz+NNL0EAAAA6MiR//9WiwD/0FZqAFCJRRDod+///4tVEIPEDIk6hdt+KotFCI08nQAAAAADxzP2iUUMiwqLACvORokEiotFDIPABIlFDDvzfOnrA4t91IM6AYt1CHYWZg8fRAAAiwKDPIIAdQhIiQKD+AF38IX2dBeLTdyF" & _
                                                    "yXQQi8YPHwDGAACNQAGD6QF19ehAkf//VotACP/Qi13whdt0FItN0IXJdA2Lw8YAAI1AAYPpAXX16BqR//9Ti0AI/9CLXfyF23QYi03chcl0EYvDDx9AAMYAAI1AAYPpAXX16PCQ//9Ti0AI/9CF/3QRi0X4i8+QxgAAjUABg+kBdfXo0JD///91+ItACP/Qi13Yhdt0EYX/dA2Lw8YAAI1AAYPvAXX16KuQ//9Ti0AI/9CLdeyLBo0MhQQAAACFyXQQi8YPHwDGAACNQAGD6QF19eiAkP//VotACP/Qi0UQX15bi+VdwgwAzMzMzMzMzMzMzMzMzMxVi+yLVQhTi10UVot1GFeNBHaNPINXVlP/dRCNBLJQiUUU6Bnq//9Xi30YjTSzV1b/dQxT6Lfn//+LXQiNBD9QU1ZT6Bjk//8zyYlFGIX/fh+NNLuL1osEi41SBIlC/McEiwAAAABBO89864tFGOsDi3UUhcB1PIX/fjiLXQyL1osKOwyDdRpAg8IEO8d88VdW/3UMVugG6///X15bXcIUADvHfQ6LVQiNDDiLDIo7DIN2C1dW/3UMVuji6v//X15bXcIUAMzMzMzMzMzMzMzMVYvsgewAAQAAi0UMD1fAU1ZXuTwAAABmDxOFAP///421AP///8dF/BAAAACNvQj////zpYtNEI2dCP///4PBEIvTK8KJTfiJRQxmDx9EAACL+cdFEAQAAACL8w8fRAAA/3QYBP80GP939P938Oj+m///AUb4i0UMEVb8/3QYBP80GP93/P93+Ojjm///AQaLRQwRVgT/dBgE/zQY/3cE/zfoypv//wFGCItFDBFWDP90GAT/NBj/dwz/dwjor5v//wFGEI1/IItFDBFW" & _
                                                    "FI12IINtEAF1iotN+IPDCINt/AEPhWr///8z9moAaib/dPWE/3T1gOh3m///AYT1AP///2oAEZT1BP///2om/3T1jP909YjoWJv//wGE9Qj///9qABGU9Qz///9qJv909ZT/dPWQ6Dmb//8BhPUQ////agARlPUU////aib/dPWc/3T1mOgam///AYT1GP///2oAEZT1HP///2om/3T1pP909aDo+5r//wGE9SD///8RlPUk////g8YFg/4PD4JZ////i10IjbUA////uSAAAACL+/OlU+hppP//U+hjpP//X15bi+VdwgwAzMzMzMzMzMzMzFWL7IPsEFNWi3UMV4t9GGoAVmoA/3UU6JSa//9qAFZqAFeJRfCL2uiEmv//agD/dRCJRfSL8moAV+hymv//agD/dRCJRfxqAP91FIlV+Ohdmv//i/iLRfQD+4PSAAP4E9Y71ncOcgQ7+HMIg0X8AINV+AGLRQgzyQtN8IkIM8kDVfyJeAQTTfhfXolQCIlIDFuL5V3CFADMzMzMzMzMzMxVi+yLTQhWM/aD+TJ+FGaQi8GZK8LR+CvIQY00joP5Mn/ui8ZeXcIEAMzMzMzMzMxVi+yD7DBTVot1CFeLfQxXVug6MAAAaiBXjUXQUOjuGgAAi9iJVQyNTgiDxjiNRdBQUVHoaBUAAAPDiQYTVQyLRQiDwBCJVgRXUFDoUBUAAIt9CIlHQI1F0FBXV4lXROjMMwAAi00MA9gTyotXMCvTi180G9k7XzRyLXcFO1cwdiaDBv+LBoNWBP8jRgSD+P91FYNGCP+NdgiDVgT/iwYjRgSD+P9064lfNIlXMF9eW4vlXcIIAMzMzMzMzMzMzMxVi+yB7AgBAACNhXj///9T" & _
                                                    "Vlf/dQxQ6GULAACNhXj///9Q6Jmi//+NhXj///9Q6I2i//+NhXj///9Q6IGi//+Nvfj+//+7AgAAAGYPH0QAAIuNeP///4uFfP///4Hp7f8AAImN+P7//4PYAImF/P7//7gIAAAAZmYPH4QAAAAAAIt0B/iLTAf8i5QFeP///4l1+A+szhCLjAV8////g+YBx0QH/AAAAAAr1oPZAIHq//8AAImUBfj+//+D2QCJjAX8/v//D7dN+IlMB/iDwAiD+HhyrIuNaP///4uFbP///4tV8A+swRAPt4Vo////g+EBiYVo////K9HHhWz///8AAAAAi030uAEAAACD2QCB6v9/AACJlXD///+D2QCJjXT///8PrMoQg+IBwfkQK8JQjYX4/v//UI2FeP///1DoTQkAAIPrAQ+FBP///4t1CDPSioTVeP///4uM1Xj///+IBFaLhNV8////D6zBCIhMVgFCwfgIg/oQctdfXluL5V3CCADMzMzMzMzMzMzMzMzMVYvsi0UIM9JWV4t9DCv4jXIRiwwHjUAEA0j8A9EPtsrB6giJSPyD7gF1519eXcIIAMzMzMzMzMzMzMzMzMzMzFWL7Fb/dQyLdQhW6LD///+NRkRQVui2AQAAXl3CCADMVYvsg+xEU1aLdQhXDxAGi0ZAiUX8DxFFvA8QRhAPEUXMDxBGIA8RRdwPEEYwDxFF7Ogaiv//BUQFAABQjUW8UOhb////i0X8jX2899CNVcwlgAAAACv+uQIAAACNWP/30MHoH8HrHyPY99sr1ovD99CJRQhmD27DZg9w0ABmD27Ai8ZmD3DYAA8fhAAAAAAAjUAgDxBA4A8QTAfgZg/bwmYP28tmD+vIDxFI4A8QQPAPEEwC" & _
                                                    "4GYP28JmD9vLZg/ryA8RSPCD6QF1xo1WQI1xAYsMOo1SBCNNCIvDI0L8C8iJSvyD7gF16F9eW4vlXcIEAMzMzMzMzMzMzMzMzMzMzFWL7IPsRI1FvFZqRGoAUOj85v//i3UIg8QMM8CLlqgAAACF0nQbZmYPH4QAAAAAAA+2jAaYAAAAiUyFvEA7wnLvjUW8x0SVvAEAAABQVuiN/v//XovlXcIEAMzMzMzMzFWL7FaLdQgzwDPSDx9EAAADBJYPtsiJDJZCwegIg/oQfO4DRkCLyMHoAoPhAzPSiU5AjQyAAwyWD7bBiQSWQsHpCIP6EHzuAU5AXl3CBADMVYvsg+xUi0UMjU2sU1aLdQgz2yvBx0X4EAAAAFeJRfAz0jP/M8CJVQiJVfyF23hRjUsBg/kCfDCLTfCNVayNDJkD0YsMho1S+A+vSggBTQiLTIYEg8ACD69KBAFN/I1L/zvBft6LVQg7w38Oi30Mi8sryIs8jw+vPIaLRfwDwgP4jUMBM9KJVQiLyIlV/IlF9IP4EX1yg334AnxDi1UMi8MrwY0UgoPCQA8fgAAAAACLBI6NUvgPr0IMjQSAweAGAUUIi0SOBIPBAg+vQgiNBIDB4AYBRfyD+RB81ItVCIP5EX0ai1UMi8MrwYtEgkQPrwSOi1UIjQSAweAGA/iLRfwDwgP4i0X0i034SYl8nayJTfiL2IP5/w+PAv///41FrFDoif7//w8QRayLRexfDxEGDxBFvA8RRhAPEEXMDxFGIA8QRdwPEUYwiUZAXluL5V3CCADMzMzMzMzMzMzMzFWL7ItVDIPsRDPADx9EAAAPtgwQiUyFvECD+BB88o1FvMdF/AEAAABQ/3UI6J/8//+L5V3CCADM" & _
                                                    "zMzMzMzMzMxVi+yB7HwBAABTVldqDP91DI1F4MZF3AAPV8DHReUAAAAAUGYP1kXdZsdF6QAAxkXrAOhZ5P//g8QMxkW8AI1F3MdF1QAAAAAPV8Bmx0XZAAAPEUW9agRQaiD/dQiNhTD///9mD9ZFzVDGRdsA6P6r//9qII1FvFBQjYUw////UOhLpf//jUXMUI1FvFCNhYT+//9Q6De3//8PV8APEUW8ikW8aiCNRbxQUI2FMP///1APEUXM6Bal//+LdRQPV8BW/3UQDxFFvIpFvI2FhP7//8ZF7ABQDxFFzMdF9QAAAABmD9ZF7WbHRfkAAMZF+wDoq7f//4vG99iD4A9QjUXsUI2FhP7//1Dok7f//4N9JAGLfSCLXRxTdRRX/3UYjYUw////UOimpP//U1frA/91GI2FhP7//1DoY7f//4vD99iD4A9QjUXsUI2FhP7//1DoS7f//zPSiF30i8aJVeiIReyLyIvCD6zBCGoQwegIiE3ti8KLzg+swRDB6BCITe6LwovOD6zBGMHoGA+2wohF8IvCwegIiEXxi8LB6BCIRfLB6hiITe+Ly4hV8zPSi8KJVegPrMEIwegIiE31i8KLyw+swRDB6BCITfaLwovLD6zBGMHoGA+2wohF+IvCwegIiEX5i8LB6BCIRfqNRexQjYWE/v//weoYUIhN94hV++ibtv//g30kAXUz/3UojYWE/v//UOg2tf//anyNhTD///9qAFDoluL//4qFMP///4PEDDPAX15bi+VdwiQAjUWsUI2FhP7//1DoArX//4t1KI1NrIvBMtu6EAAAACvwkIoEDo1JATJB/wrYg+oBdfCLRRyE23U/UFf/dRiNhTD///9Q6Eij//9qfI2F" & _
                                                    "MP///2oAUOgo4v//ioUw////g8QMD1fADxFFrIpFrF9eM8Bbi+VdwiQAhcB0DlBqAFfo/eH//4oHg8QManyNhTD///9qAFDo6OH//4qFMP///4PEDA9XwA8RRayKRaxfXrgBAAAAW4vlXcIkAMzMzMzMzMxVi+xWV4t9CA+2B5mLyIvyD7ZHAQ+kzgiZweEIC/ILyA+2RwIPpM4ImcHhCAvyC8gPtkcDD6TOCJnB4QgL8gvID7ZHBA+kzgiZweEIC/ILyA+2RwUPpM4ImcHhCAvyC8gPtkcGD6TOCJnB4QgL8gvID7ZHBw+kzgiZweEIC8EL1l9eXcIEAMzMzMzMzMzMzMxVi+xTi10IM9KLy8HpEEFWV42B////f/fxi/DB5hCLxvfji8gDzvfRg9IAM8CDwQH30hPAA8L35sHoH4vyA/YL8IvG9+MDxov6g9cAgf8AAACAciP30zPJA8MTyYPAAYPRAE8D+U6B/wAAAIBz6F+Lxl5bXcIEADPSA8MT0gPXgfoAAACAcxEzyUYDwxPJA9GB+gAAAIBy71+Lxl5bXcIEAMzMzMzMzMzMzMzMzMzMzFWL7FNWi3UIV1b/dQzojpf//1b/dRCL+Il9COiAl///Vv91FIvYiV0M6HKX//9QU1eJRRDol+n//4vYhfZ0F4t9GE5WU+h2lv//iAeNfwGF9nXvi30IiweNDIUEAAAAhcl0DYvHxgAAjUABg+kBdfXoS4L//1eLQAj/0It1DIsGjQyFBAAAAIXJdBCLxg8fAMYAAI1AAYPpAXX16CCC//9Wi0AI/9CLdRCLBo0MhQQAAACFyXQVi8YPH4QAAAAAAMYAAI1AAYPpAXX16PCB//9Wi0AI/9CLA40MhQQAAACF" & _
                                                    "yXQNi8PGAACNQAGD6QF19ejLgf//U4tACP/QX15bXcIUAMzMzMzMzMzMzMzMzMzMVYvsg+wIi0UQSPfQmVOLXQiJRfiLRQyJVfzzD35d+I1LeFYz9mYPbNuNUHg7wXdLO9NyRyvYx0UQEAAAAFdmkIs8GI1ACIt0GPyLSPiLUPwzzyNN+DPWI1X8M/kz8ol8GPiJdBj8MUj4MVD8g20QAXXOX15bi+VdwgwAi9ONSBAr0A8QDPONSSAPEFHQZg/v0WYP29MPKMJmD+/BDxEE84PGBA8QQdBmD+/QDxFR0A8QTArgDxBR4GYP79FmD9vTDyjCZg/vwQ8RRArgDxBB4GYP78IPEUHgg/4QcqVeW4vlXcIMAMzMzMzMzMzMzMzMVYvsi1UMi0UIK9BWvhAAAACLDAKNQAiJSPiLTAL8iUj8g+4BdeteXcIIAMzMzMzMVYvsi0UQVleD+BB0OYP4IHVfi3UMi30IahBWV+j/3f//ahCNRhBQjUcQUOjw3f//g8QY6FiA//8FMQUAAIlHMF9eXcIMAIt1DIt9CGoQVlfoy93//2oQjUcQVlDov93//4PEGOgngP//BSAFAACJRzBfXl3CDADMzMzMzMzMzMxVi+yD7GyLRQiNVZRTVrugAQAAM/aLSASJTfiLSAiJTfSLSAyJTeiLSBCJTfyLSBSJTfCLSBiJTeyLTQyDwQKJddxXizgr04tAHIl94IlF5IlN2IldDIlV1A8fgAAAAACD/hBzKQ+2cf4PtkH/weYIC/APtgHB5ggL8A+2QQHB5ggL8IPBBIk0GolN2OtUjV4Bg+YPjUP9g+APjX2UjTy3i0yFlIvDg+APi/HBxg+LVIWUi8HBwA0z8MHpCjPxi8KLysHI" & _
                                                    "B8HBDjPIweoDjUP4M8qLXQyD4A8D8QN0hZQDN4k36Cl///+LffyL18HKC4vPwcEHM9GLz8HJBvfXI33sM9GLDBiDwwSLRfADyiNF/APOi3XgM/iL1oldDMHKDYvGwcAKA/kDfeQz0IvGwcgCM9CLRfiLyCPGM84jTfQzyItF7IlF5APRi0Xwi034iUXsi0X8iUXwi0XoA8eJdfiLddwD+otV1EaJRfyLRfSJTfSLTdiJReiJfeCJddyB+6ACAAAPgtf+//+LRQiLTfiLVfwBSASLTfQBSAgBUBABOItN6ItV8AFIDAFQFItV7ItN5AFQGAFIHP9AYF9eW4vlXcIIAMzMzMzMzMzMzMzMzFWL7IHs4AAAAFNWi3UIu6ACAABXiV24iwaJReyLRgSJRfCLRgyLfgiJReCLRhCJRdSLRhSJRdCLRhiJRbSLRhyJRbCLRiCJReiLRiSJRfSLRiiJRcyLRiyJRciLRjCJRcSLRjSJRcCLRjiJRayLRjyLdQyJfdiNvSD///+JRagzwCv7iUXciX2gDx+AAAAAAIP4EHMfVuil+f//i8iDxgiLwolNDIlF5IkMH4lEHwTpEwEAAI1QAcdFDAAAAACNQv2D4A+LjMUg////i4TFJP///4lF+IvCg+APiU38jY0g////i5TFIP///4v6i5zFJP///4tF3IPgD4lVvMHnGI0EwYvLiUWki8IPrMgICUUMi0W8wekIC/mLyw+syAGJfeSL+tHpM9IL0MHnHzFVDAv5i0W8i03kD6zYBzPPMUUMi0X8wesHM8sz24lN5ItN+IvRD6TBA8HqHcHgAwvZi034C9CLRfyL+A+syBOJVbwz0gvQwekTi0W8M8LB5w2LVfwL+YtN+DPf" & _
                                                    "D6zKBjPCwekGi1UMM9mLTeQD0ItF3BPLg8D5g+APA5TFIP///xOMxST///+LRaQDEIlVDBNIBIkQiU3kiUgE6HR8//+LVfQz/4tN6IvaD6TKF8HrCQv6weEXi1X0C9mLTeiJXfyL2Q+s0RKJffgz/wv5weoSMX38M/+LTejB4w4L2otV9DFd+IvZD6zRDsHjEgv5weoOMX38C9qLTfiLVbgzy4td/It96PfXAxwQE0wQBCN9xItV9ItFyPfSI0X0I1XAM9CJTfiLTcwjTeiLRfgz+YtN8APfE8IDXQwTReQDXayJXfwTRagz24lF+ItF7IvQD6zIHMHiBMHpHAvYi0XsC9GLTfCL+Q+kwR6JVQwz0sHvAgvRweAeC/gz3zFVDDPSi03wi/mLRewPpMEZwe8HC9HB4BkxVQwL+ItN2DPfi1Xgi/kzfewjfdQjTewzVfAz+SNV0ItF4CNF8ItNxDPQi0UMA9+LffgTwolNrItNwItV/ANVtIlNqBN9sItNzANd/IlNxItNyIlNwItN6IlNzItN9Il99It91Il9tIt90Il9sIt92Il91It94Il90It97IlNyIvIE034i0XciV3sQItduIl92IPDCIt98Il94It9oIlV6IlN8IlF3IlduIH7IAUAAA+CG/3//4t1CItF7It92AEGi0XgEU4Ei8oBfgiLfbQRRgyLRdQBRhCLRdARRhQBfhiLRbARRhwBTiCLRfQRRiSLRcwBRiiLRcgRRiyLRcQBRjCLRcARRjSLTawBTjiLTagRTjz/hsAAAABfXluL5V3CCADMzMzMzMzMzMzMzMzMzFWL7FOLXQhWVw+2ewcPtkMKD7ZzCw+2Uw/B5wgL+A+2SwMPtkMNwecIC/jB" & _
                                                    "5ggPtgPB5wgL+MHiCA+2Qw4L8MHhCA+2QwHB5ggL8A+2QwTB5ggL8A+2QwIL0A+2QwXB4ggL0A+2QwjB4ggL0A+2QwYLyIl7BA+2QwnB4QgLyIlzCA+2QwzB4QhfC8iJUwxeiQtbXcIEAMzMzMzMzMzMzMxVi+yLRQxQUP91COgQ6v//XcIIAMzMzMzMzMzMzMzMzFWL7ItFEFNWi3UIjUh4V4t9DI1WeDvxdwQ70HMLjU94O/F3MDvXciwr+LsQAAAAK/CLFDgrEItMOAQbSASNQAiJVDD4iUww/IPrAXXkX15bXcIMAIvXjUgQi94r0CvYK/64BAAAAI12II1JIA8QQdAPEEw34GYP+8gPEU7gDxBMCuAPEEHgZg/7yA8RTAvgg+gBddJfXltdwgwAzMzMzMxVi+xW6Od4//+LdQgFiAUAAFD/NuhXAAAAiQbo0Hj//wWIBQAAUP92BOhCAAAAiUYE6Lp4//8FiAUAAFD/dgjoLAAAAIlGCOikeP//BYgFAABQ/3YM6BYAAACJRgxeXcIEAMzMzMzMzMzMzMzMzMzMVYvsi1UMU4tdCIvDwegYi8tWwekID7bJD7Y0EIvDwegQD7bAD7YMEcHmCA+2BBALxsHgCAvBD7bLweAIXlsPtgwRC8FdwggAzMzMzMzMzMxVi+yLTQxTi10IVoPDEMdFDAQAAABXg8EDDx+AAAAAAA+2Qf6NWyCZjUkIi/CL+g+2QfUPpPcImcHmCAPwiXPQE/qJe9QPtkH3mYvwi/oPtkH4mQ+kwgjB4AgD8Ilz2BP6iXvcD7ZB+pmL8Iv6D7ZB+Q+k9wiZweYIA/CJc+AT+ol75A+2QfyZi/CL+g+2QfsPpPcImcHmCAPwiXPoE/qD" & _
                                                    "bQwBiXvsD4V0////i00IX15bgWF4/38AAMdBfAAAAABdwggAzMzMzMzMzMzMzMzMVYvsg+wMU4tdDA9XwFZXi30QixOL8otDBIvIZg8TRfQDNxNPBDvydQY7yHUE6xg7yHcPcgQ78nMJuAEAAAAz0usLZg8TRfSLRfSLVfiLfQiJTwSJN4tLCIt1EIlN/ItLDIlN+ItOCANN/IlNCItODBNN+ItdCAPYiV0IE8o7XfyLXQx1BTtLDHQjO0sMdxNyCItDCDlFCHMJuAEAAAAz0usLZg8TRfSLVfiLRfSLdQiJTwyJdwiLSxCLdRCJTfyLSxSJTfiLThADTfyJTQiLThQTTfiLXQgD2IldCBPKO138i10MdQU7SxR0IztLFHcTcgiLQxA5RQhzCbgBAAAAM9LrC2YPE0X0i1X4i0X0i3UIiU8UiXcQi0sYi3UQiU38i0sciU34i04YA038iU0Ii04cE034i10IA9iJXQgTyjtd/ItdDHUFO0scdCM7Sxx3E3IIi0MYOUUIcwm4AQAAADPS6wtmDxNF9ItV+ItF9It1CIlPHIl3GItLIIt1EIlN/ItLJIlN+ItOIANN/IlNCItOJBNN+ItdCAPYiV0IE8o7XfyLXQx1BTtLJHQjO0skdxNyCItDIDlFCHMJuAEAAAAz0usLZg8TRfSLVfiLRfSLdQiJdyCLdRCJTySLSyiLWyyLdigD8YlNDItNEItJLBPLA/ATyjt1DHUEO8t0LDvLdx1yBTt1DHMWiXcouAEAAACJTywz0l9eW4vlXcIMAGYPE0X0i1X4i0X0iXcoiU8sX15bi+VdwgwAzMzMzMzMVYvsg+wIU4tdDA9XwFZXi30QixOL8otDBIvIZg8TRfgDNxNP" & _
                                                    "BDvydQY7yHUE6xg7yHcPcgQ78nMJuAEAAAAz0usLZg8TRfiLRfiLVfyLfQiJTwSLTRCJN4txCANzCItJDBNLDAPwE8o7cwh1BTtLDHQgO0sMdxByBTtzCHMJuAEAAAAz0usLZg8TRfiLVfyLRfiJTwyLTRCJdwiLcRADcxCLSRQTSxQD8BPKO3MQdQU7SxR0IDtLFHcQcgU7cxBzCbgBAAAAM9LrC2YPE0X4i1X8i0X4iU8UiXcQi0sYi1sciU0Mi00Qi3EYA3UMi0kcE8sD8BPKO3UMdQQ7y3QsO8t3HXIFO3UMcxaJdxi4AQAAAIlPHDPSX15bi+VdwgwAZg8TRfiLVfyLRfiJdxiJTxxfXluL5V3CDADMzMzMzMxVi+yLTQjHAQAAAADHQQQAAAAAiwGJQQiLQQSJQQyLQQiJQRCLQQyJQRSLQRCJQRiLQRSJQRyLQRiJQSCLQRyJQSSLQSCJQSiLQSSJQSxdwgQAzMzMzMzMzMzMzMzMzMxVi+yLRQjHAAAAAADHQAQAAAAAx0AIAAAAAMdADAAAAADHQBAAAAAAx0AUAAAAAMdAGAAAAADHQBwAAAAAXcIEAMzMzMzMzMzMzMzMzMzMzFWL7ItNDLoFAAAAU4tdCFYr2Y1BKFeJXQgPH4AAAAAAizQDi1wDBIt4BIsIO993LnIiO/F3KDvfchp3BDvxchSLXQiD6AiD6gF51V9eM8BbXcIIAF9eg8j/W13CCABfXrgBAAAAW13CCADMzMzMzMxVi+yLTQy6AwAAAFOLXQhWK9mNQRhXiV0IDx+AAAAAAIs0A4tcAwSLeASLCDvfdy5yIjvxdyg733IadwQ78XIUi10Ig+gIg+oBedVfXjPAW13CCABfXoPI" & _
                                                    "/1tdwggAX164AQAAAFtdwggAzMzMzMzMVYvsi1UIM8APH4QAAAAAAIsMwgtMwgR1D0CD+AZy8bgBAAAAXcIEADPAXcIEAMzMVYvsi1UIM8APH4QAAAAAAIsMwgtMwgR1D0CD+ARy8bgBAAAAXcIEADPAXcIEAMzMVYvsg+wQU4tdELlAAAAAVot1CCvLV4t9DGYPbsOJTRCLB4tXBIlF+IlV/PMPfk34Zg/zyGYP1g7o437//4tNEIlF8ItHCIlV9ItXDIlF+IlV/PMPfk34Zg9uw2YP88jzD35F8GYP68hmD9ZOCOiufv//i00QiUXwi0cQiVX0i1cUiUX4iVX88w9+TfhmD27DZg/zyPMPfkXwZg/ryGYP1k4Q6Hl+//+LTRCJRfCLRxiJVfSLVxyJRfiJVfzzD35N+GYPbsNmD/PI8w9+RfBmD+vIZg/WThjoRH7//4tNEIlF8ItHIIlV9ItXJIlF+IlV/PMPfk34Zg9uw2YP88jzD35F8GYP68hmD9ZOIOgPfv//iUXwi0coiVX0i1csiUX4iVX88w9+TfiLTRBmD27DZg/zyPMPfkXwZg/ryGYP1k4o6Np9//9fXluL5V3CDADMVYvsg+wQU4tdELlAAAAAVot1CCvLV4t9DGYPbsOJTRCLB4tXBIlF+IlV/PMPfk34Zg/zyGYP1g7ok33//4tNEIlF8ItHCIlV9ItXDIlF+IlV/PMPfk34Zg9uw2YP88jzD35F8GYP68hmD9ZOCOheff//i00QiUXwi0cQiVX0i1cUiUX4iVX88w9+TfhmD27DZg/zyPMPfkXwZg/ryGYP1k4Q6Cl9//+LTRCJRfCLRxiJVfSLVxyJRfiJVfzzD35N+GYPbsNmD/PI8w9+"
Private Const STR_THUNK5                As String = "RfBmD+vIZg/WThjo9Hz//19eW4vlXcIMAMzMzMzMzMzMzMzMVYvsg+xoU1aLdQyNXjBT6Ez9//+FwA+FZAIAAFcPHwCNRZgPV8BQZg8TRfjon/v//41FyFDolvv//1ONRZhQ6Gzi//9T6Ib7//+LFov6A32Yi0YEi8gTTZw7+nUGO8h1BOsbO8h3D3IEO/pzCbgBAAAAM9LrDg9XwGYPE0X4i0X4i1X8i14MiT6LfggDfaCJTgSLyxNNpAP4E8o7fgh1BDvLdCI7y3cQcgU7fghzCbgBAAAAM9LrDg9XwGYPE0X4i1X8i0X4i14UiX4Ii34QA32oiU4Mi8sTTawD+BPKO34QdQQ7y3QiO8t3EHIFO34Qcwm4AQAAADPS6w4PV8BmDxNF+ItV/ItF+IteHIl+EIt+GAN9sIlOFIvLE020A/gTyjt+GHUEO8t0IjvLdxByBTt+GHMJuAEAAAAz0usOD1fAZg8TRfiLVfyLRfiLXiSJfhiLfiADfbiJThyLyxNNvAP4E8o7fiB1BDvLdCI7y3cQcgU7fiBzCbgBAAAAM9LrDg9XwGYPE0X4i1X8i0X4i14siX4gi34oA33AiU4ki8sTTcQD+BPKO34odQQ7y3QiO8t3EHIFO34ocwm4AQAAADPS6w4PV8BmDxNF+ItV/ItF+IteNIl+KIt+MAN9yIlOLIvLE03MA/gTyjt+MHUEO8t0IjvLdxByBTt+MHMJuAEAAAAz0usOD1fAZg8TRfiLVfyLRfiLXjyJfjCLfjgDfdCJTjSLyxNN1AP4E8o7fjh1BDvLdCI7y3cQcgU7fjhzCbgBAAAAM9LrDg9XwGYPE0X4i1X8i0X4iU48jV4wi03YA8iJfjiLRdwTwgFOQFMRRkTo6fr//4XAD4Sh/f//" & _
                                                    "X+grbf//BbAAAABQVujv+f//hcB+J+gWbf//BbAAAABQVlboORQAAOgEbf//BbAAAABQVujI+f//hcB/2Vb/dQjoOxAAAF5bi+VdwggAzMzMVYvsg+woU4tdCFZXi30MV1PoehAAAItHLA9XwIlF5ItHMIlF6ItHNIlF7ItHOIlF8ItHPIlF9I1F2GoBUFBmDxNF2MdF4AAAAADo8fv//4vwjUXYUFNT6GT3//+LTzgD8ItHMItXPIlF5DPAC0c0iUXojUXYagFQUMdF4AAAAACJTeyJVfDHRfQAAAAA6K77//8D8I1F2FBTU+gh9///A/DHReQAAAAAi0cgD1fAiUXYi0ckiUXci0coiUXgi0c4iUXwi0c8iUX0jUXYUFNTZg8TRejo5/b//4tPJAPwM8CJTdgLRyiJRdyLRzCLVzSLyolF+DPAC0csiUXgi0c4iUXoi0c8iUXsM8ALRyCJRfSNRdhQU1OJTeSJVfDon/b//4tPLAPwi1c0M8ALRzAPV8CJRdyLRyCJRfCNRdiJTdgzyQtPKFBTU4lV4MdF5AAAAABmDxNF6IlN9OjBFAAAi1ckK/CLRzAPV8CJRdixIItHNIlF3ItHOIlF4ItHPIlF5ItHIGYPE0Xo6IJ4//8LVyyJRfCNRdhQU1OJVfTofhQAAItVDCvwi080M8ALRziLXySJRdwzwAtHPIlN2ItPIDP/iUXgi0Ioi1IsiU3ksSDoG3j//wvYx0XwAAAAAIld6Av6i10MiX3si30Ii0MwiUX0jUXYUFdX6CMUAAAr8MdF4AAAAACLQziJRdiLQzyJRdyLQySJReSLQyiJReiLQyyJReyLQzSJRfSNRdhQV1fHRfAAAAAA6OQTAAAr8Hkg6Jtq" & _
                                                    "//+DwBBQV1focPX//wPweOxfXluL5V3CCAAPHwCF9nUUV+h2av//g8AQUOit9///g/gBdNzoY2r//4PAEFBXV+iYEwAAK/Dr1MzMzMxVi+xW/3UQi3UI/3UMVujd8v//C8J1Df91FFboAPf//4XAeAr/dRRWVuhSEQAAXl3CEADMzMzMzMzMzMzMzMzMVYvsVv91EIt1CP91DFbo3fT//wvCdQ3/dRRW6DD3//+FwHgK/3UUVlboIhMAAF5dwhAAzMzMzMzMzMzMzMzMzFWL7IHsyAAAAFaLdQxW6G33//+FwHQP/3UI6NH1//9ei+VdwgwAV1aNhTj///9Q6OwMAACLfRCNhWj///9XUOjcDAAAjUXIUOij9f//jUWYx0XIAQAAAFDHRcwAAAAA6Iz1//+NhWj///9QjYU4////UOgp9v//i9CF0g+EvgEAAFOLjTj///8PV8CD4QFmDxNF+IPJAHUvjYU4////UOi8CwAAi0XIg+ABg8gAD4S/AAAAV41FyFBQ6LLx//+L8Iva6bEAAACLhWj///+D4AGDyAB1L42FaP///1DofwsAAItFmIPgAYPIAA+EEQEAAFeNRZhQUOh18f//i/CL2ukDAQAAhdIPjo8AAACNhWj///9QjYU4////UFDo4A8AAI2FOP///1DoNAsAAI1FmFCNRchQ6Gf1//+FwHkLV41FyFBQ6Cjx//+NRZhQjUXIUFDoqg8AAItFyIPgAYPIAHQRV41FyFBQ6ATx//+L8Iva6waLXfyLdfiNRchQ6N8KAAAL8w+EmAAAAItF8IFN9AAAAICJRfDphgAAAI2FOP///1CNhWj///9QUOhRDwAAjYVo////UOilCgAAjUXIUI1FmFDo2PT/" & _
                                                    "/4XAeQtXjUWYUFDomfD//41FyFCNRZhQUOgbDwAAi0WYg+ABg8gAdBFXjUWYUFDodfD//4vwi9rrBotd/It1+I1FmFDoUAoAAAvzdA2LRcCBTcQAAACAiUXAjYVo////UI2FOP///1DobPT//4vQhdIPhUT+//9bjUXIUP91COjVCgAAX16L5V3CDADMzMzMzMzMzMzMzMzMVYvsgeyIAAAAVot1DFboPfX//4XAdA//dQjo0fP//16L5V3CDABXVo2FeP///1Do7AoAAIt9EI1FmFdQ6N8KAACNRdhQ6Kbz//+NRbjHRdgBAAAAUMdF3AAAAADoj/P//41FmFCNhXj///9Q6D/0//+L0IXSD4SwAQAAUw8fQACLjXj///8PV8CD4QFmDxNF+IPJAHUvjYV4////UOi+CQAAi0XYg+ABg8gAD4S2AAAAV41F2FBQ6JTx//+L8Iva6agAAACLRZiD4AGDyAB1LI1FmFDohwkAAItFuIPgAYPIAA+ECAEAAFeNRbhQUOhd8f//i/CL2un6AAAAhdIPjowAAACNRZhQjYV4////UFDomw8AAI2FeP///1DoPwkAAI1FuFCNRdhQ6ILz//+FwHkLV41F2FBQ6BPx//+NRbhQjUXYUFDoZQ8AAItF2IPgAYPIAHQRV41F2FBQ6O/w//+L8Iva6waLXfyLdfiNRdhQ6OoIAAAL8w+EkgAAAItF8IFN9AAAAICJRfDpgAAAAI2FeP///1CNRZhQUOgPDwAAjUWYUOi2CAAAjUXYUI1FuFDo+fL//4XAeQtXjUW4UFDoivD//41F2FCNRbhQUOjcDgAAi0W4g+ABg8gAdBFXjUW4UFDoZvD//4vwi9rrBotd/It1+I1FuFDo" & _
                                                    "YQgAAAvzdA2LRdCBTdQAAACAiUXQjUWYUI2FeP///1DokPL//4vQhdIPhVb+//9bjUXYUP91COjpCAAAX16L5V3CDADMVYvsgezAAAAAU1aLdRRXVuirBgAA/3UQi9iNhUD/////dQxQ6NcDAACNhXD///9Q6IsGAACL+IX/dAiBx4ABAADrDo2FQP///1DocQYAAIv4O/tzGI2FQP///1D/dQjoHAgAAF9eW4vlXcIQAI1FoFDo2vD//41F0FDo0fD//4vHK8OL2MHrBoPgP3QYUI1FoFaNBNhQ6KXy//+JRN3QiVTd1OsNjUWgVo0E2FDozgcAAItdCFPolfD//8cDAQAAAMdDBAAAAACB/4ABAAB3ElaNRaBQ6Cbx//+FwA+IggAAAI2FcP///1CNRdBQ6A7x//+FwHgWdUiNhUD///9QjUWgUOj48P//hcB/NI1FoFCNhUD///9QUOhDCwAAC8J0DlONhXD///9QUOgxCwAAjUXQUI2FcP///1BQ6CALAACLddCNRdBQweYf6HEGAACNRaBQ6GgGAAAJdcxPi3UU6WT///+NhUD///9QU+gPBwAAX15bi+VdwhAAzMzMzMzMVYvsgeyAAAAAU1aLdRRXVuh7BQAA/3UQi9iNRYD/dQxQ6LoDAACNRaBQ6GEFAACL+IX/dAiBxwABAADrC41FgFDoSgUAAIv4O/tzFY1FgFD/dQjoCAcAAF9eW4vlXcIQAI1FwFDoxu///41F4FDove///4vHK8OL2MHrBoPgP3QYUI1FwFaNBNhQ6IHy//+JRN3giVTd5OsNjUXAVo0E2FDougYAAItdCFPoge///8cDAQAAAMdDBAAAAAAPH0AAgf8AAQAAdw5WjUXAUOge" & _
                                                    "8P//hcB4c41FoFCNReBQ6A3w//+FwHgTdTyNRYBQjUXAUOj67///hcB/K41FwFCNRYBQUOjoCwAAC8J0C1ONRaBQUOjZCwAAjUXgUI1FoFBQ6MsLAACLdeCNReBQweYf6GwFAACNRcBQ6GMFAAAJddxPi3UU6Xf///+NRYBQU+gNBgAAX15bi+VdwhAAzMzMzFWL7IPsYI1FoP91EP91DFDoCwEAAI1FoFD/dQjof/L//4vlXcIMAMzMzMzMzMzMzFWL7IPsQI1FwP91EP91DFDoOwIAAI1FwFD/dQjoH/X//4vlXcIMAMzMzMzMzMzMzFWL7IPsYI1FoP91DFDozgUAAI1FoFD/dQjoIvL//4vlXcIIAMzMzMzMzMzMzMzMzFWL7IPsQI1FwP91DFDoPgcAAI1FwFD/dQjowvT//4vlXcIIAMzMzMzMzMzMzMzMzFWL7Fb/dRCLdQj/dQxW6K0IAAALwnQK/3UUVlboD+r//15dwhAAzMzMzMzMzMzMzFWL7Fb/dRCLdQj/dQxW6I0KAAALwnQK/3UUVlboH+z//15dwhAAzMzMzMzMzMzMzFWL7IPsYFMPV8BWZg8TRdiLRdxXZg8TRdAz/4td1IlF/DP2jUf7g/8GD1fAZg8TRfSLVfQPQ/A79w+H0gAAAItNEIvHDxBF0CvGDxFFwI0cwYtF+IlF8IlV+GYPH0QAAIP+Bg+DowAAAP9zBItFDP8z/3TwBP808I1FsFDo39L//4PsEIvMg+wQDxAADxAIi8QPEQEPEEXADxFN4A8RAI1FoFDoeG7//2YPc9kMDxAQZg9+yA8owmYPc9gMZg9+wQ8RVcCJTfwPEVXQO8h3E3IIi0XYO0Xocwm4AQAAADPJ6w4P" & _
                                                    "V8BmDxNF6ItN7ItF6ItV+APQi0XwiVX4E8FGg+sIiUXwO/cPhlT///+LXdTrA4tF+ItNCIt10Ik0+Yvxi8qL0IlV3Ilc/gRHi3XYi138iXXQiV3UiU3YiVX8g/8LD4Lb/v//i0UIX4lwWF6JWFxbi+VdwgwAzMzMzMzMzMxVi+yD7GBTD1fAVmYPE0XYi0XcV2YPE0XQM/+LXdSJRfwz9o1H/YP/BA9XwGYPE0X0i1X0D0PwO/cPh9IAAACLTRCLxw8QRdArxg8RRcCNHMGLRfiJRfCJVfhmDx9EAACD/gQPg6MAAAD/cwSLRQz/M/908AT/NPCNRbBQ6H/R//+D7BCLzIPsEA8QAA8QCIvEDxEBDxBFwA8RTeAPEQCNRaBQ6Bht//9mD3PZDA8QEGYPfsgPKMJmD3PYDGYPfsEPEVXAiU38DxFV0DvIdxNyCItF2DtF6HMJuAEAAAAzyesOD1fAZg8TReiLTeyLReiLVfgD0ItF8IlV+BPBRoPrCIlF8Dv3D4ZU////i13U6wOLRfiLTQiLddCJNPmL8YvKi9CJVdyJXP4ER4t12Itd/Il10Ild1IlN2IlV/IP/Bw+C2/7//4tFCF+JcDheiVg8W4vlXcIMAMzMzMzMzMzMVYvsVleLfQhX6JIAAACL8IX2dQZfXl3CBACLVPf4i8qLRPf8M/8LyHQTZg8fRAAAD6zCAUfR6IvKC8h188HmBo1GwAPHX15dwgQAzMzMzMxVi+xWV4t9CFfocgAAAIvwhfZ1Bl9eXcIEAItU9/iLyotE9/wz/wvIdBNmDx9EAAAPrMIBR9Hoi8oLyHXzweYGjUbAA8dfXl3CBADMzMzMzFWL7ItVCLgFAAAADx9EAACLDMILTMIE" & _
                                                    "dQWD6AF58kBdwgQAzMzMzMzMzMzMzMzMzFWL7ItVCLgDAAAADx9EAACLDMILTMIEdQWD6AF58kBdwgQAzMzMzMzMzMzMzMzMzFWL7IPsCItFCA9XwFOL2GYPE0X4g8AwO8N2OItN+FZXi338iU0Ii3D4g+gIi86LUAQPrNEBC00I0eoL14kIi/6JUATB5x/HRQgAAAAAO8N31V9eW4vlXcIEAMzMzMzMzFWL7IPsCItFCA9XwFOL2GYPE0X4g8AgO8N2OItN+FZXi338iU0Ii3D4g+gIi86LUAQPrNEBC00I0eoL14kIi/6JUATB5x/HRQgAAAAAO8N31V9eW4vlXcIEAMzMzMzMzFWL7ItVDItNCIsCiQGLQgSJQQSLQgiJQQiLQgyJQQyLQhCJQRCLQhSJQRSLQhiJQRiLQhyJQRyLQiCJQSCLQiSJQSSLQiiJQSiLQiyJQSxdwggAzMzMzMzMzMzMzMzMzFWL7ItVDItNCIsCiQGLQgSJQQSLQgiJQQiLQgyJQQyLQhCJQRCLQhSJQRSLQhiJQRiLQhyJQRxdwggAzMzMzMxVi+yD7GBTD1fAM8lWZg8TRdiLRdxXZg8TRdCLfdSJTeiJRfAz9o1B+4P5Bg9XwGYPE0X4i138D0PwO/EPhxkBAACLVQyLwQ8QRdArxold9A8RRcCNBMKLVfiJReyJVfyL+Sv+O/cPh+oAAAD/cAT/MItFDP908AT/NPCNRbBQ6LzN//8PEAAPEUXQO/dzQ4tN3IvBi1XUi/rB6B8BRfyLRdiD0wDB7x8PpMEBiV30M9sDwAvZC/iJXdyLRdAPpMIBiX3YA8CJVdSJRdAPEEXQ6waLXdyLfdiD7BCLxIPsEA8RAIvEDxBFwA8R" & _
                                                    "AI1FoFDoC2n//w8QCA8owWYPc9gMZg9+wA8RTcCJRfAPEU3QO8N3EHIFOX3Ycwm4AQAAADPJ6w4PV8BmDxNF4ItN5ItF4ItV/Itd9APQi0XsE9mJVfyLTehGg+gIiV30iUXsO/EPhgr///+LfdTrA4tV+It1CItF0IkEzotF2Il8zgRBi33wiVXYi9OJRdCJfdSJVfCJVdyJTeiD+QsPgpX+//+JflxfiUZYXluL5V3CCADMzFWL7IPsYFMPV8AzyVZmDxNF2ItF3FdmDxNF0It91IlN6IlF8DP2jUH9g/kED1fAZg8TRfiLXfwPQ/A78Q+HGQEAAItVDIvBDxBF0CvGiV30DxFFwI0EwotV+IlF7IlV/Iv5K/479w+H6gAAAP9wBP8wi0UM/3TwBP808I1FsFDoHMz//w8QAA8RRdA793NDi03ci8GLVdSL+sHoHwFF/ItF2IPTAMHvHw+kwQGJXfQz2wPAC9kL+Ild3ItF0A+kwgGJfdgDwIlV1IlF0A8QRdDrBotd3It92IPsEIvEg+wQDxEAi8QPEEXADxEAjUWgUOhrZ///DxAIDyjBZg9z2AxmD37ADxFNwIlF8A8RTdA7w3cQcgU5fdhzCbgBAAAAM8nrDg9XwGYPE0Xgi03ki0Xgi1X8i130A9CLRewT2YlV/ItN6EaD6AiJXfSJRew78Q+GCv///4t91OsDi1X4i3UIi0XQiQTOi0XYiXzOBEGLffCJVdiL04lF0Il91IlV8IlV3IlN6IP5Bw+Clf7//4l+PF+JRjheW4vlXcIIAMzMVYvsg+wMU4tdDA9XwFZXi30QixOL8otDBIvIZg8TRfQrNxtPBDvydQY7yHUE6xg7yHIPdwQ78nYJuAEAAAAz" & _
                                                    "0usLZg8TRfSLRfSLVfiLfQiJTwSJN4tzCIvOiXX4i3UQK04IiU0Ii0sMG04Mi10IK9iJXQgbyjtd+ItdDHUFO0sMdCM7SwxyE3cIi0MIOUUIdgm4AQAAADPS6wtmDxNF9ItV+ItF9It1CIlPDIl3CItzEIvOiXX8i3UQK04QiU0Ii0sUG04Ui10IK9iJXQgbyjtd/ItdDHUFO0sUdCM7SxRyE3cIi0MQOUUIdgm4AQAAADPS6wtmDxNF9ItV+ItF9It1CIlPFIl3EItzGIvOiXX8i3UQK04YiU0Ii0scG04ci10IK9iJXQgbyjtd/ItdDHUFO0scdCM7SxxyE3cIi0MYOUUIdgm4AQAAADPS6wtmDxNF9ItV+ItF9It1CIl3GIt1EIlPHItLICtOIIlNDItLJBtOJIt1DCvwG8o7cyB1BTtLJHQgO0skchB3BTtzIHYJuAEAAAAz0usLZg8TRfSLVfiLRfSJdyCJTySLcyiLSyyLXRCJdQiJTQwrcygbSywr8ItdDBvKO3UIdQQ7y3QsO8tyHXcFO3UIdhaJdyi4AQAAAIlPLDPSX15bi+VdwgwAZg8TRfSLVfiLRfSJdyiJTyxfXluL5V3CDADMzMzMVYvsg+wMU4tdDA9XwFZXi30QixOL8otDBIvIZg8TRfQrNxtPBDvydQY7yHUE6xg7yHIPdwQ78nYJuAEAAAAz0usLZg8TRfSLRfSLVfiLfQiJTwSLTRCJN4tzCIl1+CtxCItLDItdEBtLDCvwi10MG8o7dfh1BTtLDHQgO0sMchB3BTtzCHYJuAEAAAAz0usLZg8TRfSLVfiLRfSJTwyLTRCJdwiLcxCJdfwrcRCLSxSLXRAbSxQr8ItdDBvKO3X8dQU7" & _
                                                    "SxR0IDtLFHIQdwU7cxB2CbgBAAAAM9LrC2YPE0X0i1X4i0X0iU8UiXcQi0sYi/GLfRCLWxyJTQyLTRArcRiLyxtPHCvwi30IG8o7dQx1BDvLdCw7y3IddwU7dQx2Fol3GLgBAAAAiU8cM9JfXluL5V3CDABmDxNF9ItV+ItF9Il3GIlPHF9eW4vlXcIMAMzMzMzMzMzMzMzMzMzMzFWL7ItNCDPSVleLfQwz9ovHg+A/D6vGg/ggD0PWM/KD+EAPQ9bB7wYjNPkjVPkEi8ZfXl3CCADMzMzMzMzMzMxVi+yLVRSD7BAzyYXSD4TCAAAAU4tdEFaLdQhXi30Mg/ogD4KLAAAAjUP/A8I78HcJjUb/A8I7w3N5jUf/A8I78HcJjUb/A8I7x3Nni8KL1yvTg+DgiVX8i9Yr04lF8IlV+IvDi134i9eLffwr1olV9I1WEA8QAIt19IPBII1AII1SIA8QTAfgZg/vyA8RTAPgDxBMFuCLdQgPEEDwZg/vyA8RSuA7TfByyotVFIt9DItdEDvKcxsr+40EGSvzK9GKDDiNQAEySP+ITDD/g+oBde5fXluL5V3CEAAAAA==" ' 44077, 30.4.2020 15:27:45
Private Const STR_LIBSODIUM_SHA384_STATE As String = "2J4FwV2du8sH1Xw2KimaYhfdcDBaAVmROVkO99jsLxUxC8D/ZyYzZxEVWGiHSrSOp4/5ZA0uDNukT/q+HUi1Rw=="
'--- numeric
Private Const LNG_SHA256_HASHSZ         As Long = 32
Private Const LNG_SHA256_BLOCKSZ        As Long = 64
Private Const LNG_SHA384_HASHSZ         As Long = 48
Private Const LNG_SHA384_BLOCKSZ        As Long = 128
Private Const LNG_SHA384_CONTEXTSZ      As Long = 200
Private Const LNG_SHA512_HASHSZ         As Long = 64
Private Const LNG_HMAC_INNER_PAD        As Long = &H36
Private Const LNG_HMAC_OUTER_PAD        As Long = &H5C
Private Const LNG_FACILITY_WIN32        As Long = &H80070000
Private Const LNG_CHACHA20_KEYSZ        As Long = 32
Private Const LNG_CHACHA20POLY1305_IVSZ As Long = 12
Private Const LNG_CHACHA20POLY1305_TAGSZ As Long = 16
Private Const LNG_AES128_KEYSZ          As Long = 16
Private Const LNG_AES256_KEYSZ          As Long = 32
Private Const LNG_AESGCM_IVSZ           As Long = 12
Private Const LNG_AESGCM_TAGSZ          As Long = 16
Private Const LNG_LIBSODIUM_SHA512_CONTEXTSZ As Long = 64 + 16 + 128
'--- errors
Private Const ERR_OUT_OF_MEMORY         As Long = 8
Private Const ERR_TRUST_IS_REVOKED      As String = "Trust for this certificate or one of the certificates in the certificate chain has been revoked"
Private Const ERR_TRUST_IS_PARTIAL_CHAIN As String = "The certificate chain is not complete"
Private Const ERR_TRUST_IS_UNTRUSTED_ROOT As String = "The certificate or certificate chain is based on an untrusted root"
Private Const ERR_TRUST_IS_NOT_TIME_VALID As String = "The certificate has expired"
Private Const ERR_TRUST_REVOCATION_STATUS_UNKNOWN As String = "The revocation status of the certificate or one of the certificates in the certificate chain is unknown"
Private Const ERR_NO_MATCHING_ALT_NAME  As String = "No certificate subject name matches target host name"

Private m_uData                    As UcsCryptoThunkData

Private Enum UcsThunkPfnIndexEnum
    [_ucsPfnNotUsed]
    ucsPfnCurve25519ScalarMultiply
    ucsPfnCurve25519ScalarMultBase
    ucsPfnSecp256r1MakeKey
    ucsPfnSecp256r1SharedSecret
    ucsPfnSecp256r1UncompressKey
    ucsPfnSecp256r1Sign
    ucsPfnSecp256r1Verify
    ucsPfnSecp384r1MakeKey
    ucsPfnSecp384r1SharedSecret
    ucsPfnSecp384r1UncompressKey
    ucsPfnSecp384r1Sign
    ucsPfnSecp384r1Verify
    ucsPfnSha256Init
    ucsPfnSha256Update
    ucsPfnSha256Final
    ucsPfnSha384Init
    ucsPfnSha384Update
    ucsPfnSha384Final
    ucsPfnSha512Init
    ucsPfnSha512Update
    ucsPfnSha512Final
    ucsPfnChacha20Poly1305Encrypt
    ucsPfnChacha20Poly1305Decrypt
    ucsPfnAesGcmEncrypt
    ucsPfnAesGcmDecrypt
    ucsPfnRsaModExp
    [_ucsPfnMax]
End Enum

Private Type UcsCryptoThunkData
    Thunk               As Long
    Glob()              As Byte
    Pfn(1 To [_ucsPfnMax] - 1) As Long
    Ecc256KeySize       As Long
    Ecc384KeySize       As Long
#If ImplUseLibSodium Then
    HashCtx(0 To LNG_LIBSODIUM_SHA512_CONTEXTSZ - 1) As Byte
#Else
    HashCtx(0 To LNG_SHA384_CONTEXTSZ - 1) As Byte
#End If
    HashPad(0 To LNG_SHA384_BLOCKSZ - 1 + 1000) As Byte
    HashFinal(0 To LNG_SHA384_HASHSZ - 1 + 1000) As Byte
    hRandomProv         As Long
End Type

Public Type UcsRsaContextType
    hProv               As Long
    hPrivKey            As Long
    hPubKey             As Long
    HashAlgId           As Long
End Type

Public Enum UcsOsVersionEnum
    ucsOsvNt4 = 400
    ucsOsvWin98 = 410
    ucsOsvWin2000 = 500
    ucsOsvXp = 501
    ucsOsvVista = 600
    ucsOsvWin7 = 601
    ucsOsvWin8 = 602
    [ucsOsvWin8.1] = 603
    ucsOsvWin10 = 1000
End Enum

'=========================================================================
' Properties
'=========================================================================

Public Property Get OsVersion() As UcsOsVersionEnum
    Static lVersion     As Long
    Dim aVer(0 To 37)   As Long
    
    If lVersion = 0 Then
        aVer(0) = 4 * UBound(aVer)              '--- [0] = dwOSVersionInfoSize
        If GetVersionEx(aVer(0)) <> 0 Then
            lVersion = aVer(1) * 100 + aVer(2)  '--- [1] = dwMajorVersion, [2] = dwMinorVersion
        End If
    End If
    OsVersion = lVersion
End Property

'=========================================================================
' Functions
'=========================================================================

Public Function CryptoInit() As Boolean
    Const FUNC_NAME     As String = "CryptoInit"
    Dim lOffset         As Long
    Dim lIdx            As Long
    Dim hResult          As Long
    Dim sApiSource      As String
    
    With m_uData
        #If ImplUseLibSodium Then
            If GetModuleHandle("libsodium.dll") = 0 Then
                Call LoadLibrary(App.Path & "\libsodium.dll")
                If sodium_init() < 0 Then
                    hResult = ERR_OUT_OF_MEMORY
                    sApiSource = "sodium_init"
                    GoTo QH
                End If
            End If
        #Else
            If .hRandomProv = 0 Then
                If CryptAcquireContext(.hRandomProv, 0, 0, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) = 0 Then
                    hResult = Err.LastDllError
                    sApiSource = "CryptAcquireContext"
                    GoTo QH
                End If
            End If
        #End If
        If m_uData.Thunk = 0 Then
            .Ecc256KeySize = 32
            .Ecc384KeySize = 48
            '--- prepare thunk/context in executable memory
            .Thunk = pvThunkAllocate(STR_THUNK1 & STR_THUNK2 & STR_THUNK3 & STR_THUNK4 & STR_THUNK5)
            If .Thunk = 0 Then
                hResult = Err.LastDllError
                sApiSource = "VirtualAlloc"
                GoTo QH
            End If
            .Glob = FromBase64Array(STR_GLOB)
            '--- init pfns from thunk addr + offsets stored at beginning of it
            For lIdx = LBound(.Pfn) To UBound(.Pfn)
                Call CopyMemory(lOffset, ByVal UnsignedAdd(.Thunk, 4 * lIdx), 4)
                .Pfn(lIdx) = UnsignedAdd(.Thunk, lOffset)
            Next
            '--- init pfns trampolines
            Call pvPatchTrampoline(AddressOf pvCallSecpMakeKey)
            Call pvPatchTrampoline(AddressOf pvCallSecpSharedSecret)
            Call pvPatchTrampoline(AddressOf pvCallSecpUncompressKey)
            Call pvPatchTrampoline(AddressOf pvCallSecpSign)
            Call pvPatchTrampoline(AddressOf pvCallSecpVerify)
            Call pvPatchTrampoline(AddressOf pvCallCurve25519Multiply)
            Call pvPatchTrampoline(AddressOf pvCallCurve25519MulBase)
            Call pvPatchTrampoline(AddressOf pvCallSha2Init)
            Call pvPatchTrampoline(AddressOf pvCallSha2Update)
            Call pvPatchTrampoline(AddressOf pvCallSha2Final)
            Call pvPatchTrampoline(AddressOf pvCallChacha20Poly1305Encrypt)
            Call pvPatchTrampoline(AddressOf pvCallChacha20Poly1305Decrypt)
            Call pvPatchTrampoline(AddressOf pvCallAesGcmEncrypt)
            Call pvPatchTrampoline(AddressOf pvCallAesGcmDecrypt)
            Call pvPatchTrampoline(AddressOf pvCallRsaModExp)
            '--- init thunk's first 4 bytes -> global data in C/C++
            Call CopyMemory(ByVal .Thunk, VarPtr(.Glob(0)), 4)
            Call CopyMemory(.Glob(0), GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc"), 4)
            Call CopyMemory(.Glob(4), GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemRealloc"), 4)
            Call CopyMemory(.Glob(8), GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree"), 4)
        End If
    End With
    '--- success
    CryptoInit = True
QH:
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Public Sub CryptoTerminate()
    With m_uData
        #If Not ImplUseLibSodium Then
            If .hRandomProv <> 0 Then
                Call CryptReleaseContext(.hRandomProv, 0)
                .hRandomProv = 0
            End If
        #End If
    End With
End Sub

Public Function CryptoIsSupported(ByVal eAead As UcsTlsCryptoAlgorithmsEnum) As Boolean
    Const PREF          As Long = &H1000
    
    Select Case eAead
    Case ucsTlsAlgoAeadAes128, ucsTlsAlgoAeadAes256
        #If ImplUseLibSodium Then
            CryptoIsSupported = (crypto_aead_aes256gcm_is_available() <> 0 And eAead = ucsTlsAlgoAeadAes256)
        #Else
            CryptoIsSupported = True
        #End If
    Case PREF + ucsTlsAlgoAeadAes128, PREF + ucsTlsAlgoAeadAes256
        '--- signal if AES preferred over Chacha20
        #If ImplUseLibSodium Then
            CryptoIsSupported = (crypto_aead_aes256gcm_is_available() <> 0 And eAead = PREF + ucsTlsAlgoAeadAes256)
        #End If
    Case ucsTlsAlgoSignaturePss
        #If ImplUseBCrypt Then
            CryptoIsSupported = (OsVersion >= ucsOsvVista)  '--- need BCrypt for PSS padding on signatures
        #Else
            CryptoIsSupported = True
        #End If
    Case ucsTlsAlgoSignaturePkcsSha2
        CryptoIsSupported = (OsVersion >= ucsOsvXp)         '--- need PROV_RSA_AES for SHA-2
    Case Else
        CryptoIsSupported = True
    End Select
End Function

Public Function CryptoEccCurve25519MakeKey(baPrivate() As Byte, baPublic() As Byte) As Boolean
    ReDim baPrivate(0 To m_uData.Ecc256KeySize - 1) As Byte
    Debug.Assert RedimStats("CryptoEccCurve25519MakeKey.baPrivate", UBound(baPrivate) + 1)
    ReDim baPublic(0 To m_uData.Ecc256KeySize - 1) As Byte
    Debug.Assert RedimStats("CryptoEccCurve25519MakeKey.baPublic", UBound(baPublic) + 1)
    CryptoRandomBytes VarPtr(baPrivate(0)), m_uData.Ecc256KeySize
    '--- fix issues w/ specific privkeys
    baPrivate(0) = baPrivate(0) And 248
    baPrivate(UBound(baPrivate)) = (baPrivate(UBound(baPrivate)) And 127) Or 64
    #If ImplUseLibSodium Then
        Call crypto_scalarmult_curve25519_base(baPublic(0), baPrivate(0))
    #Else
        Debug.Assert pvPatchTrampoline(AddressOf pvCallCurve25519MulBase)
        pvCallCurve25519MulBase m_uData.Pfn(ucsPfnCurve25519ScalarMultBase), baPublic(0), baPrivate(0)
    #End If
    '--- success
    CryptoEccCurve25519MakeKey = True
End Function

Public Function CryptoEccCurve25519SharedSecret(baPrivate() As Byte, baPublic() As Byte) As Byte()
    Dim baRetVal()      As Byte
    
    Debug.Assert UBound(baPrivate) >= m_uData.Ecc256KeySize - 1
    Debug.Assert UBound(baPublic) >= m_uData.Ecc256KeySize - 1
    ReDim baRetVal(0 To m_uData.Ecc256KeySize - 1) As Byte
    Debug.Assert RedimStats("CryptoEccCurve25519SharedSecret.baRetVal", UBound(baRetVal) + 1)
    #If ImplUseLibSodium Then
        Call crypto_scalarmult_curve25519(baRetVal(0), baPrivate(0), baPublic(0))
    #Else
        Debug.Assert pvPatchTrampoline(AddressOf pvCallCurve25519Multiply)
        pvCallCurve25519Multiply m_uData.Pfn(ucsPfnCurve25519ScalarMultiply), baRetVal(0), baPrivate(0), baPublic(0)
    #End If
    CryptoEccCurve25519SharedSecret = baRetVal
End Function

Public Function CryptoEccSecp256r1MakeKey(baPrivate() As Byte, baPublic() As Byte) As Boolean
    Const MAX_RETRIES   As Long = 16
    Dim lIdx            As Long
    
    ReDim baPrivate(0 To m_uData.Ecc256KeySize - 1) As Byte
    Debug.Assert RedimStats("CryptoEccSecp256r1MakeKey.baPrivate", UBound(baPrivate) + 1)
    ReDim baPublic(0 To m_uData.Ecc256KeySize) As Byte
    Debug.Assert RedimStats("CryptoEccSecp256r1MakeKey.baPublic", UBound(baPublic) + 1)
    For lIdx = 1 To MAX_RETRIES
        CryptoRandomBytes VarPtr(baPrivate(0)), m_uData.Ecc256KeySize
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpMakeKey)
        If pvCallSecpMakeKey(m_uData.Pfn(ucsPfnSecp256r1MakeKey), baPublic(0), baPrivate(0)) = 1 Then
            Exit For
        End If
    Next
    If lIdx <= MAX_RETRIES Then
        baPublic = CryptoEccSecp256r1UncompressKey(baPublic)
        '--- success
        CryptoEccSecp256r1MakeKey = True
    End If
End Function

Public Function CryptoEccSecp256r1SharedSecret(baPrivate() As Byte, baPublic() As Byte) As Byte()
    Dim baRetVal()      As Byte
    
    Debug.Assert UBound(baPrivate) >= m_uData.Ecc256KeySize - 1
    Debug.Assert UBound(baPublic) >= m_uData.Ecc256KeySize
    ReDim baRetVal(0 To m_uData.Ecc256KeySize - 1) As Byte
    Debug.Assert RedimStats("CryptoEccSecp256r1SharedSecret.baRetVal", UBound(baRetVal) + 1)
    Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpSharedSecret)
    If pvCallSecpSharedSecret(m_uData.Pfn(ucsPfnSecp256r1SharedSecret), baPublic(0), baPrivate(0), baRetVal(0)) = 0 Then
        GoTo QH
    End If
    CryptoEccSecp256r1SharedSecret = baRetVal
QH:
End Function

Public Function CryptoEccSecp256r1UncompressKey(baPublic() As Byte) As Byte()
    Dim baRetVal()      As Byte
    
    ReDim baRetVal(0 To 2 * m_uData.Ecc256KeySize) As Byte
    Debug.Assert RedimStats("CryptoEccSecp256r1UncompressKey.baRetVal", UBound(baRetVal) + 1)
    Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpUncompressKey)
    If pvCallSecpUncompressKey(m_uData.Pfn(ucsPfnSecp256r1UncompressKey), baPublic(0), baRetVal(0)) = 0 Then
        GoTo QH
    End If
    CryptoEccSecp256r1UncompressKey = baRetVal
QH:
End Function

Public Function CryptoEccSecp256r1Sign(baPrivKey() As Byte, baHash() As Byte) As Byte()
    Const MAX_RETRIES   As Long = 16
    Dim baPrivate()     As Byte
    Dim baRandom()      As Byte
    Dim baRetVal()      As Byte
    Dim lIdx            As Long
    
    baPrivate = Asn1DecodePrivateKeyFromDer(baPrivKey)
    ReDim baRandom(0 To m_uData.Ecc256KeySize - 1) As Byte
    Debug.Assert RedimStats("CryptoEccSecp256r1Sign.baRandom", UBound(baRandom) + 1)
    ReDim baRetVal(0 To 2 * m_uData.Ecc256KeySize - 1) As Byte
    Debug.Assert RedimStats("CryptoEccSecp256r1Sign.baRetVal", UBound(baRetVal) + 1)
    For lIdx = 1 To MAX_RETRIES
        CryptoRandomBytes VarPtr(baRandom(0)), m_uData.Ecc256KeySize
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpSign)
        If pvCallSecpSign(m_uData.Pfn(ucsPfnSecp256r1Sign), baPrivate(0), baHash(0), baRandom(0), baRetVal(0)) <> 0 Then
            Exit For
        End If
    Next
    If lIdx < MAX_RETRIES Then
        '--- success
        CryptoEccSecp256r1Sign = baRetVal
    End If
End Function

Public Function CryptoEccSecp256r1Verify(baPublic() As Byte, baHash() As Byte, baSignature() As Byte) As Boolean
    Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpVerify)
    CryptoEccSecp256r1Verify = (pvCallSecpVerify(m_uData.Pfn(ucsPfnSecp256r1Verify), baPublic(0), baHash(0), baSignature(0)) <> 0)
End Function

Public Function CryptoEccSecp384r1MakeKey(baPrivate() As Byte, baPublic() As Byte) As Boolean
    Const MAX_RETRIES   As Long = 16
    Dim lIdx            As Long
        
    ReDim baPrivate(0 To m_uData.Ecc384KeySize - 1) As Byte
    Debug.Assert RedimStats("CryptoEccSecp384r1MakeKey.baPrivate", UBound(baPrivate) + 1)
    ReDim baPublic(0 To m_uData.Ecc384KeySize) As Byte
    Debug.Assert RedimStats("CryptoEccSecp384r1MakeKey.baPublic", UBound(baPublic) + 1)
    For lIdx = 1 To MAX_RETRIES
        CryptoRandomBytes VarPtr(baPrivate(0)), m_uData.Ecc384KeySize
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpMakeKey)
        If pvCallSecpMakeKey(m_uData.Pfn(ucsPfnSecp384r1MakeKey), baPublic(0), baPrivate(0)) = 1 Then
            Exit For
        End If
    Next
    If lIdx <= MAX_RETRIES Then
        baPublic = CryptoEccSecp384r1UncompressKey(baPublic)
        '--- success
        CryptoEccSecp384r1MakeKey = True
    End If
End Function

Public Function CryptoEccSecp384r1SharedSecret(baPrivate() As Byte, baPublic() As Byte) As Byte()
    Dim baRetVal()      As Byte
    
    Debug.Assert UBound(baPrivate) >= m_uData.Ecc384KeySize - 1
    Debug.Assert UBound(baPublic) >= m_uData.Ecc384KeySize
    ReDim baRetVal(0 To m_uData.Ecc384KeySize - 1) As Byte
    Debug.Assert RedimStats("CryptoEccSecp384r1SharedSecret.baRetVal", UBound(baRetVal) + 1)
    Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpSharedSecret)
    If pvCallSecpSharedSecret(m_uData.Pfn(ucsPfnSecp384r1SharedSecret), baPublic(0), baPrivate(0), baRetVal(0)) = 0 Then
        GoTo QH
    End If
    CryptoEccSecp384r1SharedSecret = baRetVal
QH:
End Function

Public Function CryptoEccSecp384r1UncompressKey(baPublic() As Byte) As Byte()
    Dim baRetVal()      As Byte
    
    ReDim baRetVal(0 To 2 * m_uData.Ecc384KeySize) As Byte
    Debug.Assert RedimStats("CryptoEccSecp384r1UncompressKey.baRetVal", UBound(baRetVal) + 1)
    Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpUncompressKey)
    If pvCallSecpUncompressKey(m_uData.Pfn(ucsPfnSecp384r1UncompressKey), baPublic(0), baRetVal(0)) = 0 Then
        GoTo QH
    End If
    CryptoEccSecp384r1UncompressKey = baRetVal
QH:
End Function

Public Function CryptoEccSecp384r1Sign(baPrivKey() As Byte, baHash() As Byte) As Byte()
    Const MAX_RETRIES   As Long = 16
    Dim baPrivate()     As Byte
    Dim baRandom()      As Byte
    Dim baRetVal()      As Byte
    Dim lIdx            As Long
    
    baPrivate = Asn1DecodePrivateKeyFromDer(baPrivKey)
    ReDim baRandom(0 To m_uData.Ecc384KeySize - 1) As Byte
    Debug.Assert RedimStats("CryptoEccSecp384r1Sign.baRandom", UBound(baRandom) + 1)
    ReDim baRetVal(0 To 2 * m_uData.Ecc384KeySize - 1) As Byte
    Debug.Assert RedimStats("CryptoEccSecp384r1Sign.baRetVal", UBound(baRetVal) + 1)
    For lIdx = 1 To MAX_RETRIES
        CryptoRandomBytes VarPtr(baRandom(0)), m_uData.Ecc384KeySize
        Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpSign)
        If pvCallSecpSign(m_uData.Pfn(ucsPfnSecp384r1Sign), baPrivate(0), baHash(0), baRandom(0), baRetVal(0)) <> 0 Then
            Exit For
        End If
    Next
    If lIdx < MAX_RETRIES Then
        '--- success
        CryptoEccSecp384r1Sign = baRetVal
    End If
End Function

Public Function CryptoEccSecp384r1Verify(baPublic() As Byte, baHash() As Byte, baSignature() As Byte) As Boolean
    Debug.Assert pvPatchTrampoline(AddressOf pvCallSecpVerify)
    CryptoEccSecp384r1Verify = (pvCallSecpVerify(m_uData.Pfn(ucsPfnSecp384r1Verify), baPublic(0), baHash(0), baSignature(0)) <> 0)
End Function

Public Function CryptoHashSha256(baInput() As Byte, ByVal lPos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim lCtxPtr         As Long
    Dim lPtr            As Long
    Dim baRetVal()      As Byte
    
    If Size < 0 Then
        Size = pvArraySize(baInput) - lPos
    Else
        Debug.Assert pvArraySize(baInput) >= lPos + Size
    End If
    If Size > 0 Then
        lPtr = VarPtr(baInput(lPos))
    End If
    ReDim baRetVal(0 To LNG_SHA256_HASHSZ - 1) As Byte
    Debug.Assert RedimStats("CryptoHashSha256.baRetVal", UBound(baRetVal) + 1)
    #If ImplUseLibSodium Then
        Call crypto_hash_sha256(baRetVal(0), ByVal lPtr, Size)
    #Else
        With m_uData
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Init)
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Update)
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Final)
            lCtxPtr = VarPtr(.HashCtx(0))
            pvCallSha2Init .Pfn(ucsPfnSha256Init), lCtxPtr
            pvCallSha2Update .Pfn(ucsPfnSha256Update), lCtxPtr, lPtr, Size
            pvCallSha2Final .Pfn(ucsPfnSha256Final), lCtxPtr, baRetVal(0)
        End With
    #End If
    CryptoHashSha256 = baRetVal
End Function

Public Function CryptoHashSha384(baInput() As Byte, ByVal lPos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim lCtxPtr         As Long
    Dim lPtr            As Long
    Dim baRetVal()      As Byte
    
    If Size < 0 Then
        Size = pvArraySize(baInput) - lPos
    Else
        Debug.Assert pvArraySize(baInput) >= lPos + Size
    End If
    If Size > 0 Then
        lPtr = VarPtr(baInput(lPos))
    End If
    ReDim baRetVal(0 To LNG_SHA384_HASHSZ - 1) As Byte
    Debug.Assert RedimStats("CryptoHashSha384.baRetVal", UBound(baRetVal) + 1)
    With m_uData
        lCtxPtr = VarPtr(.HashCtx(0))
        #If ImplUseLibSodium Then
            Call crypto_hash_sha384_init(.HashCtx)
            Call crypto_hash_sha512_update(ByVal lCtxPtr, ByVal lPtr, Size)
            Call crypto_hash_sha512_final(ByVal lCtxPtr, .HashFinal(0))
            Call CopyMemory(baRetVal(0), .HashFinal(0), LNG_SHA384_HASHSZ)
        #Else
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Init)
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Update)
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Final)
            pvCallSha2Init .Pfn(ucsPfnSha384Init), lCtxPtr
            pvCallSha2Update .Pfn(ucsPfnSha384Update), lCtxPtr, lPtr, Size
            pvCallSha2Final .Pfn(ucsPfnSha384Final), lCtxPtr, baRetVal(0)
        #End If
    End With
    CryptoHashSha384 = baRetVal
End Function

Public Function CryptoHashSha512(baInput() As Byte, ByVal lPos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim lCtxPtr         As Long
    Dim lPtr            As Long
    Dim baRetVal()      As Byte
    
    If Size < 0 Then
        Size = pvArraySize(baInput) - lPos
    Else
        Debug.Assert pvArraySize(baInput) >= lPos + Size
    End If
    If Size > 0 Then
        lPtr = VarPtr(baInput(lPos))
    End If
    ReDim baRetVal(0 To LNG_SHA512_HASHSZ - 1) As Byte
    Debug.Assert RedimStats("CryptoHashSha512.baRetVal", UBound(baRetVal) + 1)
    With m_uData
        lCtxPtr = VarPtr(.HashCtx(0))
        #If ImplUseLibSodium Then
            Call crypto_hash_sha512_init(ByVal lCtxPtr)
            Call crypto_hash_sha512_update(ByVal lCtxPtr, ByVal lPtr, Size)
            Call crypto_hash_sha512_final(ByVal lCtxPtr, .HashFinal(0))
            Call CopyMemory(baRetVal(0), .HashFinal(0), LNG_SHA512_HASHSZ)
        #Else
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Init)
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Update)
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Final)
            pvCallSha2Init .Pfn(ucsPfnSha512Init), lCtxPtr
            pvCallSha2Update .Pfn(ucsPfnSha512Update), lCtxPtr, lPtr, Size
            pvCallSha2Final .Pfn(ucsPfnSha512Final), lCtxPtr, baRetVal(0)
        #End If
    End With
    CryptoHashSha512 = baRetVal
End Function

Public Function CryptoHmacSha256(baKey() As Byte, baInput() As Byte, ByVal lPos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim lCtxPtr         As Long
    Dim lPtr            As Long
    Dim lIdx            As Long
    
    Debug.Assert UBound(baKey) < LNG_SHA256_BLOCKSZ
    If Size < 0 Then
        Size = pvArraySize(baInput) - lPos
    Else
        Debug.Assert pvArraySize(baInput) >= lPos + Size
    End If
    If Size > 0 Then
        lPtr = VarPtr(baInput(lPos))
    End If
    With m_uData
        lCtxPtr = VarPtr(.HashCtx(0))
        ReDim baRetVal(0 To LNG_SHA256_HASHSZ - 1) As Byte
        Debug.Assert RedimStats("CryptoHmacSha256.baRetVal", UBound(baRetVal) + 1)
        #If ImplUseLibSodium Then
            '-- inner hash
            Call crypto_hash_sha256_init(ByVal lCtxPtr)
            Call FillMemory(.HashPad(0), LNG_SHA256_BLOCKSZ, LNG_HMAC_INNER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_INNER_PAD
            Next
            Call crypto_hash_sha256_update(ByVal lCtxPtr, .HashPad(0), LNG_SHA256_BLOCKSZ)
            Call crypto_hash_sha256_update(ByVal lCtxPtr, ByVal lPtr, Size)
            Call crypto_hash_sha256_final(ByVal lCtxPtr, .HashFinal(0))
            '-- outer hash
            Call crypto_hash_sha256_init(ByVal lCtxPtr)
            Call FillMemory(.HashPad(0), LNG_SHA256_BLOCKSZ, LNG_HMAC_OUTER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_OUTER_PAD
            Next
            Call crypto_hash_sha256_update(ByVal lCtxPtr, .HashPad(0), LNG_SHA256_BLOCKSZ)
            Call crypto_hash_sha256_update(ByVal lCtxPtr, .HashFinal(0), LNG_SHA256_HASHSZ)
            Call crypto_hash_sha256_final(ByVal lCtxPtr, baRetVal(0))
        #Else
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Init)
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Update)
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Final)
            '-- inner hash
            pvCallSha2Init .Pfn(ucsPfnSha256Init), lCtxPtr
            Call FillMemory(.HashPad(0), LNG_SHA256_BLOCKSZ, LNG_HMAC_INNER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_INNER_PAD
            Next
            pvCallSha2Update .Pfn(ucsPfnSha256Update), lCtxPtr, VarPtr(.HashPad(0)), LNG_SHA256_BLOCKSZ
            pvCallSha2Update .Pfn(ucsPfnSha256Update), lCtxPtr, lPtr, Size
            pvCallSha2Final .Pfn(ucsPfnSha256Final), lCtxPtr, .HashFinal(0)
            '-- outer hash
            pvCallSha2Init .Pfn(ucsPfnSha256Init), lCtxPtr
            Call FillMemory(.HashPad(0), LNG_SHA256_BLOCKSZ, LNG_HMAC_OUTER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_OUTER_PAD
            Next
            pvCallSha2Update .Pfn(ucsPfnSha256Update), lCtxPtr, VarPtr(.HashPad(0)), LNG_SHA256_BLOCKSZ
            pvCallSha2Update .Pfn(ucsPfnSha256Update), lCtxPtr, VarPtr(.HashFinal(0)), LNG_SHA256_HASHSZ
            pvCallSha2Final .Pfn(ucsPfnSha256Final), lCtxPtr, baRetVal(0)
        #End If
    End With
    CryptoHmacSha256 = baRetVal
End Function

Public Function CryptoHmacSha384(baKey() As Byte, baInput() As Byte, ByVal lPos As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim lCtxPtr         As Long
    Dim lPtr            As Long
    Dim lIdx            As Long
    
    Debug.Assert UBound(baKey) < LNG_SHA384_BLOCKSZ
    If Size < 0 Then
        Size = pvArraySize(baInput) - lPos
    Else
        Debug.Assert pvArraySize(baInput) >= lPos + Size
    End If
    If Size > 0 Then
        lPtr = VarPtr(baInput(lPos))
    End If
    With m_uData
        lCtxPtr = VarPtr(.HashCtx(0))
        ReDim baRetVal(0 To LNG_SHA384_HASHSZ - 1) As Byte
        Debug.Assert RedimStats("CryptoHmacSha384.baRetVal", UBound(baRetVal) + 1)
        #If ImplUseLibSodium Then
            '-- inner hash
            Call crypto_hash_sha384_init(.HashCtx)
            Call FillMemory(.HashPad(0), LNG_SHA384_BLOCKSZ, LNG_HMAC_INNER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_INNER_PAD
            Next
            Call crypto_hash_sha512_update(ByVal lCtxPtr, .HashPad(0), LNG_SHA384_BLOCKSZ)
            Call crypto_hash_sha512_update(ByVal lCtxPtr, ByVal lPtr, Size)
            Call crypto_hash_sha512_final(ByVal lCtxPtr, .HashFinal(0))
            '-- outer hash
            Call crypto_hash_sha384_init(.HashCtx)
            Call FillMemory(.HashPad(0), LNG_SHA384_BLOCKSZ, LNG_HMAC_OUTER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_OUTER_PAD
            Next
            Call crypto_hash_sha512_update(ByVal lCtxPtr, .HashPad(0), LNG_SHA384_BLOCKSZ)
            Call crypto_hash_sha512_update(ByVal lCtxPtr, .HashFinal(0), LNG_SHA384_HASHSZ)
            Call crypto_hash_sha512_final(ByVal lCtxPtr, .HashFinal(0))
            Call CopyMemory(baRetVal(0), .HashFinal(0), LNG_SHA384_HASHSZ)
        #Else
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Init)
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Update)
            Debug.Assert pvPatchTrampoline(AddressOf pvCallSha2Final)
            '-- inner hash
            pvCallSha2Init .Pfn(ucsPfnSha384Init), lCtxPtr
            Call FillMemory(.HashPad(0), LNG_SHA384_BLOCKSZ, LNG_HMAC_INNER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_INNER_PAD
            Next
            pvCallSha2Update .Pfn(ucsPfnSha384Update), lCtxPtr, VarPtr(.HashPad(0)), LNG_SHA384_BLOCKSZ
            pvCallSha2Update .Pfn(ucsPfnSha384Update), lCtxPtr, lPtr, Size
            pvCallSha2Final .Pfn(ucsPfnSha384Final), lCtxPtr, .HashFinal(0)
            '-- outer hash
            pvCallSha2Init .Pfn(ucsPfnSha384Init), lCtxPtr
            Call FillMemory(.HashPad(0), LNG_SHA384_BLOCKSZ, LNG_HMAC_OUTER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_OUTER_PAD
            Next
            pvCallSha2Update .Pfn(ucsPfnSha384Update), lCtxPtr, VarPtr(.HashPad(0)), LNG_SHA384_BLOCKSZ
            pvCallSha2Update .Pfn(ucsPfnSha384Update), lCtxPtr, VarPtr(.HashFinal(0)), LNG_SHA384_HASHSZ
            pvCallSha2Final .Pfn(ucsPfnSha384Final), lCtxPtr, baRetVal(0)
        #End If
    End With
    CryptoHmacSha384 = baRetVal
End Function

Public Function CryptoAeadChacha20Poly1305Encrypt( _
            baNonce() As Byte, baKey() As Byte, _
            baAad() As Byte, ByVal lAadPos As Long, ByVal lAdSize As Long, _
            baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Dim lAdPtr          As Long
    
    Debug.Assert pvArraySize(baNonce) = LNG_CHACHA20POLY1305_IVSZ
    Debug.Assert pvArraySize(baKey) = LNG_CHACHA20_KEYSZ
    Debug.Assert pvArraySize(baBuffer) >= lPos + lSize + LNG_CHACHA20POLY1305_TAGSZ
    If lSize > 0 Then
        If lAdSize > 0 Then
            lAdPtr = VarPtr(baAad(lAadPos))
        End If
        #If ImplUseLibSodium Then
            Call crypto_aead_chacha20poly1305_ietf_encrypt(baBuffer(lPos), ByVal 0, baBuffer(lPos), lSize, 0, ByVal lAdPtr, lAdSize, 0, 0, baNonce(0), baKey(0))
        #Else
            Debug.Assert pvPatchTrampoline(AddressOf pvCallChacha20Poly1305Encrypt)
            Call pvCallChacha20Poly1305Encrypt(m_uData.Pfn(ucsPfnChacha20Poly1305Encrypt), _
                    baKey(0), baNonce(0), _
                    lAdPtr, lAdSize, _
                    baBuffer(lPos), lSize, _
                    baBuffer(lPos), baBuffer(lPos + lSize))
        #End If
    End If
    '--- success
    CryptoAeadChacha20Poly1305Encrypt = True
End Function

Public Function CryptoAeadChacha20Poly1305Decrypt( _
            baNonce() As Byte, baKey() As Byte, _
            baAad() As Byte, ByVal lAadPos As Long, ByVal lAdSize As Long, _
            baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Debug.Assert pvArraySize(baNonce) = LNG_CHACHA20POLY1305_IVSZ
    Debug.Assert pvArraySize(baKey) = LNG_CHACHA20_KEYSZ
    Debug.Assert pvArraySize(baBuffer) >= lPos + lSize
    #If ImplUseLibSodium Then
        If crypto_aead_chacha20poly1305_ietf_decrypt(baBuffer(lPos), ByVal 0, 0, baBuffer(lPos), lSize, 0, baAad(lAadPos), lAdSize, 0, baNonce(0), baKey(0)) = 0 Then
            '--- success
            CryptoAeadChacha20Poly1305Decrypt = True
        End If
    #Else
        Debug.Assert pvPatchTrampoline(AddressOf pvCallChacha20Poly1305Decrypt)
        If pvCallChacha20Poly1305Decrypt(m_uData.Pfn(ucsPfnChacha20Poly1305Decrypt), _
                baKey(0), baNonce(0), _
                baAad(lAadPos), lAdSize, _
                baBuffer(lPos), lSize - LNG_CHACHA20POLY1305_TAGSZ, _
                baBuffer(lPos + lSize - LNG_CHACHA20POLY1305_TAGSZ), baBuffer(lPos)) = 0 Then
            '--- success
            CryptoAeadChacha20Poly1305Decrypt = True
        End If
    #End If
End Function

Public Function CryptoAeadAesGcmEncrypt( _
            baNonce() As Byte, baKey() As Byte, _
            baAad() As Byte, ByVal lAadPos As Long, ByVal lAdSize As Long, _
            baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Dim lAdPtr          As Long
    
    Debug.Assert pvArraySize(baNonce) = LNG_AESGCM_IVSZ
    #If ImplUseLibSodium Then
        Debug.Assert pvArraySize(baKey) = LNG_AES256_KEYSZ
    #Else
        Debug.Assert pvArraySize(baKey) = LNG_AES128_KEYSZ Or pvArraySize(baKey) = LNG_AES256_KEYSZ
    #End If
    Debug.Assert pvArraySize(baBuffer) >= lPos + lSize + LNG_AESGCM_TAGSZ
    If lSize > 0 Then
        If lAdSize > 0 Then
            lAdPtr = VarPtr(baAad(lAadPos))
        End If
        #If ImplUseLibSodium Then
            Call crypto_aead_aes256gcm_encrypt(baBuffer(lPos), ByVal 0, baBuffer(lPos), lSize, 0, ByVal lAdPtr, lAdSize, 0, 0, baNonce(0), baKey(0))
        #Else
            Debug.Assert pvPatchTrampoline(AddressOf pvCallAesGcmEncrypt)
            Call pvCallAesGcmEncrypt(m_uData.Pfn(ucsPfnAesGcmEncrypt), _
                    baBuffer(lPos), baBuffer(lPos + lSize), _
                    baBuffer(lPos), lSize, _
                    lAdPtr, lAdSize, _
                    baNonce(0), baKey(0), UBound(baKey) + 1)
        #End If
    End If
    '--- success
    CryptoAeadAesGcmEncrypt = True
End Function

Public Function CryptoAeadAesGcmDecrypt( _
            baNonce() As Byte, baKey() As Byte, _
            baAad() As Byte, ByVal lAadPos As Long, ByVal lAdSize As Long, _
            baBuffer() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Boolean
    Debug.Assert pvArraySize(baNonce) = LNG_AESGCM_IVSZ
    #If ImplUseLibSodium Then
        Debug.Assert pvArraySize(baKey) = LNG_AES256_KEYSZ
    #Else
        Debug.Assert pvArraySize(baKey) = LNG_AES128_KEYSZ Or pvArraySize(baKey) = LNG_AES256_KEYSZ
    #End If
    Debug.Assert pvArraySize(baBuffer) >= lPos + lSize
    #If ImplUseLibSodium Then
        If crypto_aead_aes256gcm_decrypt(baBuffer(lPos), ByVal 0, 0, baBuffer(lPos), lSize, 0, baAad(lAadPos), lAdSize, 0, baNonce(0), baKey(0)) = 0 Then
            '--- success
            CryptoAeadAesGcmDecrypt = True
        End If
    #Else
        Debug.Assert pvPatchTrampoline(AddressOf pvCallAesGcmDecrypt)
        If pvCallAesGcmDecrypt(m_uData.Pfn(ucsPfnAesGcmDecrypt), _
                baBuffer(lPos), _
                baBuffer(lPos), lSize - LNG_AESGCM_TAGSZ, _
                baBuffer(lPos + lSize - LNG_AESGCM_TAGSZ), _
                baAad(lAadPos), lAdSize, _
                baNonce(0), baKey(0), UBound(baKey) + 1) = 0 Then
            '--- success
            CryptoAeadAesGcmDecrypt = True
        End If
    #End If
End Function

Public Sub CryptoRandomBytes(ByVal lPtr As Long, ByVal lSize As Long)
    #If ImplUseLibSodium Then
        Call randombytes_buf(lPtr, lSize)
    #Else
        Call CryptGenRandom(m_uData.hRandomProv, lSize, lPtr)
    #End If
End Sub

Public Function CryptoRsaModExp(baBase() As Byte, baExp() As Byte, baModulo() As Byte) As Byte()
    Dim baRetVal()      As Byte

    ReDim baRetVal(0 To UBound(baBase)) As Byte
    Call pvCallRsaModExp(m_uData.Pfn(ucsPfnRsaModExp), UBound(baBase) + 1, baBase(0), baExp(0), baModulo(0), baRetVal(0))
    CryptoRsaModExp = baRetVal
End Function

'= RSA helpers ===========================================================

Public Function CryptoRsaInitContext(uCtx As UcsRsaContextType, baPrivKey() As Byte, baCert() As Byte, baPubKey() As Byte, Optional ByVal SignatureType As Long) As Boolean
    Const FUNC_NAME     As String = "CryptoRsaInitContext"
    Dim lHashAlgId      As Long
    Dim hProv           As Long
    Dim lPkiPtr         As Long
    Dim lKeyPtr         As Long
    Dim lKeySize        As Long
    Dim uKeyBlob        As CRYPT_BLOB_DATA
    Dim hPrivKey        As Long
    Dim pCertContext    As Long
    Dim lPtr            As Long
    Dim hPubKey         As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    
    Select Case SignatureType \ &H100
    Case 0
        '--- no hash
    Case 1
        lHashAlgId = CALG_MD5
    Case 2
        lHashAlgId = CALG_SHA1
    Case 4
        lHashAlgId = CALG_SHA_256
    Case 5
        lHashAlgId = CALG_SHA_384
    Case 6
        lHashAlgId = CALG_SHA_512
    Case Else
        GoTo QH
    End Select
    If CryptAcquireContext(hProv, 0, 0, IIf(lHashAlgId >= CALG_SHA_256, PROV_RSA_AES, PROV_RSA_FULL), CRYPT_VERIFYCONTEXT) = 0 Then
        hResult = Err.LastDllError
        '-- no PROV_RSA_AES on Win2000 and below
        If hResult <> NTE_PROV_TYPE_NOT_DEF Then
            sApiSource = "CryptAcquireContext"
        End If
        GoTo QH
    End If
    If pvArraySize(baPrivKey) > 0 Then
        If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_PRIVATE_KEY_INFO, baPrivKey(0), UBound(baPrivKey) + 1, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lPkiPtr, 0) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptDecodeObjectEx(PKCS_PRIVATE_KEY_INFO)"
            GoTo QH
        End If
        Call CopyMemory(uKeyBlob, ByVal UnsignedAdd(lPkiPtr, 16), Len(uKeyBlob)) '--- dereference PCRYPT_PRIVATE_KEY_INFO->PrivateKey
        If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_RSA_PRIVATE_KEY, ByVal uKeyBlob.pbData, uKeyBlob.cbData, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lKeyPtr, lKeySize) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptDecodeObjectEx(PKCS_RSA_PRIVATE_KEY)"
            GoTo QH
        End If
        If CryptImportKey(hProv, ByVal lKeyPtr, lKeySize, 0, 0, hPrivKey) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptImportKey"
            GoTo QH
        End If
    End If
    If pvArraySize(baCert) > 0 Then
        pCertContext = CertCreateCertificateContext(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, baCert(0), UBound(baCert) + 1)
        If pCertContext = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CertCreateCertificateContext"
            GoTo QH
        End If
        Call CopyMemory(lPtr, ByVal UnsignedAdd(pCertContext, 12), 4)       '--- dereference pCertContext->pCertInfo
        lPtr = UnsignedAdd(lPtr, 56)                                    '--- &pCertContext->pCertInfo->SubjectPublicKeyInfo
        If CryptImportPublicKeyInfo(hProv, X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, ByVal lPtr, hPubKey) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptImportPublicKeyInfo#1"
            GoTo QH
        End If
    ElseIf pvArraySize(baPubKey) > 0 Then
        If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, X509_PUBLIC_KEY_INFO, baPubKey(0), UBound(baPubKey) + 1, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lKeyPtr, 0) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptDecodeObjectEx(X509_PUBLIC_KEY_INFO)"
            GoTo QH
        End If
        If CryptImportPublicKeyInfo(hProv, X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, ByVal lKeyPtr, hPubKey) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptImportPublicKeyInfo#2"
            GoTo QH
        End If
    End If
    '--- commit
    uCtx.hProv = hProv: hProv = 0
    uCtx.hPrivKey = hPrivKey: hPrivKey = 0
    uCtx.hPubKey = hPubKey: hPubKey = 0
    uCtx.HashAlgId = lHashAlgId
    '--- success
    CryptoRsaInitContext = True
QH:
    If hPrivKey <> 0 Then
        Call CryptDestroyKey(hPrivKey)
    End If
    If hPubKey <> 0 Then
        Call CryptDestroyKey(hPubKey)
    End If
    If pCertContext <> 0 Then
        Call CertFreeCertificateContext(pCertContext)
    End If
    If hProv <> 0 Then
        Call CryptReleaseContext(hProv, 0)
    End If
    If lPkiPtr <> 0 Then
        Call LocalFree(lPkiPtr)
    End If
    If lKeyPtr <> 0 Then
        Call LocalFree(lKeyPtr)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Public Sub CryptoRsaTerminateContext(uCtx As UcsRsaContextType)
    If uCtx.hPrivKey <> 0 Then
        Call CryptDestroyKey(uCtx.hPrivKey)
        uCtx.hPrivKey = 0
    End If
    If uCtx.hPubKey <> 0 Then
        Call CryptDestroyKey(uCtx.hPubKey)
        uCtx.hPubKey = 0
    End If
    If uCtx.hProv <> 0 Then
        Call CryptReleaseContext(uCtx.hProv, 0)
        uCtx.hProv = 0
    End If
End Sub

Public Function CryptoRsaSign(uCtx As UcsRsaContextType, baMessage() As Byte) As Byte()
    Const FUNC_NAME     As String = "CryptoRsaSign"
    Const MAX_SIG_SIZE  As Long = MAX_RSA_KEY / 8
    Dim baRetVal()      As Byte
    Dim hHash           As Long
    Dim lSize           As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    
    If CryptCreateHash(uCtx.hProv, uCtx.HashAlgId, 0, 0, hHash) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptCreateHash"
        GoTo QH
    End If
    lSize = pvArraySize(baMessage)
    If lSize > 0 Then
        If CryptHashData(hHash, baMessage(0), lSize, 0) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptHashData"
            GoTo QH
        End If
    End If
    ReDim baRetVal(0 To MAX_SIG_SIZE - 1) As Byte
    Debug.Assert RedimStats(FUNC_NAME & ".baRetVal", UBound(baRetVal) + 1)
    lSize = UBound(baRetVal) + 1
    If CryptSignHash(hHash, AT_KEYEXCHANGE, 0, 0, baRetVal(0), lSize) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptSignHash"
        GoTo QH
    End If
    If UBound(baRetVal) <> lSize - 1 Then
        ReDim Preserve baRetVal(0 To lSize - 1) As Byte
        Debug.Assert RedimStats(FUNC_NAME & ".baRetVal", UBound(baRetVal) + 1)
    End If
    pvArrayReverse baRetVal
    CryptoRsaSign = baRetVal
QH:
    If hHash <> 0 Then
        Call CryptDestroyHash(hHash)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Public Function CryptoRsaVerify(uCtx As UcsRsaContextType, baMessage() As Byte, baSignature() As Byte) As Boolean
    Const FUNC_NAME     As String = "CryptoRsaVerify"
    Dim hHash           As Long
    Dim lSize           As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    Dim baRevSig()      As Byte
    
    If CryptCreateHash(uCtx.hProv, uCtx.HashAlgId, 0, 0, hHash) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptCreateHash"
        GoTo QH
    End If
    lSize = pvArraySize(baMessage)
    If lSize > 0 Then
        If CryptHashData(hHash, baMessage(0), lSize, 0) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptHashData"
            GoTo QH
        End If
    End If
    baRevSig = baSignature
    pvArrayReverse baRevSig
    If CryptVerifySignature(hHash, baRevSig(0), UBound(baRevSig) + 1, uCtx.hPubKey, 0, 0) = 0 Then
        hResult = Err.LastDllError
        '--- don't raise error on NTE_BAD_SIGNATURE
        If hResult <> NTE_BAD_SIGNATURE Then
            sApiSource = "CryptVerifySignature"
        End If
        GoTo QH
    End If
    '--- success
    CryptoRsaVerify = True
QH:
    If hHash <> 0 Then
        Call CryptDestroyHash(hHash)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Public Function CryptoRsaEncrypt(ByVal hKey As Long, baPlainText() As Byte) As Byte()
    Const FUNC_NAME     As String = "CryptoRsaEncrypt"
    Const MAX_RSA_BYTES As Long = MAX_RSA_KEY / 8
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    Dim lAlignedSize    As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    
    lSize = pvArraySize(baPlainText)
    lAlignedSize = (lSize + MAX_RSA_BYTES - 1 And -MAX_RSA_BYTES) + MAX_RSA_BYTES
    ReDim baRetVal(0 To lAlignedSize - 1) As Byte
    Debug.Assert RedimStats(FUNC_NAME & ".baRetVal", UBound(baRetVal) + 1)
    Call CopyMemory(baRetVal(0), baPlainText(0), lSize)
    If CryptEncrypt(hKey, 0, 1, 0, baRetVal(0), lSize, lAlignedSize) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptEncrypt"
        GoTo QH
    End If
    ReDim Preserve baRetVal(0 To lSize - 1) As Byte
    Debug.Assert RedimStats(FUNC_NAME & ".baRetVal", UBound(baRetVal) + 1)
    pvArrayReverse baRetVal
    CryptoRsaEncrypt = baRetVal
QH:
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

#If ImplUseBCrypt Then

Public Function CryptoRsaPssSign(baPrivKey() As Byte, baMessage() As Byte, ByVal lSignatureType As Long) As Byte()
    Const FUNC_NAME     As String = "CryptoRsaPssSign"
    Dim baRetVal()      As Byte
    Dim lPkiPtr         As Long
    Dim lKeyPtr         As Long
    Dim lKeySize        As Long
    Dim uKeyBlob        As CRYPT_BLOB_DATA
    Dim hAlgRSA         As Long
    Dim hKey            As Long
    Dim uPadInfo        As BCRYPT_PSS_PADDING_INFO
    Dim lSize           As Long
    Dim baHash()        As Byte
    Dim hResult         As Long
    Dim sApiSource      As String
    
    If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_PRIVATE_KEY_INFO, baPrivKey(0), UBound(baPrivKey) + 1, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lPkiPtr, 0) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptDecodeObjectEx(PKCS_PRIVATE_KEY_INFO)"
        GoTo QH
    End If
    Call CopyMemory(uKeyBlob, ByVal UnsignedAdd(lPkiPtr, 16), Len(uKeyBlob)) '--- dereference PCRYPT_PRIVATE_KEY_INFO->PrivateKey
    If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_RSA_PRIVATE_KEY, ByVal uKeyBlob.pbData, uKeyBlob.cbData, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lKeyPtr, lKeySize) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptDecodeObjectEx(PKCS_RSA_PRIVATE_KEY)"
        GoTo QH
    End If
    hResult = BCryptOpenAlgorithmProvider(hAlgRSA, StrPtr("RSA"), 0, 0)
    If hResult < 0 Then
        sApiSource = "BCryptOpenAlgorithmProvider"
        GoTo QH
    End If
    hResult = BCryptImportKeyPair(hAlgRSA, 0, StrPtr("CAPIPRIVATEBLOB"), hKey, lKeyPtr, lKeySize, 0)
    If hResult < 0 Then
        sApiSource = "BCryptImportKeyPair"
        GoTo QH
    End If
    Select Case lSignatureType
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA256
        uPadInfo.pszAlgId = StrPtr("SHA256")
        uPadInfo.cbSalt = 32
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA384
        uPadInfo.pszAlgId = StrPtr("SHA384")
        uPadInfo.cbSalt = 48
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA512, TLS_SIGNATURE_RSA_PSS_PSS_SHA512
        uPadInfo.pszAlgId = StrPtr("SHA512")
        uPadInfo.cbSalt = 64
    End Select
    pvArrayHash uPadInfo.cbSalt, baMessage, baHash
    hResult = BCryptSignHash(hKey, uPadInfo, baHash(0), UBound(baHash) + 1, ByVal 0, 0, lSize, BCRYPT_PAD_PSS)
    If hResult < 0 Then
        sApiSource = "BCryptSignHash"
        GoTo QH
    End If
    ReDim baRetVal(0 To lSize - 1) As Byte
    Debug.Assert RedimStats(FUNC_NAME & ".baRetVal", UBound(baRetVal) + 1)
    hResult = BCryptSignHash(hKey, uPadInfo, baHash(0), UBound(baHash) + 1, baRetVal(0), UBound(baRetVal) + 1, lSize, BCRYPT_PAD_PSS)
    If hResult < 0 Then
        sApiSource = "BCryptSignHash#2"
        GoTo QH
    End If
    CryptoRsaPssSign = baRetVal
QH:
    If hKey <> 0 Then
        Call BCryptDestroyKey(hKey)
    End If
    If hAlgRSA <> 0 Then
        Call BCryptCloseAlgorithmProvider(hAlgRSA, 0)
    End If
    If lKeyPtr <> 0 Then
        Call LocalFree(lKeyPtr)
    End If
    If lPkiPtr <> 0 Then
        Call LocalFree(lPkiPtr)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Public Function CryptoRsaPssVerify(baCert() As Byte, baMessage() As Byte, baSignature() As Byte, ByVal lSignatureType As Long) As Boolean
    Const FUNC_NAME     As String = "CryptoRsaPssVerify"
    Dim pCertContext    As Long
    Dim lPtr            As Long
    Dim hKey            As Long
    Dim uPadInfo        As BCRYPT_PSS_PADDING_INFO
    Dim baHash()        As Byte
    Dim hResult         As Long
    Dim sApiSource      As String

    pCertContext = CertCreateCertificateContext(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, baCert(0), UBound(baCert) + 1)
    If pCertContext = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CertCreateCertificateContext"
        GoTo QH
    End If
    Call CopyMemory(lPtr, ByVal UnsignedAdd(pCertContext, 12), 4)       '--- dereference pCertContext->pCertInfo
    lPtr = UnsignedAdd(lPtr, 56)
    If CryptImportPublicKeyInfoEx2(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, lPtr, 0, 0, hKey) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptImportPublicKeyInfoEx2"
        GoTo QH
    End If
    Select Case lSignatureType
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA256
        uPadInfo.pszAlgId = StrPtr("SHA256")
        uPadInfo.cbSalt = 32
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA384
        uPadInfo.pszAlgId = StrPtr("SHA384")
        uPadInfo.cbSalt = 48
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA512, TLS_SIGNATURE_RSA_PSS_PSS_SHA512
        uPadInfo.pszAlgId = StrPtr("SHA512")
        uPadInfo.cbSalt = 64
    End Select
    pvArrayHash uPadInfo.cbSalt, baMessage, baHash
    hResult = BCryptVerifySignature(hKey, uPadInfo, baHash(0), UBound(baHash) + 1, baSignature(0), UBound(baSignature) + 1, BCRYPT_PAD_PSS)
    If hResult < 0 Then
        If hResult <> STATUS_INVALID_SIGNATURE And hResult <> ERROR_INVALID_DATA Then
            sApiSource = "BCryptSignHash"
        End If
        GoTo QH
    End If
    CryptoRsaPssVerify = True
QH:
    If hKey <> 0 Then
        Call BCryptDestroyKey(hKey)
    End If
    If pCertContext <> 0 Then
        Call CertFreeCertificateContext(pCertContext)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

#Else

Public Function CryptoRsaPssSign(baPrivKey() As Byte, baMessage() As Byte, ByVal lSignatureType As Long) As Byte()
    Dim baRetVal()      As Byte
    Dim baBuffer()      As Byte
    Dim lKeyLen         As Long
    Dim lSize           As Long
    Dim lHalfSize       As Long
    Dim baModulo()      As Byte
    Dim baPrivExp()     As Byte
    Dim lHashSize       As Long
    Dim baHash()        As Byte
    Dim lSaltSize       As Long
    Dim baSalt()        As Byte
    Dim baDecr()        As Byte
    Dim lIdx            As Long
    Dim lPos            As Long
    Dim bMask           As Byte
    
    '--- retrieve keylen, modulo and private exponent from RSA private key
    baBuffer = Asn1DecodePrivateKeyFromDer(baPrivKey)
    Debug.Assert UBound(baBuffer) >= 16
    Call CopyMemory(lKeyLen, baBuffer(12), 4)
    lSize = (lKeyLen + 7) \ 8
    lHalfSize = (lKeyLen + 15) \ 16
    ReDim baModulo(0 To lSize - 1) As Byte
    Debug.Assert UBound(baBuffer) - 20 >= UBound(baModulo)
    Call CopyMemory(baModulo(0), baBuffer(20), UBound(baModulo) + 1)
    pvArrayReverse baModulo
    ReDim baPrivExp(0 To lSize - 1) As Byte
    Debug.Assert UBound(baBuffer) >= 20 + lSize + 5 * lHalfSize + UBound(baPrivExp)
    Call CopyMemory(baPrivExp(0), baBuffer(20 + lSize + 5 * lHalfSize), UBound(baPrivExp) + 1)
    pvArrayReverse baPrivExp
    '--- figure out hash and salt size
    Select Case lSignatureType
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA256
        lHashSize = 32
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA384
        lHashSize = 48
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA512, TLS_SIGNATURE_RSA_PSS_PSS_SHA512
        lHashSize = 64
    Case Else
        GoTo QH
    End Select
    lSaltSize = lHashSize
    '--- 2. Let |mHash| = |Hash(M)|, an octet string of length hLen.
    pvArrayHash lHashSize, baMessage, baHash
    '--- 3. If |emLen| < |hLen + sLen + 2|, output "encoding error" and stop.
    If lSize < lHashSize + lSaltSize + 2 Then
        GoTo QH
    End If
    '--- 4. Generate a random octet string salt of length sLen; if |sLen| = 0, then salt is the empty string.
    If lSaltSize > 0 Then
        ReDim baSalt(0 To lSaltSize - 1) As Byte
        CryptoRandomBytes VarPtr(baSalt(0)), lSaltSize
    Else
        baSalt = vbNullString
    End If
    '--- 5. Let |M'| = (0x)00 00 00 00 00 00 00 00 || mHash || salt;
    ReDim baBuffer(8 + lHashSize + lSaltSize - 1) As Byte
    Call CopyMemory(baBuffer(8), baHash(0), lHashSize)
    Call CopyMemory(baBuffer(8 + lHashSize), baSalt(0), lSaltSize)
    '--- 6. Let |H| = Hash(M'), an octet string of length hLen.
    pvArrayHash lHashSize, baBuffer, baHash
    '--- 7. Generate an octet string |PS| consisting of |emLen - sLen - hLen - 2| zero octets. The length of PS may be 0.
    '--- 8. Let |DB| = PS || 0x01 || salt; DB is an octet string of length |emLen - hLen - 1|.
    ReDim baDecr(0 To lSize - 1) As Byte
    baDecr(lSize - lHashSize - lSaltSize - 2) = &H1
    Call CopyMemory(baDecr(lSize - lHashSize - lSaltSize - 1), baSalt(0), lSaltSize)
    Call CopyMemory(baDecr(lSize - lHashSize - 1), baHash(0), lHashSize)
    '--- 9. Let |dbMask| = MGF(H, emLen - hLen - 1).
    '--- 10. Let |maskedDB| = DB \xor dbMask.
    ReDim baSeed(0 To lHashSize - 1 + 4) As Byte '--- leave 4 more bytes at the end for counter
    Call CopyMemory(baSeed(0), baDecr(lSize - lHashSize - 1), lHashSize)
    Do
        pvArrayHash lHashSize, baSeed, baHash
        For lIdx = 0 To UBound(baHash)
            baDecr(lPos) = baDecr(lPos) Xor baHash(lIdx)
            lPos = lPos + 1
            If lPos >= lSize - lHashSize - 1 Then
                Exit Do
            End If
        Next
        pvArrayIncCounter baSeed, lHashSize + 3
    Loop
    '--- 11. Set the leftmost |8 * emLen - emBits| bits of the leftmost octet in |maskedDB| to zero.
    bMask = &HFF \ (2 ^ (lSize * 8 - lKeyLen))
    baDecr(0) = baDecr(0) And (bMask \ 2)
    '--- 12. Let |EM| = maskedDB || H || 0xbc.
    baDecr(lSize - 1) = &HBC
    '--- 13. Output EM.
    ReDim baRetVal(0 To lSize - 1) As Byte
    Debug.Assert pvPatchTrampoline(AddressOf pvCallRsaModExp)
    Call pvCallRsaModExp(m_uData.Pfn(ucsPfnRsaModExp), lSize, baDecr(0), baPrivExp(0), baModulo(0), baRetVal(0))
QH:
    CryptoRsaPssSign = baRetVal
End Function

Public Function CryptoRsaPssVerify(baCert() As Byte, baMessage() As Byte, baSignature() As Byte, ByVal lSignatureType As Long) As Boolean
    Const FUNC_NAME     As String = "CryptoRsaPssVerify"
    Dim baPubKey()      As Byte
    Dim lKeyLen         As Long
    Dim lSize           As Long
    Dim baBuffer()      As Byte
    Dim baModulo()      As Byte
    Dim baDecr()        As Byte
    Dim baPubExp()      As Byte
    Dim baSeed()        As Byte
    Dim lHashSize       As Long
    Dim baHash()        As Byte
    Dim lSaltSize       As Long
    Dim baSalt()        As Byte
    Dim lPos            As Long
    Dim lIdx            As Long
    Dim bMask           As Byte
    
    baPubKey = Asn1DecodePublicKeyFromDer(baCert, KeyLen:=lKeyLen, KeyBlob:=baBuffer)
    lSize = (lKeyLen + 7) \ 8
    '--- check signature size
    If UBound(baSignature) + 1 <> lSize Then
        GoTo QH
    End If
    '--- retrieve RSA public exponent
    ReDim baPubExp(0 To 3) As Byte
    Debug.Assert RedimStats(FUNC_NAME & ".baPubExp", UBound(baPubExp) + 1)
    Call CopyMemory(baPubExp(0), baBuffer(16), UBound(baPubExp) + 1)        '--- 16 = sizeof(PUBLICKEYSTRUC) + offset(RSAPUBKEY, pubexp)
    baPubExp = pvArrayReverseResize(baPubExp, lSize)
    '--- retrieve RSA key modulo
    ReDim baModulo(0 To lSize - 1) As Byte
    Debug.Assert RedimStats(FUNC_NAME & ".baModulo", UBound(baModulo) + 1)
    Debug.Assert UBound(baBuffer) - 20 >= UBound(baModulo)
    Call CopyMemory(baModulo(0), baBuffer(20), UBound(baModulo) + 1)        '--- 20 = sizeof(PUBLICKEYSTRUC) + sizeof(RSAPUBKEY)
    pvArrayReverse baModulo
    '--- decrypt RSA signature
    ReDim baDecr(0 To lSize - 1) As Byte
    Debug.Assert RedimStats(FUNC_NAME & ".baDecr", UBound(baDecr) + 1)
    Debug.Assert pvPatchTrampoline(AddressOf pvCallRsaModExp)
    Call pvCallRsaModExp(m_uData.Pfn(ucsPfnRsaModExp), lSize, baSignature(0), baPubExp(0), baModulo(0), baDecr(0))
    '--- from RFC 8017, Section 9.1.2.
    Select Case lSignatureType
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA256, TLS_SIGNATURE_RSA_PSS_PSS_SHA256
        lHashSize = 32
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA384, TLS_SIGNATURE_RSA_PSS_PSS_SHA384
        lHashSize = 48
    Case TLS_SIGNATURE_RSA_PSS_RSAE_SHA512, TLS_SIGNATURE_RSA_PSS_PSS_SHA512
        lHashSize = 64
    Case Else
        GoTo QH
    End Select
    lSaltSize = lHashSize
    '--- 3. If |emLen| < |hLen + sLen + 2|, output "inconsistent" and stop.
    If lSize < lHashSize + lSaltSize + 2 Then
        GoTo QH
    End If
    '--- 4. If the rightmost octet of |EM| does not have hexadecimal value 0xbc, output "inconsistent" and stop.
    If baDecr(lSize - 1) <> &HBC Then
        GoTo QH
    End If
    '--- 5. Let |maskedDB| be the leftmost |emLen - hLen - 1| octets of |EM|, and let |H| be the next |hLen| octets.
    '--- 6. If the leftmost |8 * emLen - emBits| bits of the leftmost octet in |maskedDB| are not all equal to zero,
    '---    output "inconsistent" and stop.
    bMask = &HFF \ (2 ^ (lSize * 8 - lKeyLen))
    If (baDecr(0) And Not bMask) <> 0 Then
        GoTo QH
    End If
    '--- 7. Let |dbMask| = MGF(H, emLen - hLen - 1).
    '--- 8. Let |DB| = maskedDB \xor dbMask.
    ReDim baSeed(0 To lHashSize - 1 + 4) As Byte '--- leave 4 more bytes at the end for counter
    Call CopyMemory(baSeed(0), baDecr(lSize - lHashSize - 1), lHashSize)
    Do
        pvArrayHash lHashSize, baSeed, baHash
        For lIdx = 0 To UBound(baHash)
            baDecr(lPos) = baDecr(lPos) Xor baHash(lIdx)
            lPos = lPos + 1
            If lPos >= lSize - lHashSize - 1 Then
                Exit Do
            End If
        Next
        pvArrayIncCounter baSeed, lHashSize + 3
    Loop
    '--- 9. Set the leftmost |8 * emLen - emBits| bits of the leftmost octet in |DB| to zero.
    '--- note: troubles w/ sign bit so use (bMask \ 2) to clear MSB
    baDecr(0) = baDecr(0) And (bMask \ 2)
    '--- 10. If the |emLen - hLen - sLen - 2| leftmost octets of |DB| are not zero or if the octet at position
    '---     |emLen - hLen - sLen - 1| (the leftmost position is "position 1") does not have hexadecimal
    '---     value 0x01, output "inconsistent" and stop.
    For lIdx = 0 To lPos - lHashSize - 2
        If baDecr(lIdx) <> 0 Then
            Exit For
        End If
    Next
    If lIdx <> lPos - lHashSize - 1 Then
        GoTo QH
    End If
    If baDecr(lPos - lHashSize - 1) <> &H1 Then
        GoTo QH
    End If
    '--- 11. Let |salt| be the last |sLen| octets of |DB|.
    ReDim baSalt(0 To lSaltSize - 1) As Byte
    Call CopyMemory(baSalt(0), baDecr(lPos - lSaltSize), lSaltSize)
    '--- 12. Let |M'| = (0x)00 00 00 00 00 00 00 00 || mHash || salt
    ReDim baSeed(0 To 8 + lHashSize + lSaltSize - 1) As Byte
    pvArrayHash lHashSize, baMessage, baHash
    Call CopyMemory(baSeed(8), baHash(0), lHashSize)
    Call CopyMemory(baSeed(8 + lHashSize), baSalt(0), lSaltSize)
    '--- 13. Let |H'| = Hash(M'), an octet string of length |hLen|.
    pvArrayHash lHashSize, baSeed, baHash
    '--- |H| is still not de-masked in decrypted buffer
    ReDim baBuffer(0 To lHashSize - 1) As Byte
    Call CopyMemory(baBuffer(0), baDecr(lPos), lHashSize)
    '--- 14. If |H| = |H'|, output "consistent." Otherwise, output "inconsistent."
    If StrConv(baHash, vbUnicode) <> StrConv(baBuffer, vbUnicode) Then
        GoTo QH
    End If
    '--- success
    CryptoRsaPssVerify = True
QH:
End Function

#End If

Public Function Asn1DecodePrivateKeyFromDer(baPrivKey() As Byte) As Byte()
    Const FUNC_NAME     As String = "Asn1DecodePrivateKeyFromDer"
    Dim baRetVal()      As Byte
    Dim lPkiPtr         As Long
    Dim uKeyBlob        As CRYPT_BLOB_DATA
    Dim lKeyPtr         As Long
    Dim lKeySize        As Long
    Dim uEccKeyInfo     As CRYPT_ECC_PRIVATE_KEY_INFO
    Dim lSize           As Long
    Dim hResult         As Long
    Dim sApiSource      As String

    If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_PRIVATE_KEY_INFO, baPrivKey(0), UBound(baPrivKey) + 1, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lPkiPtr, 0) <> 0 Then
        Call CopyMemory(uKeyBlob, ByVal UnsignedAdd(lPkiPtr, 16), Len(uKeyBlob)) '--- dereference PCRYPT_PRIVATE_KEY_INFO->PrivateKey
        If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_RSA_PRIVATE_KEY, ByVal uKeyBlob.pbData, uKeyBlob.cbData, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lKeyPtr, lKeySize) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptDecodeObjectEx(PKCS_RSA_PRIVATE_KEY)"
            GoTo QH
        End If
        ReDim baRetVal(0 To lKeySize - 1) As Byte
        Debug.Assert RedimStats(FUNC_NAME & ".baRetVal", UBound(baRetVal) + 1)
        Call CopyMemory(baRetVal(0), ByVal lKeyPtr, lKeySize)
    ElseIf CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, X509_ECC_PRIVATE_KEY, baPrivKey(0), UBound(baPrivKey) + 1, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lKeyPtr, 0) <> 0 Then
        Call CopyMemory(uEccKeyInfo, ByVal lKeyPtr, Len(uEccKeyInfo))
        ReDim baRetVal(0 To uEccKeyInfo.PrivateKey.cbData - 1) As Byte
        Debug.Assert RedimStats(FUNC_NAME & ".baRetVal", UBound(baRetVal) + 1)
        Call CopyMemory(baRetVal(0), ByVal uEccKeyInfo.PrivateKey.pbData, uEccKeyInfo.PrivateKey.cbData)
    ElseIf Err.LastDllError = ERROR_FILE_NOT_FOUND Then
        '--- no X509_ECC_PRIVATE_KEY struct type on NT4 -> decode manually
        Call CopyMemory(lSize, baPrivKey(6), 1)
        If 7 + lSize <= UBound(baPrivKey) Then
            ReDim baRetVal(0 To lSize - 1) As Byte
            Debug.Assert RedimStats(FUNC_NAME & ".baRetVal", UBound(baRetVal) + 1)
            Call CopyMemory(baRetVal(0), baPrivKey(7), lSize)
        Else
            hResult = ERROR_FILE_NOT_FOUND
            sApiSource = "CryptDecodeObjectEx(X509_ECC_PRIVATE_KEY)"
            GoTo QH
        End If
    Else
        hResult = Err.LastDllError
        sApiSource = "CryptDecodeObjectEx(X509_ECC_PRIVATE_KEY)"
        GoTo QH
    End If
    Asn1DecodePrivateKeyFromDer = baRetVal
QH:
    If lPkiPtr <> 0 Then
        Call LocalFree(lPkiPtr)
    End If
    If lKeyPtr <> 0 Then
        Call LocalFree(lPkiPtr)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Public Function Asn1DecodePublicKeyFromDer(baCert() As Byte, Optional AlgoObjId As String, Optional KeyLen As Long, Optional KeyBlob As Variant) As Byte()
    Const FUNC_NAME     As String = "Asn1DecodePublicKeyFromDer"
    Dim baRetVal()      As Byte
    Dim pCertContext    As Long
    Dim lPtr            As Long
    Dim uPublicKeyInfo  As CERT_PUBLIC_KEY_INFO
    Dim hProv           As Long
    Dim hKey            As Long
    Dim lSize           As Long
    Dim baBuffer()      As Byte
    Dim hResult         As Long
    Dim sApiSource      As String

    pCertContext = CertCreateCertificateContext(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, baCert(0), UBound(baCert) + 1)
    If pCertContext = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CertCreateCertificateContext"
        GoTo QH
    End If
    Call CopyMemory(lPtr, ByVal UnsignedAdd(pCertContext, 12), 4)       '--- dereference pCertContext->pCertInfo
    lPtr = UnsignedAdd(lPtr, 56)                                        '--- &pCertContext->pCertInfo->SubjectPublicKeyInfo
    Call CopyMemory(uPublicKeyInfo, ByVal lPtr, Len(uPublicKeyInfo))
    AlgoObjId = pvToString(uPublicKeyInfo.Algorithm.pszObjId)
    ReDim baRetVal(0 To uPublicKeyInfo.PublicKey.cbData - 1) As Byte
    Debug.Assert RedimStats(FUNC_NAME & ".baRetVal", UBound(baRetVal) + 1)
    Call CopyMemory(baRetVal(0), ByVal uPublicKeyInfo.PublicKey.pbData, uPublicKeyInfo.PublicKey.cbData)
    '--- don't report failure on keylen and blob retrieval
    If CryptAcquireContext(hProv, 0, 0, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0 Then
        If CryptImportPublicKeyInfo(hProv, X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, ByVal lPtr, hKey) <> 0 Then
            Call CryptGetKeyParam(hKey, KP_KEYLEN, KeyLen, 4, 0)
            If Not IsMissing(KeyBlob) Then
                If CryptExportKey(hKey, 0, PUBLICKEYBLOB, 0, ByVal 0, lSize) <> 0 Then
                    ReDim baBuffer(0 To lSize - 1) As Byte
                    If CryptExportKey(hKey, 0, PUBLICKEYBLOB, 0, baBuffer(0), lSize) <> 0 Then
                        KeyBlob = baBuffer
                    End If
                End If
            End If
        End If
    End If
    '--- success
    Asn1DecodePublicKeyFromDer = baRetVal
QH:
    If hKey <> 0 Then
        Call CryptDestroyKey(hKey)
    End If
    If hProv <> 0 Then
        Call CryptReleaseContext(hProv, 0)
    End If
    If pCertContext <> 0 Then
        Call CertFreeCertificateContext(pCertContext)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Public Sub RemoveCollection(ByVal oCol As Collection, Index As Variant)
    If Not oCol Is Nothing Then
        pvCallCollectionRemove oCol, Index
    End If
End Sub

Public Function SearchCollection(ByVal oCol As Collection, Index As Variant, Optional RetVal As Variant) As Boolean
    Dim vItem           As Variant
    
    If oCol Is Nothing Then
        GoTo QH
    ElseIf pvCallCollectionItem(oCol, Index, vItem) < 0 Then
        GoTo QH
    End If
    If IsObject(vItem) Then
        Set RetVal = vItem
    Else
        RetVal = vItem
    End If
    '--- success
    SearchCollection = True
QH:
End Function

Public Function FromBase64Array(sText As String) As Byte()
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    lSize = (Len(sText) \ 4) * 3
    ReDim baRetVal(0 To lSize - 1) As Byte
    Debug.Assert RedimStats("FromBase64Array.baRetVal", UBound(baRetVal) + 1)
    pvThunkAllocate sText, VarPtr(baRetVal(0))
    If Right$(sText, 2) = "==" Then
        ReDim Preserve baRetVal(0 To lSize - 3)
        Debug.Assert RedimStats("FromBase64Array.baRetVal", UBound(baRetVal) + 1)
    ElseIf Right$(sText, 1) = "=" Then
        ReDim Preserve baRetVal(0 To lSize - 2)
        Debug.Assert RedimStats("FromBase64Array.baRetVal", UBound(baRetVal) + 1)
    End If
    FromBase64Array = baRetVal
End Function

'= private ===============================================================

Private Function pvArraySize(baArray() As Byte) As Long
    Dim lPtr            As Long
    
    '--- peek long at ArrPtr(baArray)
    Call CopyMemory(lPtr, ByVal ArrPtr(baArray), 4)
    If lPtr <> 0 Then
        pvArraySize = UBound(baArray) + 1
    End If
End Function

Private Sub pvArrayReverse(baData() As Byte)
    Dim lIdx            As Long
    Dim bTemp           As Byte
    
    For lIdx = 0 To UBound(baData) \ 2
        bTemp = baData(lIdx)
        baData(lIdx) = baData(UBound(baData) - lIdx)
        baData(UBound(baData) - lIdx) = bTemp
    Next
End Sub

Private Sub pvArrayHash(ByVal lHashSize As Long, baInput() As Byte, baRetVal() As Byte)
    Select Case lHashSize
    Case 32
        baRetVal = CryptoHashSha256(baInput, 0)
    Case 48
        baRetVal = CryptoHashSha384(baInput, 0)
    Case 64
        baRetVal = CryptoHashSha512(baInput, 0)
    Case Else
        Err.Raise vbObjectError, , "Invalid hash size"
    End Select
End Sub

Private Function pvArrayReverseResize(baData() As Byte, ByVal lSize As Long) As Byte()
    Dim baRetVal()      As Byte
    Dim lIdx            As Long
    Dim lJdx            As Long
    
    ReDim baRetVal(0 To lSize - 1) As Byte
    For lIdx = 0 To UBound(baData)
        lJdx = lSize - 1 - lIdx
        If lJdx >= 0 Then
            baRetVal(lJdx) = baData(lIdx)
        End If
    Next
    pvArrayReverseResize = baRetVal
End Function

Private Sub pvArrayIncCounter(baInput() As Byte, ByVal lPos As Long)
    Do While lPos >= 0
        If baInput(lPos) < 255 Then
            baInput(lPos) = baInput(lPos) + 1
            Exit Do
        Else
            baInput(lPos) = 0
            lPos = lPos - 1
        End If
    Loop
End Sub

Private Function pvThunkAllocate(sText As String, Optional ByVal ThunkPtr As Long) As Long
    Static Map(0 To &H3FF) As Long
    Dim baInput()       As Byte
    Dim lIdx            As Long
    Dim lChar           As Long
    Dim lPtr            As Long
    
    If ThunkPtr <> 0 Then
        pvThunkAllocate = ThunkPtr
    Else
        pvThunkAllocate = VirtualAlloc(0, (Len(sText) \ 4) * 3, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        If pvThunkAllocate = 0 Then
            Exit Function
        End If
    End If
    '--- init decoding maps
    If Map(65) = 0 Then
        baInput = StrConv("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", vbFromUnicode)
        For lIdx = 0 To UBound(baInput)
            lChar = baInput(lIdx)
            Map(&H0 + lChar) = lIdx * (2 ^ 2)
            Map(&H100 + lChar) = (lIdx And &H30) \ (2 ^ 4) Or (lIdx And &HF) * (2 ^ 12)
            Map(&H200 + lChar) = (lIdx And &H3) * (2 ^ 22) Or (lIdx And &H3C) * (2 ^ 6)
            Map(&H300 + lChar) = lIdx * (2 ^ 16)
        Next
    End If
    '--- base64 decode loop
    baInput = StrConv(Replace(Replace(sText, vbCr, vbNullString), vbLf, vbNullString), vbFromUnicode)
    lPtr = pvThunkAllocate
    For lIdx = 0 To UBound(baInput) - 3 Step 4
        lChar = Map(baInput(lIdx + 0)) Or Map(&H100 + baInput(lIdx + 1)) Or Map(&H200 + baInput(lIdx + 2)) Or Map(&H300 + baInput(lIdx + 3))
        Call CopyMemory(ByVal lPtr, lChar, 3)
        lPtr = UnsignedAdd(lPtr, 3)
    Next
End Function

Private Function pvPatchTrampoline(ByVal Pfn As Long) As Boolean
    Dim bInIDE          As Boolean
 
    Debug.Assert pvSetTrue(bInIDE)
    If bInIDE Then
        Call CopyMemory(Pfn, ByVal UnsignedAdd(Pfn, &H16), 4)
    Else
        Call VirtualProtect(Pfn, 8, PAGE_EXECUTE_READWRITE, 0)
    End If
    ' 0:  58                      pop    eax
    ' 1:  59                      pop    ecx
    ' 2:  50                      push   eax
    ' 3:  ff e1                   jmp    ecx
    ' 5:  90                      nop
    ' 6:  90                      nop
    ' 7:  90                      nop
    Call CopyMemory(ByVal Pfn, -802975883527609.7192@, 8)
    '--- success
    pvPatchTrampoline = True
End Function

Private Function pvPatchMethodTrampoline(ByVal Pfn As Long, ByVal lMethodIdx As Long) As Boolean
    Dim bInIDE          As Boolean

    Debug.Assert pvSetTrue(bInIDE)
    If bInIDE Then
        '--- note: IDE is not large-address aware
        Call CopyMemory(Pfn, ByVal Pfn + &H16, 4)
    Else
        Call VirtualProtect(Pfn, 12, PAGE_EXECUTE_READWRITE, 0)
    End If
    ' 0: 8B 44 24 04          mov         eax,dword ptr [esp+4]
    ' 4: 8B 00                mov         eax,dword ptr [eax]
    ' 6: FF A0 00 00 00 00    jmp         dword ptr [eax+lMethodIdx*4]
    Call CopyMemory(ByVal Pfn, -684575231150992.4725@, 8)
    Call CopyMemory(ByVal (Pfn Xor &H80000000) + 8 Xor &H80000000, lMethodIdx * 4, 4)
    '--- success
    pvPatchMethodTrampoline = True
End Function

Private Function pvSetTrue(bValue As Boolean) As Boolean
    bValue = True
    pvSetTrue = True
End Function

Private Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function

Public Function ReadBinaryFile(sFile As String) As Byte()
    Dim baBuffer()      As Byte
    Dim nFile           As Integer
    
    baBuffer = vbNullString
    If GetFileAttributes(sFile) <> -1 Then
        nFile = FreeFile
        Open sFile For Binary Access Read Shared As nFile
        If LOF(nFile) > 0 Then
            ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
            Debug.Assert RedimStats("ReadBinaryFile.baBuffer", UBound(baBuffer) + 1)
            Get nFile, , baBuffer
        End If
        Close nFile
    End If
    ReadBinaryFile = baBuffer
End Function

#If ImplUseLibSodium Then
    Private Sub crypto_hash_sha384_init(baCtx() As Byte)
        Static baSha384State() As Byte
        
        If pvArraySize(baSha384State) = 0 Then
            baSha384State = FromBase64Array(STR_LIBSODIUM_SHA384_STATE)
        End If
        Debug.Assert pvArraySize(baCtx) >= crypto_hash_sha512_statebytes()
        Call crypto_hash_sha512_init(baCtx(0))
        Call CopyMemory(baCtx(0), baSha384State(0), UBound(baSha384State) + 1)
    End Sub
#End If

'= trampolines ===========================================================

Private Function pvCallCurve25519Multiply(ByVal Pfn As Long, pSecretPtr As Byte, pPubKeyPtr As Byte, pPrivKeyPtr As Byte) As Long
    ' void cf_curve25519_mul(uint8_t out[32], const uint8_t priv[32], const uint8_t pub[32])
End Function

Private Function pvCallCurve25519MulBase(ByVal Pfn As Long, pPubKeyPtr As Byte, pPrivKeyPtr As Byte) As Long
    ' void cf_curve25519_mul_base(uint8_t out[32], const uint8_t priv[32])
End Function

Private Function pvCallSecpMakeKey(ByVal Pfn As Long, pPubKeyPtr As Byte, pPrivKeyPtr As Byte) As Long
    ' int ecc_make_key(uint8_t p_publicKey[ECC_BYTES+1], uint8_t p_privateKey[ECC_BYTES]);
    ' int ecc_make_key384(uint8_t p_publicKey[ECC_BYTES_384+1], const uint8_t p_privateKey[ECC_BYTES_384])
End Function

Private Function pvCallSecpSharedSecret(ByVal Pfn As Long, pPubKeyPtr As Byte, pPrivKeyPtr As Byte, pSecretPtr As Byte) As Long
    ' int ecdh_shared_secret(const uint8_t p_publicKey[ECC_BYTES+1], const uint8_t p_privateKey[ECC_BYTES], uint8_t p_secret[ECC_BYTES]);
    ' int ecdh_shared_secret384(const uint8_t p_publicKey[ECC_BYTES_384+1], const uint8_t p_privateKey[ECC_BYTES_384], uint8_t p_secret[ECC_BYTES_384])
End Function

Private Function pvCallSecpUncompressKey(ByVal Pfn As Long, pPubKeyPtr As Byte, pUncompressedKeyPtr As Byte) As Long
    ' int ecdh_uncompress_key(const uint8_t p_publicKey[ECC_BYTES + 1], uint8_t p_uncompressedKey[2 * ECC_BYTES + 1])
    ' int ecdh_uncompress_key384(const uint8_t p_publicKey[ECC_BYTES_384 + 1], uint8_t p_uncompressedKey[2 * ECC_BYTES_384 + 1])
End Function

Private Function pvCallSecpSign(ByVal Pfn As Long, pPrivKeyPtr As Byte, pHashPtr As Byte, pRandomPtr As Byte, pSignaturePtr As Byte) As Long
    ' int ecdsa_sign(const uint8_t p_privateKey[ECC_BYTES], const uint8_t p_hash[ECC_BYTES], uint64_t k[NUM_ECC_DIGITS], uint8_t p_signature[ECC_BYTES*2])
    ' int ecdsa_sign384(const uint8_t p_privateKey[ECC_BYTES_384], const uint8_t p_hash[ECC_BYTES_384], uint64_t k[NUM_ECC_DIGITS_384], uint8_t p_signature[ECC_BYTES_384*2])
End Function

Private Function pvCallSecpVerify(ByVal Pfn As Long, pPubKeyPtr As Byte, pHashPtr As Byte, pSignaturePtr As Byte) As Long
    ' int ecdsa_verify(const uint8_t p_publicKey[ECC_BYTES+1], const uint8_t p_hash[ECC_BYTES], const uint8_t p_signature[ECC_BYTES*2])
    ' int ecdsa_verify384(const uint8_t p_publicKey[ECC_BYTES_384+1], const uint8_t p_hash[ECC_BYTES_384], const uint8_t p_signature[ECC_BYTES_384*2])
End Function

Private Function pvCallSha2Init(ByVal Pfn As Long, ByVal lCtxPtr As Long) As Long
    ' void cf_sha256_init(cf_sha256_context *ctx)
    ' void cf_sha384_init(cf_sha384_context *ctx)
    ' void cf_sha512_init(cf_sha512_context *ctx)
End Function

Private Function pvCallSha2Update(ByVal Pfn As Long, ByVal lCtxPtr As Long, ByVal lDataPtr As Long, ByVal lSize As Long) As Long
    ' void cf_sha256_update(cf_sha256_context *ctx, const void *data, size_t nbytes)
    ' void cf_sha384_update(cf_sha384_context *ctx, const void *data, size_t nbytes)
    ' void cf_sha512_update(cf_sha512_context *ctx, const void *data, size_t nbytes)
End Function

Private Function pvCallSha2Final(ByVal Pfn As Long, ByVal lCtxPtr As Long, pHashPtr As Byte) As Long
    ' void cf_sha256_digest_final(cf_sha256_context *ctx, uint8_t hash[LNG_SHA256_HASHSZ])
    ' void cf_sha384_digest_final(cf_sha384_context *ctx, uint8_t hash[LNG_SHA384_HASHSZ])
    ' void cf_sha512_digest_final(cf_sha512_context *ctx, uint8_t hash[LNG_SHA384_HASHSZ])
End Function

Private Function pvCallChacha20Poly1305Encrypt( _
            ByVal Pfn As Long, pKeyPtr As Byte, pNoncePtr As Byte, _
            ByVal lHeaderPtr As Long, ByVal lHeaderSize As Long, _
            pPlaintTextPtr As Byte, ByVal lPlaintTextSize As Long, _
            pCipherTextPtr As Byte, pTagPtr As Byte) As Long
    ' void cf_chacha20poly1305_encrypt(const uint8_t key[32], const uint8_t nonce[12], const uint8_t *header, size_t nheader,
    '                                  const uint8_t *plaintext, size_t nbytes, uint8_t *ciphertext, uint8_t tag[16])
End Function

Private Function pvCallChacha20Poly1305Decrypt( _
            ByVal Pfn As Long, pKeyPtr As Byte, pNoncePtr As Byte, _
            pHeaderPtr As Byte, ByVal lHeaderSize As Long, _
            pCipherTextPtr As Byte, ByVal lCipherTextSize As Long, _
            pTagPtr As Byte, pPlaintTextPtr As Byte) As Long
    ' int cf_chacha20poly1305_decrypt(const uint8_t key[32], const uint8_t nonce[12], const uint8_t *header, size_t nheader,
    '                                 const uint8_t *ciphertext, size_t nbytes, const uint8_t tag[16], uint8_t *plaintext)
End Function

Private Function pvCallAesGcmEncrypt( _
            ByVal Pfn As Long, pCipherTextPtr As Byte, pTagPtr As Byte, pPlaintTextPtr As Byte, ByVal lPlaintTextSize As Long, _
            ByVal lHeaderPtr As Long, ByVal lHeaderSize As Long, pNoncePtr As Byte, pKeyPtr As Byte, ByVal lKeySize As Long) As Long
    ' void cf_aesgcm_encrypt(uint8_t *c, uint8_t *mac, const uint8_t *m, const size_t mlen, const uint8_t *ad, const size_t adlen,
    '                        const uint8_t *npub, const uint8_t *k, size_t klen)
End Function

Private Function pvCallAesGcmDecrypt( _
            ByVal Pfn As Long, pPlaintTextPtr As Byte, pCipherTextPtr As Byte, ByVal lCipherTextSize As Long, pTagPtr As Byte, _
            pHeaderPtr As Byte, ByVal lHeaderSize As Long, pNoncePtr As Byte, pKeyPtr As Byte, ByVal lKeySize As Long) As Long
    ' void cf_aesgcm_decrypt(uint8_t *m, const uint8_t *c, const size_t clen, const uint8_t *mac, const uint8_t *ad, const size_t adlen,
    '                        const uint8_t *npub, const uint8_t *k, const size_t klen)
End Function

Private Function pvCallRsaModExp(ByVal Pfn As Long, ByVal lSize As Long, pBasePtr As Byte, pExpPtr As Byte, pModuloPtr As Byte, pResultPtr As Byte) As Long
    ' void rsa_modexp(uint32_t maxbytes, const uint8_t *b, const uint8_t *e, const uint8_t *m, uint8_t *r)
End Function

Private Function pvCallCollectionItem(ByVal oCol As Collection, Index As Variant, Optional RetVal As Variant) As Long
    Const IDX_COLLECTION_ITEM As Long = 7
    
    pvPatchMethodTrampoline AddressOf mdTlsCrypto.pvCallCollectionItem, IDX_COLLECTION_ITEM
    pvCallCollectionItem = pvCallCollectionItem(oCol, Index, RetVal)
End Function

Private Function pvCallCollectionRemove(ByVal oCol As Collection, Index As Variant) As Long
    Const IDX_COLLECTION_REMOVE As Long = 10
    
    pvPatchMethodTrampoline AddressOf mdTlsCrypto.pvCallCollectionRemove, IDX_COLLECTION_REMOVE
    pvCallCollectionRemove = pvCallCollectionRemove(oCol, Index)
End Function

'=========================================================================
' PKI
'=========================================================================

Public Function PkiPemImportCertificates(ByVal vPemFiles As Variant, cCerts As Collection, baPrivKey() As Byte) As Boolean
    Dim vElem           As Variant
    Dim sPemText        As String
    Dim cKeys           As Collection
    
    If VarType(vPemFiles) = vbString Then
        vPemFiles = Array(vPemFiles)
    End If
    For Each vElem In vPemFiles
        sPemText = StrConv(CStr(ReadBinaryFile(CStr(vElem))), vbUnicode)
        PkiPemGetTextPortions sPemText, "PRIVATE KEY", cKeys
        PkiPemGetTextPortions sPemText, "EC PRIVATE KEY", cKeys
        PkiPemGetTextPortions sPemText, "CERTIFICATE", cCerts
    Next
    If SearchCollection(cKeys, 1, RetVal:=baPrivKey) Then
        '--- success
        PkiPemImportCertificates = True
    End If
End Function

Public Function PkiPemImportRootCaCertStore(sCaBundlePemFile As String) As Long
    Const FUNC_NAME     As String = "PkiPemImportRootCaCertStore"
    Dim hCertStore      As Long
    Dim cCerts          As Collection
    Dim vElem           As Variant
    Dim baCert()        As Byte
    Dim hResult         As Long
    Dim sApiSource      As String
    
    Set cCerts = PkiPemGetTextPortions(StrConv(CStr(ReadBinaryFile(CStr(sCaBundlePemFile))), vbUnicode), "CERTIFICATE")
    If cCerts.Count = 0 Then
        GoTo QH
    End If
    hCertStore = CertOpenStore(CERT_STORE_PROV_MEMORY, 0, 0, CERT_STORE_CREATE_NEW_FLAG, 0)
    If hCertStore = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CertOpenStore"
        GoTo QH
    End If
    For Each vElem In cCerts
        baCert = vElem
        If CertAddEncodedCertificateToStore(hCertStore, X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, baCert(0), UBound(baCert) + 1, CERT_STORE_ADD_USE_EXISTING, 0) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CertAddEncodedCertificateToStore"
            GoTo QH
        End If
    Next
    '--- commit
    PkiPemImportRootCaCertStore = hCertStore
    hCertStore = 0
QH:
    If hCertStore <> 0 Then
        Call CertCloseStore(hCertStore, 0)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Private Function PkiPemGetTextPortions(sContents As String, sBoundary As String, Optional RetVal As Collection) As Collection
    Dim vSplit          As Variant
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim bInside         As Boolean
    Dim lStart          As Long
    Dim lSize           As Long
    Dim sPortion        As String
    
    If RetVal Is Nothing Then
        Set RetVal = New Collection
    End If
    vSplit = Split(Replace(sContents, vbCr, vbNullString), vbLf)
    For lIdx = 0 To UBound(vSplit)
        If Not bInside Then
            If InStr(vSplit(lIdx), "-----BEGIN " & sBoundary & "-----") > 0 Then
                lStart = lIdx + 1
                lSize = 0
                bInside = True
            End If
        Else
            If InStr(vSplit(lIdx), "-----END " & sBoundary & "-----") > 0 Then
                sPortion = String$(lSize, 0)
                lSize = 1
                For lJdx = lStart To lIdx - 1
                    If InStr(vSplit(lJdx), ":") = 0 Then
                        Mid$(sPortion, lSize, Len(vSplit(lJdx))) = vSplit(lJdx)
                        lSize = lSize + Len(vSplit(lJdx))
                    End If
                Next
                If Not SearchCollection(RetVal, sPortion) Then
                    RetVal.Add FromBase64Array(sPortion), sPortion
                End If
                bInside = False
            ElseIf InStr(vSplit(lIdx), ":") = 0 Then
                lSize = lSize + Len(vSplit(lIdx))
            End If
        End If
    Next
    Set PkiPemGetTextPortions = RetVal
End Function

Public Function PkiPkcs12ImportCertificates(sPfxFile As String, sPassword As String, cCerts As Collection, baPrivKey() As Byte) As Boolean
    Const FUNC_NAME     As String = "PkiPkcs12ImportCertificates"
    Dim baPfx()         As Byte
    Dim uBlob           As CRYPT_BLOB_DATA
    Dim hPfxStore       As Long
    Dim pCertContext    As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    
    baPfx = ReadBinaryFile(sPfxFile)
    If UBound(baPfx) < 0 Then
        GoTo QH
    End If
    uBlob.cbData = UBound(baPfx) + 1
    uBlob.pbData = VarPtr(baPfx(0))
    hPfxStore = PFXImportCertStore(uBlob, StrPtr(sPassword), CRYPT_EXPORTABLE)
    If hPfxStore = 0 And Err.LastDllError <> NTE_BAD_ALGID Then
        hPfxStore = PFXImportCertStore(baPfx(0), 0, CRYPT_EXPORTABLE)
    End If
    If hPfxStore = 0 Then
        sApiSource = "PFXImportCertStore"
        hResult = Err.LastDllError
        GoTo QH
    End If
    Do
        pCertContext = CertEnumCertificatesInStore(hPfxStore, pCertContext)
        If pCertContext = 0 Then
            Exit Do
        End If
        If PkiAppendCertContext(pCertContext, cCerts, baPrivKey) Then
            '--- success
            PkiPkcs12ImportCertificates = True
        End If
    Loop
QH:
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Public Function PkiGenerSelfSignedCertificate(cCerts As Collection, baPrivKey() As Byte, Optional ByVal Subject As String) As Boolean
    Const FUNC_NAME     As String = "PkiGenerSelfSignedCertificate"
    Dim hProv           As Long
    Dim hKey            As Long
    Dim sName           As String
    Dim baName()        As Byte
    Dim lSize           As Long
    Dim uName           As CRYPT_BLOB_DATA
    Dim uExpire         As SYSTEMTIME
    Dim uInfo           As CRYPT_KEY_PROV_INFO
    Dim pCertContext    As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    
    Call CryptAcquireContext(0, 0, 0, PROV_RSA_FULL, CRYPT_DELETEKEYSET)
    If CryptAcquireContext(hProv, 0, 0, PROV_RSA_FULL, CRYPT_NEWKEYSET) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptAcquireContext"
        GoTo QH
    End If
    If CryptGenKey(hProv, AT_SIGNATURE, RSA1024BIT_KEY Or CRYPT_EXPORTABLE, hKey) = 0 Then
        GoTo QH
    End If
    If Left$(Subject, 3) <> "CN=" Then
        If LenB(Subject) = 0 Then
            Subject = LCase$(Environ$("COMPUTERNAME") & IIf(LenB(Environ$("USERDNSDOMAIN")) <> 0, "." & Environ$("USERDNSDOMAIN"), vbNullString))
        End If
        sName = "CN=""" & Replace(Subject, """", """""") & """" & ",OU=""" & Replace(Environ$("USERDOMAIN") & "\" & Environ$("USERNAME"), """", """""") & """,O=""VbAsyncSocket Self-Signed Certificate"""
    Else
        sName = Subject
    End If
    If CertStrToName(X509_ASN_ENCODING, StrPtr(sName), CERT_OID_NAME_STR, 0, ByVal 0, lSize, 0) = 0 Then
        GoTo QH
    End If
    ReDim baName(0 To lSize - 1) As Byte
    Debug.Assert RedimStats(FUNC_NAME & ".baName", UBound(baName) + 1)
    If CertStrToName(X509_ASN_ENCODING, StrPtr(sName), CERT_OID_NAME_STR, 0, baName(0), lSize, 0) = 0 Then
        GoTo QH
    End If
    With uName
        .cbData = lSize
        .pbData = VarPtr(baName(0))
    End With
    Call GetSystemTime(uExpire)
    uExpire.wYear = uExpire.wYear + 1
    With uInfo
        .dwProvType = PROV_RSA_FULL
        .dwKeySpec = AT_SIGNATURE
    End With
    pCertContext = CertCreateSelfSignCertificate(hProv, uName, 0, uInfo, 0, ByVal 0, uExpire, 0)
    If PkiAppendCertContext(pCertContext, cCerts, baPrivKey) Then
        '--- success
        PkiGenerSelfSignedCertificate = True
    End If
QH:
    If hKey <> 0 Then
        Call CryptDestroyKey(hKey)
    End If
    If hProv <> 0 Then
        Call CryptReleaseContext(hProv, 0)
        Call CryptAcquireContext(0, 0, 0, PROV_RSA_FULL, CRYPT_DELETEKEYSET)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

'= private ===============================================================

Private Function PkiAppendCertContext(ByVal pCertContext As Long, cCerts As Collection, baPrivKey() As Byte) As Boolean
    Dim uCertContext    As CERT_CONTEXT
    Dim baBuffer()      As Byte
    
    Call CopyMemory(uCertContext, ByVal pCertContext, Len(uCertContext))
    If uCertContext.cbCertEncoded > 0 Then
        ReDim baBuffer(0 To uCertContext.cbCertEncoded - 1) As Byte
        Debug.Assert RedimStats("PkiAppendCertContext.baBuffer", UBound(baBuffer) + 1)
        Call CopyMemory(baBuffer(0), ByVal uCertContext.pbCertEncoded, uCertContext.cbCertEncoded)
        If cCerts Is Nothing Then
            Set cCerts = New Collection
        End If
        cCerts.Add baBuffer
    End If
    If PkiExportPrivateKey(pCertContext, baPrivKey) Then
        If cCerts.Count > 1 Then
            '--- move certificate w/ private key to the beginning of the collection
            baBuffer = cCerts.Item(cCerts.Count)
            cCerts.Remove cCerts.Count
            cCerts.Add baBuffer, Before:=1
        End If
        '--- success
        PkiAppendCertContext = True
    End If
End Function

Private Function PkiExportPrivateKey(ByVal pCertContext As Long, baPrivKey() As Byte) As Boolean
    Const FUNC_NAME     As String = "PkiExportPrivateKey"
    Dim dwFlags         As Long
    Dim hProvOrKey      As Long
    Dim lKeySpec        As Long
    Dim lFree           As Long
    Dim hCngKey         As Long
    Dim hNewKey         As Long
    Dim lSize           As Long
    Dim baBuffer()      As Byte
    Dim uKeyInfo        As CRYPT_KEY_PROV_INFO
    Dim hProv           As Long
    Dim hKey            As Long
    Dim lMagic          As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    
    '--- note: this function allows using CRYPT_ACQUIRE_PREFER_NCRYPT_KEY_FLAG too for key export w/ all CNG API calls
    dwFlags = CRYPT_ACQUIRE_CACHE_FLAG Or CRYPT_ACQUIRE_SILENT_FLAG Or CRYPT_ACQUIRE_ALLOW_NCRYPT_KEY_FLAG
    If CryptAcquireCertificatePrivateKey(pCertContext, dwFlags, 0, hProvOrKey, lKeySpec, lFree) = 0 Then
        GoTo QH
    End If
    If lKeySpec < 0 Then
        hCngKey = hProvOrKey: hProvOrKey = 0
        hNewKey = PkiCloneKeyWithExportPolicy(hCngKey, NCRYPT_ALLOW_EXPORT_FLAG Or NCRYPT_ALLOW_PLAINTEXT_EXPORT_FLAG)
        hResult = NCryptExportKey(hNewKey, 0, StrPtr("PRIVATEBLOB"), ByVal 0, ByVal 0, 0, lSize, 0)
        If hResult < 0 Then
            sApiSource = "NCryptExportKey(PRIVATEBLOB)"
            GoTo QH
        End If
        ReDim baBuffer(0 To lSize - 1) As Byte
        Debug.Assert RedimStats(FUNC_NAME & ".baBuffer", UBound(baBuffer) + 1)
        hResult = NCryptExportKey(hNewKey, 0, StrPtr("PRIVATEBLOB"), ByVal 0, baBuffer(0), UBound(baBuffer) + 1, lSize, 0)
        If hResult < 0 Then
            sApiSource = "NCryptExportKey(PRIVATEBLOB)#2"
            GoTo QH
        End If
        Call CopyMemory(lMagic, baBuffer(0), 4)
        Select Case lMagic
        Case BCRYPT_RSAPRIVATE_MAGIC
            hResult = NCryptExportKey(hNewKey, 0, StrPtr("RSAFULLPRIVATEBLOB"), ByVal 0, ByVal 0, 0, lSize, 0)
            If hResult < 0 Then
                sApiSource = "NCryptExportKey(RSAFULLPRIVATEBLOB)"
                GoTo QH
            End If
            ReDim baBuffer(0 To lSize - 1) As Byte
            Debug.Assert RedimStats(FUNC_NAME & ".baBuffer", UBound(baBuffer) + 1)
            hResult = NCryptExportKey(hNewKey, 0, StrPtr("RSAFULLPRIVATEBLOB"), ByVal 0, baBuffer(0), UBound(baBuffer) + 1, lSize, 0)
            If hResult < 0 Then
                sApiSource = "NCryptExportKey(RSAFULLPRIVATEBLOB)#2"
                GoTo QH
            End If
            baPrivKey = PkiExportRsaPrivateKey(baBuffer, CNG_RSA_PRIVATE_KEY_BLOB)
        Case BCRYPT_ECDH_PRIVATE_P256_MAGIC, BCRYPT_ECDH_PRIVATE_P384_MAGIC, BCRYPT_ECDH_PRIVATE_P521_MAGIC
            Call CopyMemory(lSize, baBuffer(4), 4)
            Debug.Assert 8 + 3 * lSize <= UBound(baBuffer) + 1
            Call CopyMemory(baBuffer(0), baBuffer(8 + 2 * lSize), lSize)
            ReDim Preserve baBuffer(0 To lSize - 1) As Byte
            Debug.Assert RedimStats(FUNC_NAME & ".baBuffer", UBound(baBuffer) + 1)
            baPrivKey = PkiExportEccPrivateKey(baBuffer, lMagic)
        Case Else
            Debug.Print "Unknown CNG private key magic (0x" & Hex$(lMagic) & ")"
        End Select
    Else
        If CertGetCertificateContextProperty(pCertContext, CERT_KEY_PROV_INFO_PROP_ID, ByVal 0, lSize) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CertGetCertificateContextProperty(CERT_KEY_PROV_INFO_PROP_ID)"
            GoTo QH
        End If
        ReDim baBuffer(0 To lSize - 1) As Byte
        Debug.Assert RedimStats(FUNC_NAME & ".baBuffer", UBound(baBuffer) + 1)
        If CertGetCertificateContextProperty(pCertContext, CERT_KEY_PROV_INFO_PROP_ID, baBuffer(0), lSize) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CertGetCertificateContextProperty(CERT_KEY_PROV_INFO_PROP_ID)#2"
            GoTo QH
        End If
        Call CopyMemory(uKeyInfo, baBuffer(0), Len(uKeyInfo))
        If CryptAcquireContext(hProv, uKeyInfo.pwszContainerName, uKeyInfo.pwszProvName, uKeyInfo.dwProvType, uKeyInfo.dwFlags) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptAcquireContext"
            GoTo QH
        End If
        If CryptGetUserKey(hProv, uKeyInfo.dwKeySpec, hKey) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptGetUserKey"
            GoTo QH
        End If
        If CryptExportKey(hKey, 0, PRIVATEKEYBLOB, 0, ByVal 0, lSize) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptExportKey(PRIVATEKEYBLOB)"
            GoTo QH
        End If
        ReDim baBuffer(0 To lSize - 1) As Byte
        Debug.Assert RedimStats(FUNC_NAME & ".baBuffer", UBound(baBuffer) + 1)
        If CryptExportKey(hKey, 0, PRIVATEKEYBLOB, 0, baBuffer(0), lSize) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptExportKey(PRIVATEKEYBLOB)#2"
            GoTo QH
        End If
        Call CopyMemory(lMagic, baBuffer(8), 4)
        Select Case lMagic
        Case BCRYPT_RSAPRIVATE_MAGIC
            baPrivKey = PkiExportRsaPrivateKey(baBuffer, PKCS_RSA_PRIVATE_KEY)
        Case BCRYPT_ECDH_PRIVATE_P256_MAGIC, BCRYPT_ECDH_PRIVATE_P384_MAGIC, BCRYPT_ECDH_PRIVATE_P521_MAGIC
            baPrivKey = PkiExportEccPrivateKey(baBuffer, lMagic)
        Case Else
            Debug.Print "Unknown CAPI private key magic (0x" & Hex$(lMagic) & ")"
        End Select
    End If
    '--- success
    PkiExportPrivateKey = True
QH:
    If hKey <> 0 Then
        Call CryptDestroyKey(hKey)
    End If
    If hProv <> 0 Then
        Call CryptReleaseContext(hProv, 0)
    End If
    If hProvOrKey <> 0 And lFree <> 0 Then
        Call CryptReleaseContext(hProvOrKey, 0)
    End If
    If hCngKey <> 0 And lFree <> 0 Then
        Call NCryptFreeObject(hCngKey)
    End If
    If hNewKey <> 0 Then
        Call NCryptFreeObject(hNewKey)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Private Function PkiExportRsaPrivateKey(baPrivBlob() As Byte, ByVal lStructType As Long) As Byte()
    Const FUNC_NAME     As String = "PkiExportRsaPrivateKey"
    Dim baRetVal()      As Byte
    Dim baRsaPrivKey()  As Byte
    Dim uPrivKey        As CRYPT_PRIVATE_KEY_INFO
    Dim lSize           As Long
    Dim sObjId          As String
    Dim hResult         As Long
    Dim sApiSource      As String
    
    If CryptEncodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, lStructType, baPrivBlob(0), 0, 0, ByVal 0, lSize) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptEncodeObjectEx"
        GoTo QH
    End If
    ReDim baRsaPrivKey(0 To lSize - 1)
    Debug.Assert RedimStats(FUNC_NAME & ".baRsaPrivKey", UBound(baRsaPrivKey) + 1)
    If CryptEncodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, lStructType, baPrivBlob(0), 0, 0, baRsaPrivKey(0), lSize) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptEncodeObjectEx#2"
        GoTo QH
    End If
    sObjId = StrConv(szOID_RSA_RSA, vbFromUnicode)
    With uPrivKey
        .Algorithm.pszObjId = StrPtr(sObjId)
        .PrivateKey.pbData = VarPtr(baRsaPrivKey(0))
        .PrivateKey.cbData = UBound(baRsaPrivKey) + 1
    End With
    If CryptEncodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_PRIVATE_KEY_INFO, uPrivKey, 0, 0, ByVal 0, lSize) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptEncodeObjectEx(PKCS_PRIVATE_KEY_INFO)"
        GoTo QH
    End If
    ReDim baRetVal(0 To lSize - 1)
    Debug.Assert RedimStats(FUNC_NAME & ".baRetVal", UBound(baRetVal) + 1)
    If CryptEncodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_PRIVATE_KEY_INFO, uPrivKey, 0, 0, baRetVal(0), lSize) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptEncodeObjectEx(PKCS_PRIVATE_KEY_INFO)#2"
        GoTo QH
    End If
    PkiExportRsaPrivateKey = baRetVal
QH:
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Private Function PkiExportEccPrivateKey(baPrivBlob() As Byte, ByVal lMagic As Long) As Byte()
    Const FUNC_NAME     As String = "PkiExportEccPrivateKey"
    Dim baRetVal()      As Byte
    Dim sObjId          As String
    Dim uEccPrivKey     As CRYPT_ECC_PRIVATE_KEY_INFO
    Dim lSize           As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    
    sObjId = StrConv(Switch(lMagic = BCRYPT_ECDH_PRIVATE_P521_MAGIC, szOID_ECC_CURVE_P521, _
                            lMagic = BCRYPT_ECDH_PRIVATE_P384_MAGIC, szOID_ECC_CURVE_P384, _
                            True, szOID_ECC_CURVE_P256), vbFromUnicode)
    With uEccPrivKey
        .dwVersion = 1
        .PrivateKey.pbData = VarPtr(baPrivBlob(0))
        .PrivateKey.cbData = UBound(baPrivBlob) + 1
        .szCurveOid = StrPtr(sObjId)
    End With
    If CryptEncodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, X509_ECC_PRIVATE_KEY, uEccPrivKey, 0, 0, ByVal 0, lSize) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptEncodeObjectEx(X509_ECC_PRIVATE_KEY)"
        GoTo QH
    End If
    ReDim baRetVal(0 To lSize - 1)
    Debug.Assert RedimStats(FUNC_NAME & ".baRetVal", UBound(baRetVal) + 1)
    If CryptEncodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, X509_ECC_PRIVATE_KEY, uEccPrivKey, 0, 0, baRetVal(0), lSize) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptEncodeObjectEx(X509_ECC_PRIVATE_KEY)#2"
        GoTo QH
    End If
    PkiExportEccPrivateKey = baRetVal
QH:
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Private Function PkiCloneKeyWithExportPolicy(ByVal hKey As Long, ByVal lPolicy As Long) As Long
    Const FUNC_NAME     As String = "PkiCloneKeyWithExportPolicy"
    Const STR_PASSWORD  As String = "0000"
    Dim baPkcs8()       As Byte
    Dim uParams         As NCryptBufferDesc
    Dim sSecret         As String
    Dim sObjId          As String
    Dim uPbeParams      As CRYPT_PKCS12_PBE_PARAMS
    Dim lSize           As Long
    Dim hProv           As Long
    Dim sKeyName        As String
    Dim hRetVal         As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    
    '--- export PKCS#8 password protected blob
    ReDim uParams.Buffers(0 To 2) As NCryptBuffer
    Debug.Assert RedimStats(FUNC_NAME & ".uParams.Buffers", 0)
    uParams.cBuffers = UBound(uParams.Buffers) + 1
    uParams.pBuffers = VarPtr(uParams.Buffers(0))
    sSecret = STR_PASSWORD
    With uParams.Buffers(0)
        .BufferType = NCRYPTBUFFER_PKCS_SECRET
        .pvBuffer = StrPtr(sSecret)
        .cbBuffer = LenB(sSecret) + 2
    End With
    sObjId = StrConv(szOID_PKCS_12_pbeWithSHA1And3KeyTripleDES, vbFromUnicode)
    With uParams.Buffers(1)
        .BufferType = NCRYPTBUFFER_PKCS_ALG_OID
        .pvBuffer = StrPtr(sObjId)
        .cbBuffer = LenB(sObjId) + 1
    End With
    uPbeParams.cbSalt = 8
    uPbeParams.iIterations = 2048
    With uParams.Buffers(2)
        .BufferType = NCRYPTBUFFER_PKCS_ALG_PARAM
        .pvBuffer = VarPtr(uPbeParams)
        .cbBuffer = 8 + uPbeParams.cbSalt
    End With
    hResult = NCryptExportKey(hKey, 0, StrPtr("PKCS8_PRIVATEKEY"), uParams, ByVal 0, 0, lSize, 0)
    If hResult < 0 Then
        sApiSource = "NCryptExportKey(PKCS8_PRIVATEKEY)"
        GoTo QH
    End If
    ReDim baPkcs8(0 To lSize - 1) As Byte
    Debug.Assert RedimStats(FUNC_NAME & ".baPkcs8", UBound(baPkcs8) + 1)
    hResult = NCryptExportKey(hKey, 0, StrPtr("PKCS8_PRIVATEKEY"), uParams, baPkcs8(0), UBound(baPkcs8) + 1, lSize, 0)
    If hResult < 0 Then
        sApiSource = "NCryptExportKey(PKCS8_PRIVATEKEY)#2"
        GoTo QH
    End If
    '--- retrieve more key props
    hResult = NCryptGetProperty(hKey, StrPtr("Provider Handle"), hProv, 4, lSize, 0)
    If hResult < 0 Then
        sApiSource = "NCryptGetProperty(Provider Handle)"
        GoTo QH
    End If
    hResult = NCryptGetProperty(hKey, StrPtr("Name"), ByVal 0, 0, lSize, 0)
    If hResult < 0 Then
        sApiSource = "NCryptGetProperty(Name)"
        GoTo QH
    End If
    ReDim baBuffer(0 To lSize - 1) As Byte
    Debug.Assert RedimStats(FUNC_NAME & ".baBuffer", UBound(baBuffer) + 1)
    hResult = NCryptGetProperty(hKey, StrPtr("Name"), baBuffer(0), UBound(baBuffer) + 1, lSize, 0)
    If hResult < 0 Then
        sApiSource = "NCryptGetProperty(Name)#2"
        GoTo QH
    End If
    '--- remove trailing terminating zero too
    sKeyName = Replace(CStr(baBuffer), vbNullChar, vbNullString)
    '--- import PKCS#8 blob and set Export Policy before finalizing
    ReDim uParams.Buffers(0 To 1) As NCryptBuffer
    Debug.Assert RedimStats(FUNC_NAME & ".uParams.Buffers", 0)
    uParams.cBuffers = UBound(uParams.Buffers) + 1
    uParams.pBuffers = VarPtr(uParams.Buffers(0))
    sSecret = STR_PASSWORD
    With uParams.Buffers(0)
        .BufferType = NCRYPTBUFFER_PKCS_SECRET
        .pvBuffer = StrPtr(sSecret)
        .cbBuffer = LenB(sSecret) + 2
    End With
    With uParams.Buffers(1)
        .BufferType = NCRYPTBUFFER_PKCS_KEY_NAME
        .pvBuffer = StrPtr(sKeyName)
        .cbBuffer = LenB(sKeyName) + 2
    End With
    hResult = NCryptImportKey(hProv, 0, StrPtr("PKCS8_PRIVATEKEY"), uParams, hRetVal, baPkcs8(0), UBound(baPkcs8) + 1, NCRYPT_OVERWRITE_KEY_FLAG Or NCRYPT_DO_NOT_FINALIZE_FLAG)
    If hResult < 0 Then
        sApiSource = "NCryptImportKey(PKCS8_PRIVATEKEY)"
        GoTo QH
    End If
    hResult = NCryptSetProperty(hRetVal, StrPtr("Export Policy"), lPolicy, 4, NCRYPT_PERSIST_FLAG)
    If hResult < 0 Then
        sApiSource = "NCryptSetProperty(Export Policy)"
        GoTo QH
    End If
    hResult = NCryptFinalizeKey(hRetVal, 0)
    If hResult < 0 Then
        sApiSource = "NCryptFinalizeKey"
        GoTo QH
    End If
    PkiCloneKeyWithExportPolicy = hRetVal
QH:
    If hProv <> 0 Then
        Call NCryptFreeObject(hProv)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Public Function PkiCertChainValidate(sRemoteHostName As String, cCerts As Collection, ByVal hRootStore As Long, sError As String) As Boolean
    Const FUNC_NAME     As String = "PkiCertChainValidate"
    Dim hCertStore      As Long
    Dim lIdx            As Long
    Dim baCert()        As Byte
    Dim pCertContext    As Long
    Dim pChainContext   As Long
    Dim uChain          As CERT_CHAIN_CONTEXT
    Dim lPtr            As Long
    Dim dwErrorStatus   As Long
    Dim uChainParams    As CERT_CHAIN_PARA
    Dim uInfo           As CERT_INFO
    Dim uExtension      As CERT_EXTENSION
    Dim lAltInfoPtr     As Long
    Dim uAltInfo        As CERT_ALT_NAME_INFO
    Dim uEntry          As CERT_ALT_NAME_ENTRY
    Dim sDnsName        As String
    Dim bValidName      As Boolean
    Dim uEngineConfig   As CERT_CHAIN_ENGINE_CONFIG
    Dim hChainEngine    As Long
    Dim uChainElem      As CERT_CHAIN_ELEMENT
    Dim pExistContext   As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    
    '--- load server X.509 certificates to an in-memory certificate store
    hCertStore = CertOpenStore(CERT_STORE_PROV_MEMORY, 0, 0, CERT_STORE_CREATE_NEW_FLAG, 0)
    If hCertStore = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CertOpenStore"
        GoTo QH
    End If
    For lIdx = 1 To cCerts.Count
        baCert = cCerts.Item(lIdx)
        If CertAddEncodedCertificateToStore(hCertStore, X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, baCert(0), UBound(baCert) + 1, CERT_STORE_ADD_USE_EXISTING, 0) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CertAddEncodedCertificateToStore"
            GoTo QH
        End If
    Next
    '--- search remote host FQDN in any X.509 certificate's "Subject Alternative Name" list of DNS names (incl. wildcards)
    Do
        pCertContext = CertEnumCertificatesInStore(hCertStore, pCertContext)
        If pCertContext = 0 Then
            sError = Replace(ERR_NO_MATCHING_ALT_NAME, "%1", sRemoteHostName)
            GoTo QH
        End If
        Call CopyMemory(lPtr, ByVal UnsignedAdd(pCertContext, 12), 4)               '--- dereference pCertContext->pCertInfo->cExtension
        Call CopyMemory(uInfo, ByVal lPtr, Len(uInfo))
        lPtr = CertFindExtension(szOID_SUBJECT_ALT_NAME2, uInfo.cExtension, uInfo.rgExtension)
        If lPtr <> 0 Then
            Call CopyMemory(uExtension, ByVal lPtr, Len(uExtension))
            If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, szOID_SUBJECT_ALT_NAME2, ByVal uExtension.Value.pbData, _
                        uExtension.Value.cbData, CRYPT_DECODE_ALLOC_FLAG Or CRYPT_DECODE_NOCOPY_FLAG, 0, lAltInfoPtr, 0) = 0 Then
                hResult = Err.LastDllError
                sApiSource = "CryptDecodeObjectEx(szOID_SUBJECT_ALT_NAME2)"
                GoTo QH
            End If
            Call CopyMemory(uAltInfo, ByVal lAltInfoPtr, Len(uAltInfo))
            For lIdx = 0 To uAltInfo.cAltEntry - 1
                lPtr = UnsignedAdd(uAltInfo.rgAltEntry, lIdx * Len(uEntry))         '--- dereference lAltInfoPtr->rgAltEntry[lidx].dwAltNameChoice
                Call CopyMemory(uEntry, ByVal lPtr, Len(uEntry))
                If uEntry.dwAltNameChoice = CERT_ALT_NAME_DNS_NAME Then
                    sDnsName = LCase$(pvToStringW(uEntry.pwszDNSName))
                    If Left$(sDnsName, 1) = "*" Then
                        If LCase$(sRemoteHostName) Like sDnsName And Not LCase$(sRemoteHostName) Like "*." & sDnsName Then
                            bValidName = True
                            Exit Do
                        End If
                    Else
                        If LCase$(sRemoteHostName) = sDnsName Then
                            bValidName = True
                            Exit Do
                        End If
                    End If
                End If
            Next
            Call LocalFree(lAltInfoPtr)
            lAltInfoPtr = 0
        End If
    Loop
    '--- build custom chain engine that trusts the additional root CA certificates if provided
    If hRootStore <> 0 Then
        uEngineConfig.cbSize = Len(uEngineConfig) - IIf(OsVersion < ucsOsvWin7, 12, 0)
        uEngineConfig.cAdditionalStore = 1
        uEngineConfig.rghAdditionalStore = VarPtr(hRootStore)
        If CertCreateCertificateChainEngine(uEngineConfig, hChainEngine) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CertCreateCertificateChainEngine"
            GoTo QH
        End If
    End If
    '--- for the matched server certificate try to build a chain of certificates from the ones in the in-memory certificate store
    '---    and check this chain for revokation, expiry or missing link to a trust anchor
    uChainParams.cbSize = Len(uChainParams)
    If CertGetCertificateChain(hChainEngine, pCertContext, 0, hCertStore, uChainParams, 0, 0, pChainContext) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CertGetCertificateChain"
        GoTo QH
    End If
    Call CopyMemory(uChain, ByVal pChainContext, Len(uChain))       '--- dereference pChainContext->rgpChain[0]->TrustStatus.dwErrorStatus
    Call CopyMemory(lPtr, ByVal uChain.rgElem, 4)
    Call CopyMemory(uChain, ByVal lPtr, Len(uChain))
    dwErrorStatus = uChain.TrustStatus.dwErrorStatus And Not CERT_TRUST_IS_NOT_TIME_NESTED
    If hRootStore <> 0 And uChain.cElems > 0 Then
        '--- check if the last certificate in the chain is from our custom hRootStore and remove untrusted flags from status
        Call CopyMemory(lPtr, ByVal UnsignedAdd(uChain.rgElem, (uChain.cElems - 1) * 4), 4)
        Call CopyMemory(uChainElem, ByVal lPtr, Len(uChainElem))
        pExistContext = CertFindCertificateInStore(hRootStore, X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, 0, CERT_FIND_EXISTING, ByVal uChainElem.pCertContext, 0)
        If pExistContext <> 0 Then
            Call CertFreeCertificateContext(pExistContext)
            pExistContext = 0
            dwErrorStatus = dwErrorStatus And Not CERT_TRUST_IS_UNTRUSTED_ROOT And Not CERT_TRUST_IS_NOT_SIGNATURE_VALID
        End If
    End If
    If dwErrorStatus <> 0 Then
        If (dwErrorStatus And CERT_TRUST_IS_REVOKED) <> 0 Then
            sError = ERR_TRUST_IS_REVOKED
        ElseIf (dwErrorStatus And CERT_TRUST_IS_PARTIAL_CHAIN) <> 0 Then
            sError = ERR_TRUST_IS_PARTIAL_CHAIN
        ElseIf (dwErrorStatus And CERT_TRUST_IS_UNTRUSTED_ROOT) <> 0 Then
            sError = ERR_TRUST_IS_UNTRUSTED_ROOT
        ElseIf (dwErrorStatus And CERT_TRUST_IS_NOT_TIME_VALID) <> 0 Then
            sError = ERR_TRUST_IS_NOT_TIME_VALID
        ElseIf (dwErrorStatus And CERT_TRUST_REVOCATION_STATUS_UNKNOWN) <> 0 Then
            sError = ERR_TRUST_REVOCATION_STATUS_UNKNOWN
        Else
            sError = "CertGetCertificateChain error mask: 0x" & Hex$(dwErrorStatus)
        End If
        GoTo QH
    End If
    '--- success
    PkiCertChainValidate = True
QH:
    If pChainContext <> 0 Then
        Call CertFreeCertificateChain(pChainContext)
    End If
    If pCertContext <> 0 Then
        Call CertFreeCertificateContext(pCertContext)
    End If
    If hCertStore <> 0 Then
        Call CertCloseStore(hCertStore, 0)
    End If
    If hChainEngine <> 0 Then
        Call CertFreeCertificateChainEngine(hChainEngine)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), FUNC_NAME & "." & sApiSource
    End If
End Function

Private Function pvToString(ByVal lPtr As Long) As String
    If lPtr <> 0 Then
        pvToString = String$(lstrlen(lPtr), 0)
        Call CopyMemory(ByVal pvToString, ByVal lPtr, Len(pvToString))
    End If
End Function

Private Function pvToStringW(ByVal lPtr As Long) As String
    If lPtr Then
        pvToStringW = String$(lstrlenW(lPtr), 0)
        Call CopyMemory(ByVal StrPtr(pvToStringW), ByVal lPtr, LenB(pvToStringW))
    End If
End Function
