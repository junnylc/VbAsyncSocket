Attribute VB_Name = "mdTlsCrypto"
'=========================================================================
'
' Elliptic-curve cryptography thunks based on the following sources
'
'  1. https://github.com/esxgx/easy-ecc by Kenneth MacKay
'     BSD 2-clause license
'
'  2. https://github.com/ctz/cifra by Joseph Birr-Pixton
'     CC0 1.0 Universal license
'
'=========================================================================
Option Explicit
DefObj A-Z

#Const ImplUseLibSodium = (ASYNCSOCKET_USE_LIBSODIUM <> 0)
#Const ImplUseBCrypt = False

'=========================================================================
' API
'=========================================================================

Private Const TLS_SIGNATURE_RSA_PKCS1_SHA1              As Long = &H201
Private Const TLS_SIGNATURE_RSA_PKCS1_SHA256            As Long = &H401
Private Const TLS_SIGNATURE_RSA_PSS_RSAE_SHA256         As Long = &H804
Private Const TLS_SIGNATURE_RSA_PSS_RSAE_SHA384         As Long = &H805
'--- for CryptAcquireContext
Private Const PROV_RSA_FULL                             As Long = 1
Private Const PROV_RSA_AES                              As Long = 24
Private Const CRYPT_VERIFYCONTEXT                       As Long = &HF0000000
'--- for CryptDecodeObjectEx
Private Const X509_ASN_ENCODING                         As Long = 1
Private Const PKCS_7_ASN_ENCODING                       As Long = &H10000
Private Const X509_PUBLIC_KEY_INFO                      As Long = 8
Private Const PKCS_RSA_PRIVATE_KEY                      As Long = 43
Private Const PKCS_PRIVATE_KEY_INFO                     As Long = 44
Private Const CRYPT_DECODE_ALLOC_FLAG                   As Long = &H8000
'--- for CryptCreateHash
Private Const CALG_SHA1                                 As Long = &H8004&
Private Const CALG_SHA_256                              As Long = &H800C&
'--- for CryptSignHash
Private Const AT_KEYEXCHANGE                            As Long = 1
Private Const MAX_RSA_KEY                               As Long = 8192     '--- in bits
'--- for CryptVerifySignature
Private Const NTE_BAD_SIGNATURE                         As Long = &H80090006
'--- for BCryptSignHash
Private Const BCRYPT_PAD_PSS                            As Long = 8
'--- for BCryptVerifySignature
Private Const STATUS_INVALID_SIGNATURE                  As Long = &HC000A000
Private Const ERROR_INVALID_DATA                        As Long = &HC000000D
'--- for thunks
Private Const MEM_COMMIT                                As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE                    As Long = &H40
#If ImplUseBCrypt Then
    Private Const BCRYPT_SECP256R1_PARTSZ               As Long = 32
    Private Const BCRYPT_SECP256R1_PRIVATE_KEYSZ        As Long = BCRYPT_SECP256R1_PARTSZ * 3
    Private Const BCRYPT_SECP256R1_COMPRESSED_PUBLIC_KEYSZ As Long = 1 + BCRYPT_SECP256R1_PARTSZ
    Private Const BCRYPT_SECP256R1_UNCOMPRESSED_PUBLIC_KEYSZ As Long = 1 + BCRYPT_SECP256R1_PARTSZ * 2
    Private Const BCRYPT_SECP256R1_TAG_COMPRESSED_POS   As Long = 2
    Private Const BCRYPT_SECP256R1_TAG_COMPRESSED_NEG   As Long = 3
    Private Const BCRYPT_SECP256R1_TAG_UNCOMPRESSED     As Long = 4
    Private Const BCRYPT_ECDH_PUBLIC_P256_MAGIC         As Long = &H314B4345
    Private Const BCRYPT_ECDH_PRIVATE_P256_MAGIC        As Long = &H324B4345
#End If

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextW" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGenRandom Lib "advapi32" (ByVal hProv As Long, ByVal dwLen As Long, ByVal pbBuffer As Long) As Long
Private Declare Function CryptImportKey Lib "advapi32" (ByVal hProv As Long, pbData As Any, ByVal dwDataLen As Long, ByVal hPubKey As Long, ByVal dwFlags As Long, phKey As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32" (ByVal hProv As Long, ByVal AlgId As Long, ByVal hKey As Long, ByVal dwFlags As Long, phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32" (ByVal hHash As Long, pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32" (ByVal hHash As Long) As Long
Private Declare Function CryptSignHash Lib "advapi32" Alias "CryptSignHashA" (ByVal hHash As Long, ByVal dwKeySpec As Long, ByVal szDescription As Long, ByVal dwFlags As Long, pbSignature As Any, pdwSigLen As Long) As Long
Private Declare Function CryptVerifySignature Lib "advapi32" Alias "CryptVerifySignatureA" (ByVal hHash As Long, pbSignature As Any, ByVal dwSigLen As Long, ByVal hPubKey As Long, ByVal szDescription As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, pbData As Any, pdwDataLen As Long, dwBufLen As Long) As Long
Private Declare Function CryptImportPublicKeyInfo Lib "crypt32" (ByVal hCryptProv As Long, ByVal dwCertEncodingType As Long, pInfo As Any, phKey As Long) As Long
Private Declare Function CryptDecodeObjectEx Lib "crypt32" (ByVal dwCertEncodingType As Long, ByVal lpszStructType As Long, pbEncoded As Any, ByVal cbEncoded As Long, ByVal dwFlags As Long, ByVal pDecodePara As Long, pvStructInfo As Any, pcbStructInfo As Long) As Long
Private Declare Function CryptEncodeObjectEx Lib "crypt32" (ByVal dwCertEncodingType As Long, ByVal lpszStructType As Long, pvStructInfo As Any, ByVal dwFlags As Long, ByVal pEncodePara As Long, pvEncoded As Any, pcbEncoded As Long) As Long
Private Declare Function CertCreateCertificateContext Lib "crypt32" (ByVal dwCertEncodingType As Long, pbCertEncoded As Any, ByVal cbCertEncoded As Long) As Long
Private Declare Function CertFreeCertificateContext Lib "crypt32" (ByVal pCertContext As Long) As Long
Private Declare Function CryptImportPublicKeyInfoEx2 Lib "crypt32" (ByVal dwCertEncodingType As Long, ByVal pInfo As Long, ByVal dwFlags As Long, ByVal pvAuxInfo As Long, phKey As Long) As Long
Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt" (ByRef hAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptImportKeyPair Lib "bcrypt" (ByVal hAlgorithm As Long, ByVal hImportKey As Long, ByVal pszBlobType As Long, ByRef hKey As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDestroyKey Lib "bcrypt" (ByVal hKey As Long) As Long
Private Declare Function BCryptSignHash Lib "bcrypt" (ByVal hKey As Long, pPaddingInfo As Any, pbInput As Any, ByVal cbInput As Long, pbOutput As Any, ByVal cbOutput As Long, pcbResult As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptVerifySignature Lib "bcrypt" (ByVal hKey As Long, pPaddingInfo As Any, pbHash As Any, ByVal cbHash As Long, pbSignature As Any, ByVal cbSignature As Long, ByVal dwFlags As Long) As Long
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
#If ImplUseBCrypt Then
    '--- BCrypt
    Private Declare Function BCryptExportKey Lib "bcrypt" (ByVal hKey As Long, ByVal hExportKey As Long, ByVal pszBlobType As Long, ByVal pbOutput As Long, ByVal cbOutput As Long, ByRef cbResult As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptSecretAgreement Lib "bcrypt" (ByVal hPrivKey As Long, ByVal hPubKey As Long, ByRef phSecret As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptDestroySecret Lib "bcrypt" (ByVal hSecret As Long) As Long
    Private Declare Function BCryptDeriveKey Lib "bcrypt" (ByVal hSharedSecret As Long, ByVal pwszKDF As Long, ByVal pParameterList As Long, ByVal pbDerivedKey As Long, ByVal cbDerivedKey As Long, ByRef pcbResult As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptGenerateKeyPair Lib "bcrypt" (ByVal hAlgorithm As Long, ByRef hKey As Long, ByVal dwLength As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptFinalizeKeyPair Lib "bcrypt" (ByVal hKey As Long, ByVal dwFlags As Long) As Long
#End If

Private Type CRYPT_DER_BLOB
    cbData              As Long
    pbData              As Long
End Type

Private Type BCRYPT_PSS_PADDING_INFO
    pszAlgId            As Long
    cbSalt              As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_GLOB                  As String = "////////////////AAAAAAAAAAAAAAAAAQAAAP////9LYNInPjzOO/awU8ywBh1lvIaYdlW967Pnkzqq2DXGWpbCmNhFOaH0oDPrLYF9A3fyQKRj5ea8+EdCLOHy0Rdr9VG/N2hAtsvOXjFrVzPOKxaeD3xK6+eOm38a/uJC409RJWP8wsq584SeF6et+ua8//////////8AAAAA/////5gvikKRRDdxz/vAtaXbtelbwlY58RHxWaSCP5LVXhyrmKoH2AFbgxK+hTEkw30MVXRdvnL+sd6Apwbcm3Txm8HBaZvkhke+78adwQ/MoQwkbyzpLaqEdErcqbBc2oj5dlJRPphtxjGoyCcDsMd/Wb/zC+DGR5Gn1VFjygZnKSkUhQq3JzghGy78bSxNEw04U1RzCmW7Cmp2LsnCgYUscpKh6L+iS2YaqHCLS8KjUWzHGeiS0SQGmdaFNQ70cKBqEBbBpBkIbDceTHdIJ7W8sDSzDBw5SqrYTk/KnFvzby5o7oKPdG9jpXgUeMiECALHjPr/vpDrbFCk96P5vvJ4ccYirijXmC+KQs1l7yORRDdxLztN7M/7wLW824mBpdu16Ti1SPNbwlY5GdAFtvER8VmbTxmvpII/khiBbdrVXhyrQgIDo5iqB9i+b3BFAVuDEoyy5E6+hTEk4rT/1cN9DFVviXvydF2+crGWFjv+sd6ANRLHJacG3JuUJmnPdPGbwdJK8Z7BaZvk4yVPOIZHvu+11YyLxp3BD2WcrHfMoQwkdQIrWW8s6S2D5KZuqoR0StT7Qb3cqbBctVMRg9qI+Xar32buUlE+mBAytC1txjGoPyH7mMgnA7DkDu++x39Zv8KPqD3zC+DGJacKk0eRp9VvggPgUWPKBnBuDgpnKSkU/C/S" & _
                                                    "RoUKtycmySZcOCEbLu0qxFr8bSxN37OVnRMNOFPeY6+LVHMKZaiydzy7Cmp25q7tRy7JwoE7NYIUhSxykmQD8Uyh6L+iATBCvEtmGqiRl/jQcItLwjC+VAajUWzHGFLv1hnoktEQqWVVJAaZ1iogcVeFNQ70uNG7MnCgahDI0NK4FsGkGVOrQVEIbDcemeuO30x3SCeoSJvhtbywNGNaycWzDBw5y4pB40qq2E5z42N3T8qcW6O4stbzby5o/LLvXe6Cj3RgLxdDb2OleHKr8KEUeMiE7DlkGggCx4woHmMj+v++kOm9gt7rbFCkFXnGsvej+b4rU3Lj8nhxxpxhJurOPifKB8LAIce4htEe6+DN1n3a6njRbu5/T331um8Xcqpn8AammMiixX1jCq4N+b4EmD8RG0ccEzULcRuEfQQj9XfbKJMkx0B7q8oyvL7JFQq+njxMDRCcxGcdQ7ZCPsu+1MVMKn5l/Jwpf1ns+tY6q2/LXxdYR0qMGURsZXhwYW5kIDE2LWJ5dGUgawBleHBhbmQgMzItYnl0ZSBrAAAABQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAPwAAABjfHd78mtvxTABZyv+16t2yoLJffpZR/Ct1KKvnKRywLf9kyY2P/fMNKXl8XHYMRUExyPDGJYFmgcSgOLrJ7J1CYMsGhtuWqBSO9azKeMvhFPRAO0g/LFbasu+OUpMWM/Q76r7Q00zhUX5An9QPJ+oUaNAj5KdOPW8ttohEP/z0s0ME+xfl0QXxKd+PWRdGXNggU/cIiqQiEbuuBTeXgvb4DI6CkkGJFzC06xikZXkeefIN22N1U6pb" & _
                                                    "Fb06mV6rgi6eCUuHKa0xujddB9LvYuKcD61ZkgD9g5hNVe5hsEdnuH4mBFp2Y6Umx6H6c5VKN+MoYkNv+ZCaEGZLQ+wVLsWjQECBAgQIECAGzZSCWrVMDalOL9Ao56B89f7fOM5gpsv/4c0jkNExN7py1R7lDKmwiM97kyVC0L6w04ILqFmKNkksnZboklti9Elcvj2ZIZomBbUpFzMXWW2kmxwSFD97bnaXhVGV6eNnYSQ2KsAjLzTCvfkWAW4s0UG0Cwej8o/DwLBr70DAROKazqREUFPZ9zql/LPzvC05nOWrHQi5601heL5N+gcdd9uR/EacR0pxYlvt2IOqhi+G/xWPkvG0nkgmtvA/njNWvQf3agziAfHMbESEFkngOxfYFF/qRm1Sg0t5Xqfk8mc76DgO02uKvWwyOu7PINTmWEXKwR+unfWJuFpFGNVIQx9AAAAAAA=" ' 1688, 22.4.2020 14:28:07
Private Const STR_THUNK1                As String = "MOENADAuAAAQMQAAcDEAALAxAADgMgAAIBoAAEAdAADAJgAAECcAAPAkAACAJwAAECgAAFAnAABAGQAAABkAAFAOAADQDQAAwYP4DHzyi0UIi0SF0IvlXcIEAMzMzMzM6AAAAABYLWVADQAFAEANAIsAw8zMzMzMzMzMzMzMzMzoAAAAAFgthUANAAUAQA0Aw8zMzMzMzMzMzMzMzMzMzFWL7IPsSFOLXRBT6IBTAACFwA+FKQEAAFaLdQyNRdhXVlDoKVsAAIt9CI1F2FBXjUW4UOjoWgAAjUXYUFDoDlsAAFNWVujWWgAAU1Po/1oAAOhq////UFNXV+jhVgAA6Fz///9QU1NT6NNWAADoTv///1BTV1PoBVsAAFNXV+idWgAA6Dj///9QV1dT6K9WAADoKv///1BTV1fooVYAAGoAV+iJYAAAC8J0IOgQ////UFdX6NhQAABXi/Do0FwAAMHmHwl3HIt1DOsGV+i/XAAAV1PoeFoAAOjj/v//UI1FuFBTU+iXWgAA6NL+//9QjUW4UFNT6IZaAADowf7//1BTjUW4UFDodVoAAI1FuFBXV+gKWgAA6KX+//9QjUXYUFdQ6FlaAABTV+jCXAAAVlPou1wAAI1F2FBW6LFcAABfXluL5V3CDADMzMzMzMzMzFWL7FaLdQhW6DNSAACFwHQXjUYgUOgmUgAAhcB0CrgBAAAAXl3CBAAzwF5dwgQAzFWL7IHsqAAAAFOLXQyNRbhWV1NQ6FdcAACNQyBQiUX4jYV4////UOhEXAAA/3UUjYVY////UI1FmFCNhXj///9QjUW4UOiGAwAAi10QU+g9WwAAg+gCiUUUhcB+Ww8fAFBT6ElfAAALwnUHuAEAAADrAjPAweAFjZ1Y////A9iNTZgD" & _
                                                    "yI21eP///1P32IlN/FED8I19uAP4Vlfo4QEAAFZXU/91/Oj2AAAAi0UUi10QSIlFFIXAf6hqAFPo8F4AAAvCdQWNSAHrAjPJweEFjZ1Y////A9mJTRBTjUWYA8GNvXj///9QK/mNdbgr8VdW6IwBAADoR/3//1CNRZhQjUW4UI1F2FDo9VgAAFeNRdhQUOiKWAAA/3UMjUXYUFDofVgAAOgY/f//UI1F2FBQ6M1UAAD/dfiNRdhQUOhgWAAAVo1F2FBQ6FVYAABXVo1FmANFEFNQ6EYAAACNRdhQjYVY////UI1FmFDoogUAAIt1CI1FmFBW6PVaAACNhVj///9QjUYgUOjlWgAAX15bi+VdwhAAzMzMzMzMzMzMzMzMVYvsg+wgU1ZX6JL8//+LXQiLdRBQU1aNReBQ6EBYAACNReBQUOgGWAAAjUXgUFNT6MtXAACNReBQVlbowFcAAOhb/P//i3UMi30UUFZXV+gMWAAAV41F4FDo0lcAAOg9/P//UFONReBQUOjxVwAA6Cz8//9Qi0UQUI1F4FBQ6N1XAADoGPz//1CLRRBTUFDozFcAAItFEFBWVuhhVwAA6Pz7//9QjUXgUFOLXRBT6K1XAABTV1foRVcAAOjg+///UFZXV+iXVwAAjUXgUFPo/VkAAF9eW4vlXcIQAMzMzMxVi+yD7GBTVlfosvv//4tdCIt9EFBTV41FwFDoYFcAAI1FwFBQ6CZXAACNRcBQU1Po61YAAI1FwFBXV+jgVgAA6Hv7//+LXQyLdRRQU1aNRcBQ6OlSAADoZPv//1BTVlboG1cAAOhW+///UP91CI1F4FdQ6AhXAACNReBQU1PonVYAAOg4+///UFf/dQiNReBQ6KpSAABWV+izVgAA6B77/" & _
                                                    "/9QjUXgUFdX6NJWAADoDfv//1BXi30IjUWgV1DovlYAAI1FoFBWVuhTVgAA6O76//9QU1ZW6KVWAACNRcBQjUWgUOhoVgAA6NP6//9QjUXgUI1FoFBQ6IRWAADov/r//1BXjUWgUI1F4FDocFYAAI1FwFCNReBQUOgCVgAA6J36//9QU41F4FBT6FFWAACNRaBQV+i3WAAAX15bi+VdwhAAzMzMzMzMzMzMzMzMzMxVi+yD7CBWi3UIV1b/dRDojFgAAIt9DFf/dRTogFgAAI1F4FDoV00AAItFGMdF4AEAAADHReQAAAAAhcB0ClCNReBQ6FhYAACNReBQV1bo7QIAAI1F4FBXVuhS+v//jUXgUP91FP91EOjTAgAAX16L5V3CFADMzMzMzMzMzMzMzFOLRCQMi0wkEPfhi9iLRCQI92QkFAPYi0QkCPfhA9NbwhAAzMzMzMzMzMzMzMzMzID5QHMVgPkgcwYPpcLT4MOL0DPAgOEf0+LDM8Az0sPMgPlAcxWA+SBzBg+t0NPqw4vCM9KA4R/T6MMzwDPSw8xVi+yLRRBTVot1CI1IeFeLfQyNVng78XcEO9BzC41PeDvxdzA713IsK/i7EAAAACvwixQ4AxCLTDgEE0gEjUAIiVQw+IlMMPyD6wF15F9eW13CDACL141IEIveK9Ar2Cv+uAQAAACNdiCNSSAPEEHQDxBMN+BmD9TIDxFO4A8QTArgDxBB4GYP1MgPEUwL4IPoAXXSX15bXcIMAMzMzMzMVYvsi1Ucg+wIi0UgVot1CFeLfQwD1xNFEIkWiUYEO0UQdw9yBDvXcwm4AQAAADPJ6w4PV8BmDxNF+ItN/ItF+ANFJF8TTSgDRRSJRgiLxhNNGIlODF6L5V3CJADMzM" & _
                                                    "zMVYvsi1UMi00IiwIxAYtCBDFBBItCCDFBCItCDDFBDF3CCADMzMzMzMzMzMzMzMzMVYvsg+wIi00Ii1UQU1aLAY1ZBMHqAjP2iVUQiV34jQSFBAAAAIlF/FeF0nRCi1UMi30Qg8ICZmYPH4QAAAAAAA+2Sv6NUgQPtkL7weEIC8gPtkL8weEIC8gPtkL9weEIC8iJDLNGO/dy1otF/IvXuQEAAAAz/4lNDDvwD4ONAAAAi8Yrwo0Eg4lFCA8fRAAAi1yz/Dv6dQhBM/+JTQzrBIX/dS3op/f//wWIBAAAwcMIUFPoKEgAAIvY6JH3//+LTQwPtoQIiAUAAMHgGDPY6x2D+gZ2HoP/BHUZ6HD3//8FiAQAAFBT6PRHAACL2ItFCItVEIsIRzPLg8AEi134iUUIiQyzRotNDDt1/HKCX15bi+VdwgwAzMzMzMzMzMzMVYvsg+wgjUXg/3UQUOiuUgAAjUXgUItFCFBQ6HBSAAD/dRCNReBQUOhjUgAAjUXgUItFDFBQ6FVSAACL5V3CDADMzMzMzMzMzMzMzMzMzMxVi+yD7CBTVot1CDPJV4lN7IEEzgAAAQCLBM6DVM4EAItczgQPrNgQwfsQiUXog/kPdRXHRfwBAAAAi9DHRfAAAAAAiV346yIPV8BmDxNF9ItF+IlF8ItF9GYPE0Xgi1XgiUX8i0XkiUX4g/kPjXkBagAbwPfYD6/HK1X8aiWNNMaLRfgbRfBQUuhi/P//i03oA8ET04PoAYPaAAEGi0XsEVYEi3UID6TLEMHhECkMxovPiU3sGVzGBIP5EA+CT////19eW4vlXcIEAMzMzMzMVYvsg+wQi1UMVlcPtgoPtkIBweEIC8gPtkICweEIC8gPtkIDweEIC8gPtkI" & _
                                                    "FiU3wD7ZKBMHhCAvID7ZCBsHhCAvID7ZCB8HhCAvID7ZCCYlN9A+2SgjB4QgLyA+2QgrB4QgLyA+2QgvB4QgLyA+2QgyJTfgPtkoNweAIC8gPtkIOweEIC8gPtkIPweEIC8iJTfyLTQiLOY1xBIvHweAEA/CNRfBWUOjl/P//g+4Qg8f/dC2NRfBQ6AQvAACNRfBQ6JsvAABWjUXwUOjB/P//jUXwUOioLgAAg+4Qg+8BddONRfBQ6NcuAACNRfBQ6G4vAABWjUXwUOiU/P//i3UQi1Xwi8KLTfTB6BiIBovCwegQiEYBi8LB6AiIRgKLwcHoGIhWA4hGBIvBwegQiEYFi8HB6AiIRgaITgeLTfiLwcHoGIhGCIvBwegQiEYJi8HB6AiIRgqITguLTfyLwcHoGIhGDIvBwegQiEYNi8HB6AiIRg5fiE4PXovlXcIMAMzMVYvsg+wQU1ZXi1UMi10ID7YKD7ZCAcHhCI1zBAvID7ZCAsHhCAvID7ZCA8HhCAvID7ZCBYlN8A+2SgTB4QgLyA+2QgbB4QgLyA+2QgfB4QgLyA+2QgmJTfQPtkoIweEIC8gPtkIKweEIC8gPtkILweEIC8gPtkIMiU34D7ZKDcHgCAvID7ZCDsHhCAvID7ZCD8HhCAvIjUXwVlCJTfzobfv//78BAAAAg8YQOTt2LpCNRfBQ6PdDAACNRfBQ6I5CAACNRfBQ6BUvAABWjUXwUOg7+///R4PGEDs7ctONRfBQ6MpDAACNRfBQ6GFCAABWjUXwUOgX+///i3UQi1Xwi8KLTfTB6BiIBovCwegQiEYBi8LB6AiIRgKLwcHoGIhWA4hGBIvBwegQiEYFi8HB6AiIRgaITgeLTfiLwcHoGIhGCIvBwegQiEYJ" & _
                                                    "i8HB6AiIRgqITguLTfyLwcHoGIhGDIvBwegQiEYNi8HB6AiIRg5fiE4PXluL5V3CDADMzMzMVYvsVot1CGj0AAAAagBW6OwtAACLRRCDxAyD+BB0PIP4GHQhg/ggdAb/FdTADQBqIP91DMcGDgAAAFbogPr//15dwgwAahj/dQzHBgwAAABW6Gr6//9eXcIMAGoQ/3UMxwYKAAAAVuhU+v//Xl3CDADMzMzMzMzMzMzMzMzMzMxVi+yB7AABAABW6KHy//++0EsNAIHuAEANAAPw6I/y////dSi5UEoNAMdF9BAAAAD/dSSB6QBADQCJdfgDwYlF/I2FAP///1DoM/////91CI2FAP///2oQ/3UUagz/dSD/dRz/dRj/dRD/dQxQjUX0UOg6DwAAXovlXcIkAMzMzFWL7IHsAAEAAFboIfL//77QSw0Age4AQA0AA/DoD/L///91KLlQSg0Ax0X0EAAAAP91JIHpAEANAIl1+APBiUX8jYUA////UOiz/v//ahD/dQyNhQD/////dQhqDP91IP91HP91GP91FP91EFCNRfRQ6HoQAABei+VdwiQAzMzMVYvsUVOLXRgzwIlF/IXbdHGLVRCLTQxWx0UYAQAAAFeLOYvyK/c73g9C84XAdR0PtkUUVlCLRQgDx1DoQCwAAItNDIPEDItF/ItVEIX/dQk78g9ERRiJRfyNBD47wnUX/3UI/3Ug/1Uci00Mi1UQxwEAAAAA6wIBMYtF/CvedaBfXluL5V3CHADMzMzMzMzMVYvsVot1IIvGg+gAdGCD6AEPhKwAAABTg+gBV41FFHRti30oi10kV1NqAVD/dRD/dQz/dQjotgAAAItNGFdTOE0cdC+NRv6LdRBQUVb/dQz/dQjoGP///" & _
                                                    "1dTagGNRRxQVv91DP91COiEAAAAX1teXcIkAI1G/4t1EFBRVv91DP91COjp/v//X1teXcIkAP91KItdEP91JIt9DIt1CGoBUFNXVuhIAAAA/3UojUUc/3UkagFQU1dW6DQAAABfW15dwiQA/3UoikUc/3UkMEUUjUUUagFQ/3UQ/3UM/3UI6A0AAABeXcIkAMzMzMzMzMzMVYvs/3Ugi0UcUFD/dRj/dRT/dRD/dQz/dQjoEQAAAF3CHADMzMzMzMzMzMzMzMzMVYvsi00Mi0UkU4tdFIsRVot1GFeF0nRZhfZ0VYtFEIv+K8I7xg9C+IvCA0UIV1NQ6GsqAACLRQwD3yv3g8QMATiLfRA5OItFJHUp/3UIUIX2dQ3/VSCLTQyLRSSJMesU/1Uci00Mi0UkxwEAAAAA6wOLfRA793IZU1A793UF/1Ug6wP/VRyLRSQr9wPfO/dz54X2dC6LRQyLCIvHK8GL/jvGD0L4i0UIVwPBU1Do7ikAAItFDAPfg8QMATgr94t9EHXVX15bXcIgAMzMzMzMzFWL7ItNHIPsCFeLfRiFyXR2U4tdDFaDOwB1Ef91CP91JP9VIItFEItNHIkDiwOL8YtVECvQO8GJVRgPQvAzwIl1/IX2dC+LXRQr34ld+GaQi3X8jRQ4igwTi1UYA1UIi134MgwCjRQ4QIgKO8Zy4YtdDItNHCkzK84BdRQD/olNHIXJdZFeW1+L5V3CIADMzFWL7OiY7v//ucBYDQCB6QBADQADwYtNCFFQ/3UUjUF0/3UQ/3UMakBQjUE0UOg+////XcIQAMzMzMzMzMzMzMxVi+yD7GyLTRRTVlcPtlkDD7ZBAg+2UQfB4gjB4wgL2A+2QQHB4wgL2A+2AcHjCAvYD7ZBBg" & _
                                                    "vQiV3YweIID7ZBBQvQD7ZBBMHiCAvQD7ZBColV9IlV1A+2UQvB4ggL0A+2QQnB4ggL0A+2QQjB4ggL0A+2QQ6JVfCJVdAPtlEPweIIC9APtkENweIIC9APtkEMi00IweIIC9CJVfgPtkECiVXMD7ZRA8HiCAvQD7ZBAcHiCAvQD7YBweIIC9APtkEGiVXsiVXID7ZRB8HiCAvQD7ZBBcHiCAvQD7ZBBMHiCAvQD7ZBColV6IlVxA+2UQvB4ggL0MHiCA+2QQkL0A+2QQjB4ggL0A+2QQ6JVeSJVcAPtlEPweIIC9APtkENweIIC9APtkEMi00MweIIC9CJVeAPtkECiVW8D7ZRA8HiCAvQD7ZBAcHiCAvQD7YBweIIC9APtkEGiVUIiVW4D7ZRB8HiCAvQD7ZBBcHiCAvQD7ZBBMHiCAvQD7ZBColVFIlVtA+2UQvB4ggL0A+2QQnB4ggL0A+2QQjB4ggL0A+2QQ6JVQyJVbAPtlEPweIIC9APtkENweIIC9APtkEMweIIC9CJVfyJVayLVRAPtkoDD7ZCAsHhCAvID7ZCAcHhCAvID7YCweEIC8iJTdyJTagPtnIHD7ZCBg+2egsPtkoOweYIC/DB5wgPtkIFweYIC/DHRZgKAAAAD7ZCBMHmCAvwD7ZCCgv4iXWkD7ZCCcHnCAv4D7ZCCMHnCAv4D7ZCD8HgCAvBiX2gD7ZKDcHgCAvBD7ZKDItV3MHgCAvBi03siUWc6wOLXRAD2YtNCDPTiV0QwcIQA8qJTQgzTezBwQwD2TPTiV0Qi10IwcIIA9qJVdyLVfQDVegz8oldCDPZwcYQi00UA87BwweJTRQzTejBwQwD0TPyiVX0i1UUwcYIA9aJdeyLdfADdeQz/olVFDPRwcc" & _
                                                    "Qi00MA8/BwgeJTQwzTeTBwQwD8TP+iXXwi3UMwccIA/eJfZSLffgDfeAzx4l1DDPxwcAQi038A8jBxgeJTfwzTeDBwQwD+TPHiX34i338wcAIA/iJffwz+YtNEAPKwccHM8GJTRCLTQzBwBADyIlNDDPKi1UQwcEMA9EzwolVEItVDMHACAPQiVUMM9GLTfQDzsHCB4lN9IlV6ItV3DPRi038wcIQA8qJTfwzzot19MHBDAPxM9aJdfSLdfzBwggD8ol1/DPxi03wA8/BxgeJTfCJdeSLdewz8YtNCMHGEAPOiU0IM8+LffDBwQwD+TP3iX3wi30IwcYIA/6JfQgz+YtN+APLwccHiX3gi32UM/mJTfiLTRTBxxADz4lNFDPLi134wcEMA9kz+4ld+MHHCAF9FItdFDPZi8uJXezBwQeDbZgBi134iU3sD4VA/v//AUWcAV3Mi03YA00QAVWoi1UYiU3Yi13Yi8OLTdQDTfSIGolN1ItN0ANN8MHoCIhCAYvDiU3Qi03sAU3Ii03EA03owegQiEICwesYiFoDi13Ui8OIWgTB6AiIQgWLw4lNxItNwANN5MHoEIhCBolNwItNvANN4AF1pAF9oMHrGIhaB4td0IvDiFoIiU28i024A00IwegIiEIJi8OJTbiLTbQDTRTB6BCIQgrB6xiIWguLXcyLw4lNtItNsANNDIhaDMHoCIhCDYvDiU2wi02sA038wegQiEIOwesYiFoPi13Ii8OJTayIWhDB6AiIQhGLw8HoEIhCEsHrGIhaE4tdxIvDiFoUwegIiEIVi8PB6BCIQhbB6xiIWheLXcCLw4haGMHoCIhCGYvDwegQiEIawesYiFobi128i8OIWhzB6AiIQh2Lw8HoEIhCHsHr" & _
                                                    "GIhaH4tduIvDiFogwegIiEIhi8PB6BCIQiLB6xiIWiOLXbSLw4haJMHoCIhCJYvDwegQiEImwesYiFoni12wi8OIWijB6AiIQimLw8HoEIhCKsHrGIhaK4vZiFosi8PB6AiIQi2Lw8HoEIhCLsHrGIhaL4tdqIvDiFowwegIiEIxjUo8i8PB6xjB6BCIQjKIWjOLXaSLw4haNMHoCIhCNYvDwegQiEI2wesYiFo3i12gi8OIWjjB6AiIQjmLw8HoEIhCOsHrGIhaO4tVnIvCwegIiBGIQQGLwl/B6BDB6hheiEECiFEDW4vlXcIUAMxVi+xW/3UQi3UI/3UMVuhNMAAAahD/dRSNRiBQ6H8iAACLRRiDxAzHRnQAAAAAiUZ4Xl3CFADMzMzMzMzMzMzMVYvsVot1CFf/dQz/djCNfiBXjUYQUFboRPn//4tWeDPAgAcBdQtAO8J0BoAEOAF09V9eXcIIAMzMzMzMzMzMzFWL7IPsEI1F8GoQ/3UgUOgMIgAAg8QMjUXwUGoA/3Uk/3Uc/3UY/3UU/3UQ/3UM/3UI6PkqAACL5V3CIADMzMxVi+z/dSRqAf91IP91HP91GP91FP91EP91DP91COjOKgAAXcIgAMzMzMzMzMzMzMxVi+zoCOf//7ngaw0AgekAQA0AA8GLTQhRUP91FIsB/3UQ/3UM/zCNQShQjUEYUOis9///XcIQAMzMzMzMzMzMVYvsi00Ii0UMiUEsi0UQiUEwXcIMAMzMzMzMzMzMzMxVi+xWi3UIajRqAFbobyEAAItNDMdGLAAAAACLAYlGMItFEIlGBI1GCIkOx0YoAAAAAP8x/3UUUOgTIQAAg8QYXl3CEADMzMzMzMzMzMzMzFWL7IHsIAQAAFNWV2pwj" & _
                                                    "YVw/f//x4Vg/f//QdsAAGoAUMeFZP3//wAAAADHhWj9//8BAAAAx4Vs/f//AAAAAOjsIAAAi3UMjYVg////ah9WUOiqIAAAikYfg8QYgKVg////+CQ/DECIhX////+NheD7////dRBQ6MQ2AAAPV8CNtWD+//9mDxOFYP7//429aP7//7keAAAAZg8TRYDzpbkeAAAAZg8TheD+//+NdYDHhWD+//8BAAAAjX2Ix4Vk/v//AAAAAPOluR4AAADHRYABAAAAjbXg/v//x0WEAAAAAI296P7//7v+AAAA86W5IAAAAI214Pv//4294P3///Oli8MPtsvB+AOD4QcPtrQFYP///42F4P3//9Pug+YBVlCNRYBQ6HYsAABWjYVg/v//UI2F4P7//1DoYiwAAI2F4P7//1CNRYBQjYXg/P//UOhr6///jYXg/v//UI1FgFBQ6Fo0AACNhWD+//9QjYXg/f//UI2F4P7//1DoQOv//42FYP7//1CNheD9//9QUOgsNAAAjYXg/P//UI2FYP7//1Do+TMAAI1FgFCNhWD8//9Q6OkzAACNRYBQjYXg/v//UI1FgFDopSAAAI2F4Pz//1CNheD9//9QjYXg/v//UOiLIAAAjYXg/v//UI1FgFCNheD8//9Q6MTq//+NheD+//9QjUWAUFDoszMAAI1FgFCNheD9//9Q6IMzAACNhWD8//9QjYVg/v//UI2F4P7//1DoiTMAAI2FYP3//1CNheD+//9QjUWAUOgiIAAAjYVg/v//UI1FgFBQ6GHq//+NRYBQjYXg/v//UFDoACAAAI2FYPz//1CNhWD+//9QjUWAUOjpHwAAjYXg+///UI2F4P3//1CNhWD+//9Q6M8fAACNheD8//9QjYXg/f" & _
                                                    "//UOjsMgAAVo2F4P3//1CNRYBQ6NsqAABWjYVg/v//UI2F4P7//1DoxyoAAIPrAQ+JH/7//42F4P7//1BQ6BEcAACNheD+//9QjUWAUFDocB8AAI1FgFD/dQjoxCEAAF9eW4vlXcIMAMzMzMzMzMzMzMzMVYvsg+wgjUXgxkXgCVD/dQwPV8DHRfkAAAAA/3UIDxFF4WbHRf0AAGYP1kXxxkX/AOiq/P//i+VdwggAzMzMzFWL7IHsFAEAAFOLXQiNRfBWV4t9DA9XwFBQi0MEV8ZF8ABmD9ZF8cdF+QAAAABmx0X9AADGRf8A/9CLdSSD/gx1IFb/dSCNRdBQ6FEdAACDxAxmx0XdAADGRdwAxkXfAeswjUXwUI2F7P7//1Do/hoAAFb/dSCNhez+//9Q6B4ZAACNRdBQjYXs/v//UOjOGQAAjUXwUI2FPP///1DozhoAAP91HI2FPP////91GFDozBgAAI1F0MZF4ABQV1ONRYzHRekAAAAAD1fAZsdF7QAAUGYP1kXhxkXvAOhw+///agRqDI1FjFDoQ/v//2oQjUXgUFCNRYxQ6PP6////dRSNhTz/////dRBQ6JEYAACNRcBQjYU8////UOhBGQAAi3UsjUXgVlCNRcBQUOhfQwAAMtKNRcC7AQAAAIX2dBqLfSiLyCv5igwHjUABMkj/CtEr83XxhNJ1FP91FI1FjP91MP91EFDohfr//zPbD1fADxFF8IpF8A8RRdCKRdAPEUXgikXgDxFFwIpFwGpQjYU8////agBQ6DQcAACKjTz///+NRYxqNGoAUOghHAAAik2Mg8QYi8NfXluL5V3CLABVi+yB7BQBAABTi10IjUXwVleLfQwPV8BQUItDBFfGRfAAZg/WRfHHRfk" & _
                                                    "AAAAAZsdF/QAAxkX/AP/Qi3Ukg/4MdSBW/3UgjUXQUOiRGwAAg8QMZsdF3QAAxkXcAMZF3wHrMI1F8FCNhez+//9Q6D4ZAABW/3UgjYXs/v//UOheFwAAjUXQUI2F7P7//1DoDhgAAI1F8FCNhTz///9Q6A4ZAAD/dRyNhTz/////dRhQ6AwXAACNRdDGReAAUFdTjUWMx0XpAAAAAA9XwGbHRe0AAFBmD9ZF4cZF7wDosPn//2oEagyNRYxQ6IP5//9qEI1F4FBQjUWMUOgz+f//i30UjUWMi3UoV1b/dRBQ6B/5//9XVo2FPP///1DowRYAAI1FwMZFwABQjYU8////x0XJAAAAAA9XwGbHRc0AAFBmD9ZFwcZFzwDoVBcAAP91MI1F4FCNRcBQ/3Us6HFBAAAPV8APEUXwikXwDxFF0IpF0A8RReCKReAPEUXAikXAalCNhTz///9qAFDoghoAAIqFPP///2o0jUWMagBQ6G8aAACKRYyDxBhfXluL5V3CLABVi+yLVQyLTRBWi3UIiwYzAokBi0YEM0IEiUEEi0YIM0IIiUEIi0YMM0IMiUEMXl3CDADMzMzMzMzMzMzMzMzMVYvsUVOLXQxWV4t9CGbHRfwA4YsPi8HR6IPhAYkDi1cEi8LR6IPiAcHhHwvIweIfiUsEi3cIi8bR6IPmAQvQweYfiVMIi08Mi8HR6IPhAQvwX4lzDA+2RA38weAYMQNeW4vlXcIIAMzMzMzMzMzMzFWL7ItVDFaLdQgPtg4PtkYBweEIC8gPtkYCweEIC8gPtkYDweEIC8iJCg+2TgQPtkYFweEIC8gPtkYGweEIC8gPtkYHweEIC8iJSgQPtk4ID7ZGCcHhCAvID7ZGCsHhCAvID7ZGC8Hh" & _
                                                    "CAvIiUoID7ZODA+2Rg3B4QgLyA+2Rg7B4QgLyA+2Rg/B4QgLyIlKDF5dwggAzMzMzMzMzMzMzMxVi+yD7CBWV2oQjUXgagBQ6PsYAABqEP91DI1F8FDovRgAAIt9CIPEGA8QTeAz9pCLxrkfAAAAg+AfK8iLxsH4BYsEh9PoqAF0DA8QRfBmD+/IDxFN4I1F8FBQ6JD+//9Ggf6AAAAAfMdqEI1F4FD/dRDoaRgAAIPEDF9ei+VdwgwAzMzMzMzMzMzMzMzMzMxVi+xWi3UMV4t9CIsXi8LB6BiIBovCwegQiEYBi8LB6AiIRgKIVgOLTwSLwcHoGIhGBIvBwegQiEYFi8HB6AiIRgaITgeLTwiLwcHoGIhGCIvBwegQiEYJi8HB6AiIRgqITguLTwyLwcHoGIhGDIvBwegQiEYNi8HB6AiIRg5fiE4PXl3CCADMzMzMzMzMzMxVi+yD7ERWi3UIg76oAAAAAHQGVuhnHgAAM8kPH0QAAA+2hA6IAAAAiUSNvEGD+RBy7lbHRfwAAAAA6EEdAACNRbxQVujXHAAAi1UMM8lmkIoEjogEEUGD+RBy9GisAAAAagBW6IcXAACKBoPEDF6L5V3CCADMzMzMzMzMzMzMzFWL7FaLdQhorAAAAGoAVuhcFwAAi00MahD/dRAPtgGJRkQPtkEBiUZID7ZBAolGTA+2QQOD4A+JRlAPtkEEJfwAAACJRlQPtkEFiUZYD7ZBBolGXA+2QQeD4A+JRmAPtkEIJfwAAACJRmQPtkEJiUZoD7ZBColGbA+2QQuD4A+JRnAPtkEMJfwAAACJRnQPtkENiUZ4D7ZBDolGfA+2QQ+D4A/HhoQAAAAAAAAAiYaAAAAAjYaIAAAAUOiBFgAAg8QYXl3CD"
Private Const STR_THUNK2                As String = "ADMzMzMzMzMzMxVi+zoyNv//7nwgw0AgekAQA0AA8GLTQhRUP91EI2BqAAAAP91DGoQUI2BmAAAAFDoa+v//13CDADMzMzMzMzMVYvsg+wYU1ZX6ILb////dQi+YIkNALlAAAAAge4AQA0AA/CLRQhWjXhki0Bg9+EDB4vYg9IAg8AIg+A/K8hRagBqAGiAAAAAakBXi30ID6TaA4lV/I1HIMHjA1CJVfjoDOr//4tV/IvLi8KIXe/B6BiIReiLwsHoEIhF6YvCwegIiEXqikX4iEXri8IPrMEYagjB6BiITeyLwovLD6zBEMHoEIvDiE3tD6zQCIhF7o1F6FDB6ghX6GQBAACLF4vCi3UMwegYiAaLwsHoEIhGAYvCwegIiEYCiFYDi08Ei8HB6BiIRgSLwcHoEIhGBYvBwegIiEYGiE4Hi08Ii8HB6BiIRgiLwcHoEIhGCYvBwegIiEYKiE4Li08Mi8HB6BiIRgyLwcHoEIhGDYvBwegIiEYOiE4Pi08Qi8HB6BiIRhCLwcHoEIhGEYvBwegIiEYSiE4Ti08Ui8HB6BiIRhSLwcHoEIhGFYvBwegIiEYWiE4Xi08Yi8HB6BiIRhiLwcHoEIhGGYvBwegIiEYaiE4bi08ci8HB6BiIRhyLwcHoEIhGHYvBamjB6AhqAIhGHleITh/oqRQAAIPEDF9eW4vlXcIIAMzMzMzMzMzMzMzMzMxVi+xWi3UIamhqAFbofxQAAIPEDMcGZ+YJasdGBIWuZ7vHRghy8248x0YMOvVPpcdGEH9SDlHHRhSMaAWbx0YYq9mDH8dGHBnN4FteXcIEAFWL7Oho2f//uWCJDQCB6QBADQADwYtNCFFQ/3UQjUFk/3UMakBQjUEgUOgR6f//XcIMAM" & _
                                                    "zMzMzMzMzMzMzMzMxVi+yD7ECNRcBQ/3UI6L4AAABqMI1FwFD/dQzosBMAAIPEDIvlXcIIAMzMzMzMzMxVi+xWi3UIaMgAAABqAFbovBMAAIPEDMcG2J4FwcdGBF2du8vHRggH1Xw2x0YMKimaYsdGEBfdcDDHRhRaAVmRx0YYOVkO98dGHNjsLxXHRiAxC8D/x0YkZyYzZ8dGKBEVWGjHRiyHSrSOx0Ywp4/5ZMdGNA0uDNvHRjikT/q+x0Y8HUi1R15dwgQAzMzMzMzpiwMAAMzMzMzMzMzMzMzMVYvsg+wci0UIU42YxAAAAFaLgMAAAABXv4AAAAD354vwAzOLxoPSAA+kwgPB4AOJVfyJRfiJVfToI9j///91CLkwiw0AgekAQA0AA8FQjUYQi3UIg+B/K/hXagBqAGiAAAAAaIAAAABTjUZAUOjO5v//agiNReTHReQAAAAAUFbHRegAAAAA6PQCAACLXfyLw4tV+IvKwegYiEXki8PB6BCIReWLw8HoCIhF5opF9IhF54vDD6zBGGoIwegYiE3oi8OLyohV6w+swRDB6BCLwohN6Q+s2AiIReqNReRQVsHrCOiZAgAAi14Ei8OLDolN/MHoGIt9DIgHi8PB6BCIRwGLw8HoCIhHAovDD6zBGIhfA8HoGIhPBIvDi038D6zBEMHoEIhPBYtN/IvBD6zYCIhHBovGiE8HwesIi1gIi8uLUAyLwsHoGIhHCIvCwegQiEcJi8LB6AiIRwqLwg+swRiIVwvB6BiITwyLwovLD6zBEMHoEIhPDYvDD6zQCIhHDovGiF8PweoIi1gQi8uLUBSLwsHoGIhHEIvCwegQiEcRi8LB6AiIRxKLwg+swRiIVxPB6BiITxSLwovLD6zBEMH" & _
                                                    "oEIvDiE8VD6zQCIhHFovGweoIiF8Xi1gYi8uLUByLwsHoGIhHGIvCwegQiEcZi8LB6AiIRxqLwg+swRiIVxvB6BiITxyLwovLD6zBEMHoEIhPHYvDD6zQCIhHHovGiF8fweoIi1ggi8uLUCSLwsHoGIhHIIvCwegQiEchi8LB6AiIRyKLwg+swRiIVyPB6BiITySLwovLD6zBEMHoEIhPJYvDD6zQCIhHJovGiF8nweoIi1goi8uLUCyLwsHoGIhHKIvCwegQiEcpi8LB6AiIRyqLwg+swRiIVyvB6BiITyyLwovLD6zBEMHoEIvDiE8tD6zQCMHqCIhHLovGiF8vjXc4aMgAAABqAItYMIvLi1A0i8LB6BiIRzCLwsHoEIhHMYvCwegIiEcyi8IPrMEYiFczwegYiE80i8KLyw+swRDB6BCITzWLww+s0AiIRzaIXzeLfQjB6ghXi1c8i8KLXziLy8HoGIgGi8LB6BCIRgGLwsHoCIhGAovCD6zBGIhWA8HoGIhOBIvCi8sPrMEQwegQi8OITgUPrNAIiEYGweoIiF4H6MUPAACDxAxfXluL5V3CCADMzMzMzMzMzMxVi+zo2NT//7kwiw0AgekAQA0AA8GLTQhRUP91EI2BxAAAAP91DGiAAAAAUI1BQFDoe+T//13CDADMzMzMzMzMVYvsVot1CP91DIsOjUYIUP92BItBBP/Qi1Ysi0YwA9ZIXoBEAggBdRMPH4AAAAAAhcB0CEiARAIIAXT0XcIIAFWL7FOLXQxWV4t9CA+2QxiZi8iL8g+kzggPtkMZweEImQvIC/IPpM4ID7ZDGsHhCJkLyAvyD6TOCA+2QxvB4QiZC8gL8g+2QxwPpM4ImcHhCAvyC8gPtkMdD6TOCJnB" & _
                                                    "4QgL8gvID7ZDHg+kzgiZweEIC/ILyA+2Qx8PpM4ImcHhCAvyC8iJdwSJDw+2QxCZi8iL8g+2QxEPpM4ImcHhCAvyC8gPtkMSD6TOCJnB4QgL8gvID7ZDEw+kzgiZweEIC/ILyA+2QxQPpM4ImcHhCAvIC/IPpM4ID7ZDFcHhCJkLyAvyD6TOCA+2QxbB4QiZC8gL8g+kzggPtkMXweEImQvIC/KJTwiJdwwPtkMImYvIi/IPpM4ID7ZDCcHhCJkLyAvyD7ZDCg+kzgiZweEIC/ILyA+2QwsPpM4ImcHhCAvyC8gPtkMMD6TOCJnB4QgL8gvID7ZDDQ+kzgiZweEIC/ILyA+2Qw4PpM4ImcHhCAvyC8gPtkMPD6TOCJnB4QgL8gvIiXcUiU8QD7YDmYvIi/IPtkMBD6TOCJnB4QgL8gvID7ZDAg+kzgjB4QiZC8gL8g+2QwMPpM4ImcHhCAvyC8gPtkMED6TOCJnB4QgL8gvID7ZDBQ+kzgiZweEIC/ILyA+2QwYPpM4ImcHhCAvyC8gPtkMHD6TOCJnB4QgLyAvyiXcciU8YX15bXcIIAMzMzFWL7IPsYI1F4P91DFDo3v3//41F4FDo5SUAAIXAdAgzwIvlXcIIAI1F4FDoANL//4PogFDoVyUAAIP4AXQT6O3R//+D6IBQjUXgUFDo7zEAAGoAjUXgUOjU0f//g8BAUI1FoFDoh9P//41FoFDoTtP//4XAdamKRcCLTQgkAQQCiAGNRaBQjUEBUOgRAAAAuAEAAACL5V3CCADMzMzMzMxVi+xWi3UIsShXi30MD7ZHB4hGGA+2RwaIRhmLB4tXBOjL1///iEYasSCLB4tXBOi81///iEYbiw+LRwQPrMEYiE4ciw/B6BiLRwQPr" & _
                                                    "MEQiE4diw/B6BCLRwQPrMEIiE4esSjB6AgPtgeIRh8PtkcPiEYQD7ZHDohGEYtHCItXDOhr1///iEYSsSCLRwiLVwzoW9f//4hGE4tPCItHDA+swRiIThSLTwjB6BiLRwwPrMEQiE4Vi08IwegQi0cMD6zBCIhOFrEowegID7ZHCIhGFw+2RxeIRggPtkcWiEYJi0cQi1cU6AbX//+IRgqxIItHEItXFOj21v//iEYLi08Qi0cUD6zBGIhODItPEMHoGItHFA+swRCITg2LTxDB6BCLRxQPrMEIiE4OsSjB6AgPtkcQiEYPD7ZHH4gGD7ZHHohGAYtHGItXHOii1v//iEYCsSCLRxiLVxzoktb//4hGA4tPGItHHA+swRjB6BiITgSLTxiLRxwPrMEQwegQiE4Fi08Yi0ccD6zBCMHoCIhOBg+2RxhfiEYHXl3CCADMzFWL7IPsIFNWi3UID1fAV4t9DMdF4AMAAADHReQAAAAADxFF6I1HAWYP1kX4UFboffv//1aNXiBT6EMrAADors///1CNReBQU1PoYisAAFZTU+j6KgAA6JXP//9Q6I/P//+DwCBQU1PoBCcAAFPoDgsAAIoHM/aLCyQBD7bAg+EBmTvIdQQ78nQNU+hhz///UFPoai8AAF9eW4vlXcIIAMxVi+yB7KAAAACNhWD/////dQhQ6Ej/////dQyNReBQ6Oz6//9qAI1F4FCNhWD///9QjUWgUOjW0P//jUWgUP91EOh6/f//jUWgUOiR0P//99gbwECL5V3CDADMzMzMzMxVi+yD7ECNRcBW/3UIUOjt/v//i3UMjUXAUI1GAcYGBFDoOv3//41F4FCNRiFQ6C39//+4AQAAAF6L5V3CCADMVYvsgeyAAAAAV4" & _
                                                    "t9EFfobSIAAIXAdAkzwF+L5V3CEABX6IrO//+D6IBQ6OEhAACD+AF0EOh3zv//g+iAUFdX6HwuAABqAFfoZM7//4PAQFCNRYBQ6BfQ//+NRYBQ6E7O//+D6IBQ6KUhAACD+AF0E+g7zv//g+iAUI1FgFBQ6D0uAACNRYBQ6PQhAACFwHWHVot1FI1FgFBW6IL8////dQiNRcBQ6Mb5///oAc7//4PogFCNRcBQjUWAUI1F4FDo7CcAAP91DI1FwFDooPn//+jbzf//g+iAUI1F4FCNRcBQjUXgUOhGJQAA6MHN//+D6IBQV1fodiUAAOixzf//g+iAUFeNReBQUOiiJwAAjUXgUI1GIFDoBfz//164AQAAAF+L5V3CEADMzMzMzMzMzFWL7IHssAEAAI2FUP7//1b/dQhQ6Hf9//+LdRCNhTD///9WUOgX+f//jUYgUI1FkFDoCvn//42FMP///1DoDiEAAIXAD4VWAwAAjUWQUOj9IAAAhcAPhUUDAACNhTD///9Q6BnN//+D6IBQ6HAgAACD+AEPhScDAACNRZBQ6P7M//+D6IBQ6FUgAACD+AEPhQwDAABTV+jlzP//g+iAUI1FkFCNReBQ6JQkAAD/dQyNhVD///9Q6IX4///owMz//4PogFCNReBQjYVQ////UFDoqyYAAOimzP//g+iAUI1F4FCNhTD///9QjYUQ////UOiLJgAAjYVQ/v//UI2FsP7//1DoqCoAAI2FcP7//1CNhdD+//9Q6JUqAADoYMz//4PAQFCNhXD///9Q6IAqAADoS8z//4PAYFCNRZBQ6G4qAADoOcz//1CNhXD///9QjYWw/v//UI1F4FDo4ScAAI2F0P7//1CNhbD+//9QjUWQUI2FcP///1D" & _
                                                    "oY8///+j+y///UI1F4FBQ6LMjAACNReBQjYXQ/v//UI2FsP7//1DorNT//8dF0AAAAADo0Mv//4PAQIlF1I2FUP7//4lF2I2FsP7//4lF3I2FEP///1Do/CgAAFCNhVD///9Q6O8oAABQ6HkcAACL2I2FUP///417/1dQ6PcsAAALwnQHvgEAAADrAjP2V42FEP///1Do3SwAAAvCdAe4AgAAAOsCM8AL8I1FsIt0tdBWUOiAKQAAjUYgUI2F8P7//1DocCkAAI1F4FDoRx4AAI1z/sdF4AEAAADHReQAAAAAhfYPiNIAAACNReBQjYXw/v//UI1FsFDoSsv//1aNhVD///9Q6G0sAAALwnQHvwEAAADrAjP/Vo2FEP///1DoUywAAAvCdAe4AgAAAOsCM8ALx4t8hdCF/3R3V42FcP///1Do7ygAAI1HIFCNRZBQ6OIoAACNReBQjUWQUI2FcP///1DobtP//+iZyv//UI2FcP///1CNRbBQjYWQ/v//UOhBJgAAjYXw/v//UI1FsFCNRZBQjYVw////UOjGzf//jYWQ/v//UI1F4FBQ6LUlAACD7gEPiS7////oR8r//1CNReBQUOj8IQAAjUXgUI2F8P7//1CNRbBQ6PjS//+NRbBQ6B/K//+D6IBQ6HYdAABfW4P4AXQT6ArK//+D6IBQjUWwUFDoDCoAAI2FMP///1CNRbBQ6EwdAAD32F4bwECL5V3CDAAzwF6L5V3CDADMzMzMzMzMVYvsi00Ii8HB6AeB4X9/f/8lAQEBAQPJa8AbM8FdwgQAzMzMzMzMzMzMzMzMzMzMVYvs6LjJ//+5oHcNAIHpAEANAAPBi00IUVD/dRCNQTD/dQxqEFCNQSBQ6GHZ//9dwgwAzMzM" & _
                                                    "zMzMzMzMzMzMzFWL7ItNCItFEAFBOINRPACJRRCJTQhd6aT////MzMzMVYvsVot1CIN+SAF1DVboLQAAAMdGSAIAAACLRRABRkBQ/3UMg1ZEAFbocv///15dwgwAzMzMzMzMzMzMzMzMzFWL7FaLdQiLTjCFyXQpuBAAAAArwVCNRiADwWoAUOjNAwAAg8QMjUYgUFboEAAAAMdGMAAAAABeXcIEAMzMzMxVi+yD7BCNRfBWV1D/dQzo7On//4t1CI1F8I1+EFdXUOgr6f//V1ZX6HPq//9fXovlXcIIAMzMzMzMzMzMzMzMVYvsg+wUU1aLdQiLRkiD+AF0BYP4AnUNVuhi////x0ZIAAAAAIteOItWPA+k2gNqCIvCweMDwegYi8uIReyLwsHoEIhF7YvCwegIiEXuD7bCiEXvi8IPrMEYiVX8wegYiE3wi8KLy4hd8w+swRDB6BCLw4hN8Q+s0AiIRfKNRexQweoIVuhW/v//i15Ai1ZED6TaA2oIi8LB4wPB6BiLy4hF7IvCwegQiEXti8LB6AiIRe4PtsKIRe+Lwg+swRiJVfzB6BiITfCLwovLiF3zD6zBEMHoEIvDiE3xD6zQCIhF8o1F7FDB6ghW6PH9////dQyNRhBQ6PXp//9eW4vlXcIIAMzMzMzMzMzMzMzMzMxVi+xWi3UIalBqAFboTwIAAIPEDFb/dQzok+j//8dGSAEAAABeXcIIAMzMzMzMzMxVi+yB7IAAAAC5IAAAAFOLXQxWV4vzjX2A86W+/QAAAI1FgFBQ6HYWAACD/gJ0EIP+BHQLU41FgFBQ6DEDAACD7gF53It9CI11gLkgAAAA86VfXluL5V3CCADMzMzMzMxVi+xTVot1CFdW6AH9//+L2FPo+" & _
                                                    "fz//4vQUujx/P//i/gz/ov3i8czw8HPCDPywcAIi87ByRAzwTPHM8ZfM8MzRQheW13CBADMzMzMzMzMzFWL7FaLdQj/Nuii/////3YEiQbomP////92CIlGBOiN/////3YMiUYI6IL///+JRgxeXcIEAMzMzMzMzMzMzMxVi+xTi10IVlcPtnsHD7ZDAg+2cwsPtlMPwecIC/gPtksDD7ZDDcHnCAv4weYID7ZDCMHnCAv4weIID7ZDBgvwweEID7ZDAcHmCAvwD7ZDDMHmCAvwD7ZDCgvQD7ZDBcHiCAvQD7YDweIIC9APtkMOC8iJUwwPtkMJweEIC8iJcwgPtkMEiXsEweEIXwvIXokLW13CBADMzMzMzMzMzMzMVYvsVuinxf//i3UIBZMFAABQ/zboJxYAAIkG6JDF//8FkwUAAFD/dgToEhYAAIlGBOh6xf//BZMFAABQ/3YI6PwVAACJRgjoZMX//wWTBQAAUP92DOjmFQAAiUYMXl3CBADMzMzMzMzMzMzMzMzMzFWL7ItFCIvQVot1EIX2dBVXi30MK/iKDBeNUgGISv+D7gF18l9eXcPMzMzMzMzMzFWL7ItNEIXJdB8PtkUMVovxacABAQEBV4t9CMHpAvOri86D4QPzql9ei0UIXcPMzFWL7FaLdQhW6AP7//+L0IvOM9bByRDBwgjBzggz0TPWM8JeXcIEAMzMzMzMzMzMzFWL7FaLdQj/NujC/////3YEiQbouP////92CIlGBOit/////3YMiUYI6KL///+JRgxeXcIEAMzMzMzMzMzMzMxVi+yD7EBWD1fAx0XAAQAAAFeNRcDHRcQAAAAAUA8RRcjHReABAAAAZg/WRdjHReQAAAAADxFF6GYP1kX46C7E//" & _
                                                    "9QjUXAUOj0FQAAjUXAUOhrIQAAi30IjXD/g/4BdimNReBQUOiWHwAAVo1FwFDobCUAAAvCdAtXjUXgUFDoTR8AAE6D/gF3141F4FBX6A0iAABfXovlXcIEAMzMzMzMVYvsgewAAQAAi0UMD1fAU1ZXuTwAAABmDxOFAP///421AP///8dF/BAAAACNvQj////zpYtNEI2dCP///4PBEIvTK8KJTfiJRQxmDx9EAACL+cdFEAQAAACL8w8fRAAA/3QYBP80GP939P938Ohuyf//AUb4i0UMEVb8/3QYBP80GP93/P93+OhTyf//AQaLRQwRVgT/dBgE/zQY/3cE/zfoOsn//wFGCItFDBFWDP90GAT/NBj/dwz/dwjoH8n//wFGEI1/IItFDBFWFI12IINtEAF1iotN+IPDCINt/AEPhWr///8z9moAaib/dPWE/3T1gOjnyP//AYT1AP///2oAEZT1BP///2om/3T1jP909YjoyMj//wGE9Qj///9qABGU9Qz///9qJv909ZT/dPWQ6KnI//8BhPUQ////agARlPUU////aib/dPWc/3T1mOiKyP//AYT1GP///2oAEZT1HP///2om/3T1pP909aDoa8j//wGE9SD///8RlPUk////g8YFg/4PD4JZ////i10IjbUA////uSAAAACL+/OlU+hJy///U+hDy///X15bi+VdwgwAzMzMzMzMzMzMzFWL7IPsEFNWi3UMV4t9GGoAVmoA/3UU6ATI//9qAFZqAFeJRfCL2uj0x///agD/dRCJRfSL8moAV+jix///agD/dRCJRfxqAP91FIlV+OjNx///i/iLRfQD+4PSAAP4E9Y71ncOcgQ7+HMIg0X8AINV+AGLRQgzyQtN8IkIM8k" & _
                                                    "DVfyJeAQTTfhfXolQCIlIDFuL5V3CFADMzMzMzMzMzMxVi+yB7AgBAACNhXj///9TVlf/dQxQ6KUJAACNhXj///9Q6GnK//+NhXj///9Q6F3K//+NhXj///9Q6FHK//+Nvfj+//+7AgAAAGYPH0QAAIuNeP///4uFfP///4Hp7f8AAImN+P7//4PYAImF/P7//7gIAAAAZmYPH4QAAAAAAIt0B/iLTAf8i5QFeP///4l1+A+szhCLjAV8////g+YBx0QH/AAAAAAr1oPZAIHq//8AAImUBfj+//+D2QCJjAX8/v//D7dN+IlMB/iDwAiD+HhyrIuNaP///4uFbP///4tV8A+swRAPt4Vo////g+EBiYVo////K9HHhWz///8AAAAAi030uAEAAACD2QCB6v9/AACJlXD///+D2QCJjXT///8PrMoQg+IBwfkQK8JQjYX4/v//UI2FeP///1DojQcAAIPrAQ+FBP///4t1CDPSioTVeP///4uM1Xj///+IBFaLhNV8////D6zBCIhMVgFCwfgIg/oQctdfXluL5V3CCADMzMzMzMzMzMzMzMzMVYvsi0UIM9JWV4t9DCv4jXIRiwwHjUAEA0j8A9EPtsrB6giJSPyD7gF1519eXcIIAMzMzMzMzMzMzMzMzMzMzFWL7Fb/dQyLdQhW6LD///+NRkRQVui2AQAAXl3CCADMVYvsg+xEU1aLdQhXDxAGi0ZAiUX8DxFFvA8QRhAPEUXMDxBGIA8RRdwPEEYwDxFF7OhKv///BUQEAABQjUW8UOhb////i0X8jX2899CNVcwlgAAAACv+uQIAAACNWP/30MHoH8HrHyPY99sr1ovD99CJRQhmD27DZg9w0ABmD27Ai8ZmD3DYAA8fhAAA" & _
                                                    "AAAAjUAgDxBA4A8QTAfgZg/bwmYP28tmD+vIDxFI4A8QQPAPEEwC4GYP28JmD9vLZg/ryA8RSPCD6QF1xo1WQI1xAYsMOo1SBCNNCIvDI0L8C8iJSvyD7gF16F9eW4vlXcIEAMzMzMzMzMzMzMzMzMzMzFWL7IPsRI1FvFZqRGoAUOhc+f//i3UIg8QMM8CLlqgAAACF0nQbZmYPH4QAAAAAAA+2jAaYAAAAiUyFvEA7wnLvjUW8x0SVvAEAAABQVuiN/v//XovlXcIEAMzMzMzMzFWL7FaLdQgzwDPSDx9EAAADBJYPtsiJDJZCwegIg/oQfO4DRkCLyMHoAoPhAzPSiU5AjQyAAwyWD7bBiQSWQsHpCIP6EHzuAU5AXl3CBADMVYvsg+xUi0UMjU2sU1aLdQgz2yvBx0X4EAAAAFeJRfAz0jP/M8CJVQiJVfyF23hRjUsBg/kCfDCLTfCNVayNDJkD0YsMho1S+A+vSggBTQiLTIYEg8ACD69KBAFN/I1L/zvBft6LVQg7w38Oi30Mi8sryIs8jw+vPIaLRfwDwgP4jUMBM9KJVQiLyIlV/IlF9IP4EX1yg334AnxDi1UMi8MrwY0UgoPCQA8fgAAAAACLBI6NUvgPr0IMjQSAweAGAUUIi0SOBIPBAg+vQgiNBIDB4AYBRfyD+RB81ItVCIP5EX0ai1UMi8MrwYtEgkQPrwSOi1UIjQSAweAGA/iLRfwDwgP4i0X0i034SYl8nayJTfiL2IP5/w+PAv///41FrFDoif7//w8QRayLRexfDxEGDxBFvA8RRhAPEEXMDxFGIA8QRdwPEUYwiUZAXluL5V3CCADMzMzMzMzMzMzMzFWL7ItVDIPsRDPADx9EAAAPtgwQiUyFvECD+" & _
                                                    "BB88o1FvMdF/AEAAABQ/3UI6J/8//+L5V3CCADMzMzMzMzMzMxVi+yB7HwBAABTVldqDP91DI1F4MZF3AAPV8DHReUAAAAAUGYP1kXdZsdF6QAAxkXrAOi59v//g8QMxkW8AI1F3MdF1QAAAAAPV8Bmx0XZAAAPEUW9agRQaiD/dQiNhTD///9mD9ZFzVDGRdsA6N7T//9qII1FvFBQjYUw////UOgrzf//jUXMUI1FvFCNhYT+//9Q6Bff//8PV8APEUW8ikW8aiCNRbxQUI2FMP///1APEUXM6PbM//+LdRQPV8BW/3UQDxFFvIpFvI2FhP7//8ZF7ABQDxFFzMdF9QAAAABmD9ZF7WbHRfkAAMZF+wDoi9///4vG99iD4A9QjUXsUI2FhP7//1Doc9///4N9JAGLfSCLXRxTdRRX/3UYjYUw////UOiGzP//U1frA/91GI2FhP7//1DoQ9///4vD99iD4A9QjUXsUI2FhP7//1DoK9///zPSiF30i8aJVeiIReyLyIvCD6zBCGoQwegIiE3ti8KLzg+swRDB6BCITe6LwovOD6zBGMHoGA+2wohF8IvCwegIiEXxi8LB6BCIRfLB6hiITe+Ly4hV8zPSi8KJVegPrMEIwegIiE31i8KLyw+swRDB6BCITfaLwovLD6zBGMHoGA+2wohF+IvCwegIiEX5i8LB6BCIRfqNRexQjYWE/v//weoYUIhN94hV++h73v//g30kAXUz/3UojYWE/v//UOgW3f//anyNhTD///9qAFDo9vT//4qFMP///4PEDDPAX15bi+VdwiQAjUWsUI2FhP7//1Do4tz//4t1KI1NrIvBMtu6EAAAACvwkIoEDo1JATJB/wrYg+oBdfCLRRyE23U/UF" & _
                                                    "f/dRiNhTD///9Q6CjL//9qfI2FMP///2oAUOiI9P//ioUw////g8QMD1fADxFFrIpFrF9eM8Bbi+VdwiQAhcB0DlBqAFfoXfT//4oHg8QManyNhTD///9qAFDoSPT//4qFMP///4PEDA9XwA8RRayKRaxfXrgBAAAAW4vlXcIkAMzMzMzMzMxVi+xWV4t9CA+2B5mLyIvyD7ZHAQ+kzgiZweEIC/ILyA+2RwIPpM4ImcHhCAvyC8gPtkcDD6TOCJnB4QgL8gvID7ZHBA+kzgiZweEIC/ILyA+2RwUPpM4ImcHhCAvyC8gPtkcGD6TOCJnB4QgL8gvID7ZHBw+kzgiZweEIC8EL1l9eXcIEAMzMzMzMzMzMzMxVi+yD7AiLRRBI99CZU4tdCIlF+ItFDIlV/PMPfl34jUt4VjP2Zg9s241QeDvBd0s703JHK9jHRRAQAAAAV2aQizwYjUAIi3QY/ItI+ItQ/DPPI034M9YjVfwz+TPyiXwY+Il0GPwxSPgxUPyDbRABdc5fXluL5V3CDACL041IECvQDxAM841JIA8QUdBmD+/RZg/b0w8owmYP78EPEQTzg8YEDxBB0GYP79APEVHQDxBMCuAPEFHgZg/v0WYP29MPKMJmD+/BDxFECuAPEEHgZg/vwg8RQeCD/hBypV5bi+VdwgwAzMzMzMzMzMzMzMxVi+yLVQyLRQgr0Fa+EAAAAIsMAo1ACIlI+ItMAvyJSPyD7gF1615dwggAzMzMzMxVi+yLRRBWV4P4EHQ/g/ggdAb/FdTADQCLdQyLfQhqEFZX6Bny//9qEI1GEFCNRxBQ6Ary//+DxBjoQrf//wUxBAAAiUcwX15dwgwAi3UMi30IahBWV+jl8f//ahCNRxBWUOjZ8f/"
Private Const STR_THUNK3                As String = "/g8QY6BG3//8FIAQAAIlHMF9eXcIMAMzMzFWL7IPsbItFCI1VlFNWu6AAAAAz9otIBIlN+ItICIlN9ItIDIlN6ItIEIlN/ItIFIlN8ItIGIlN7ItNDIPBAol13FeLOCvTi0AciX3giUXkiU3YiV0MiVXUDx+AAAAAAIP+EHMpD7Zx/g+2Qf/B5ggL8A+2AcHmCAvwD7ZBAcHmCAvwg8EEiTQaiU3Y61SNXgGD5g+NQ/2D4A+NfZSNPLeLTIWUi8OD4A+L8cHGD4tUhZSLwcHADTPwwekKM/GLwovKwcgHwcEOM8jB6gONQ/gzyotdDIPgDwPxA3SFlAM3iTfoGbb//4t9/IvXwcoLi8/BwQcz0YvPwckG99cjfewz0YsMGIPDBItF8APKI0X8A86LdeAz+IvWiV0MwcoNi8bBwAoD+QN95DPQi8bByAIz0ItF+IvII8YzziNN9DPIi0XsiUXkA9GLRfCLTfiJReyLRfyJRfCLRegDx4l1+It13AP6i1XURolF/ItF9IlN9ItN2IlF6Il94Il13IH7oAEAAA+C1/7//4tFCItN+ItV/AFIBItN9AFICAFQEAE4i03oi1XwAUgMAVAUi1Xsi03kAVAYAUgc/0BgX15bi+VdwggAzMzMzMzMzMzMzMzMVYvsgezgAAAAU1aLdQi7oAEAAFeJXbiLBolF7ItGBIlF8ItGDIt+CIlF4ItGEIlF1ItGFIlF0ItGGIlFtItGHIlFsItGIIlF6ItGJIlF9ItGKIlFzItGLIlFyItGMIlFxItGNIlFwItGOIlFrItGPIt1DIl92I29IP///4lFqDPAK/uJRdyJfaAPH4AAAAAAg/gQcx9W6GX7//+LyIPGCIvCiU0MiUXkiQwfiUQfBOkTAQAA" & _
                                                    "jVABx0UMAAAAAI1C/YPgD4uMxSD///+LhMUk////iUX4i8KD4A+JTfyNjSD///+LlMUg////i/qLnMUk////i0Xcg+APiVW8wecYjQTBi8uJRaSLwg+syAgJRQyLRbzB6QgL+YvLD6zIAYl95Iv60ekz0gvQwecfMVUMC/mLRbyLTeQPrNgHM88xRQyLRfzB6wczyzPbiU3ki034i9EPpMEDweodweADC9mLTfgL0ItF/Iv4D6zIE4lVvDPSC9DB6ROLRbwzwsHnDYtV/Av5i034M98PrMoGM8LB6QaLVQwz2YtN5APQi0XcE8uDwPmD4A8DlMUg////E4zFJP///4tFpAMQiVUME0gEiRCJTeSJSAToZLP//4tV9DP/i03oi9oPpMoXwesJC/rB4ReLVfQL2YtN6Ild/IvZD6zREol9+DP/C/nB6hIxffwz/4tN6MHjDgvai1X0MV34i9kPrNEOweMSC/nB6g4xffwL2otN+ItVuDPLi138i33o99cDHBATTBAEI33Ei1X0i0XI99IjRfQjVcAz0IlN+ItNzCNN6ItF+DP5i03wA98TwgNdDBNF5ANdrIld/BNFqDPbiUX4i0Xsi9APrMgcweIEwekcC9iLRewL0YtN8Iv5D6TBHolVDDPSwe8CC9HB4B4L+DPfMVUMM9KLTfCL+YtF7A+kwRnB7wcL0cHgGTFVDAv4i03YM9+LVeCL+TN97CN91CNN7DNV8DP5I1XQi0XgI0Xwi03EM9CLRQwD34t9+BPCiU2si03Ai1X8A1W0iU2oE32wi03MA138iU3Ei03IiU3Ai03oiU3Mi030iX30i33UiX20i33QiX2wi33YiX3Ui33giX3Qi33siU3Ii8gTTfiLRdyJXexAi124iX3Yg" & _
                                                    "8MIi33wiX3gi32giVXoiU3wiUXciV24gfsgBAAAD4Ib/f//i3UIi0Xsi33YAQaLReARTgSLygF+CIt9tBFGDItF1AFGEItF0BFGFAF+GItFsBFGHAFOIItF9BFGJItFzAFGKItFyBFGLItFxAFGMItFwBFGNItNrAFOOItNqBFOPP+GwAAAAF9eW4vlXcIIAMzMzMzMzMzMzMzMzMzMVYvsU4tdCFZXD7Z7Bw+2QwoPtnMLD7ZTD8HnCAv4D7ZLAw+2Qw3B5wgL+MHmCA+2A8HnCAv4weIID7ZDDgvwweEID7ZDAcHmCAvwD7ZDBMHmCAvwD7ZDAgvQD7ZDBcHiCAvQD7ZDCMHiCAvQD7ZDBgvIiXsED7ZDCcHhCAvIiXMID7ZDDMHhCF8LyIlTDF6JC1tdwgQAzMzMzMzMzMzMzFWL7ItFDFBQ/3UI6MDs//9dwggAzMzMzMzMzMzMzMzMVYvsi0UQU1aLdQiNSHhXi30MjVZ4O/F3BDvQcwuNT3g78XcwO9dyLCv4uxAAAAAr8IsUOCsQi0w4BBtIBI1ACIlUMPiJTDD8g+sBdeRfXltdwgwAi9eNSBCL3ivQK9gr/rgEAAAAjXYgjUkgDxBB0A8QTDfgZg/7yA8RTuAPEEwK4A8QQeBmD/vIDxFMC+CD6AF10l9eW13CDADMzMzMzFWL7Fbo16///4t1CAWIBAAAUP826FcAAACJBujAr///BYgEAABQ/3YE6EIAAACJRgToqq///wWIBAAAUP92COgsAAAAiUYI6JSv//8FiAQAAFD/dgzoFgAAAIlGDF5dwgQAzMzMzMzMzMzMzMzMzMxVi+yLVQxTi10Ii8PB6BiLy1bB6QgPtskPtjQQi8PB6BAPtsAPtgwRweYID7YEEA" & _
                                                    "vGweAIC8EPtsvB4AheWw+2DBELwV3CCADMzMzMzMzMzFWL7ItFDDlFCA9HRQhdwggAzMzMzMzMzMzMzMzMzMzMVYvsi00MU4tdCFaDwxDHRQwEAAAAV4PBAw8fgAAAAAAPtkH+jVsgmY1JCIvwi/oPtkH1D6T3CJnB5ggD8Ilz0BP6iXvUD7ZB95mL8Iv6D7ZB+JkPpMIIweAIA/CJc9gT+ol73A+2QfqZi/CL+g+2QfkPpPcImcHmCAPwiXPgE/qJe+QPtkH8mYvwi/oPtkH7D6T3CJnB5ggD8Ilz6BP6g20MAYl77A+FdP///4tNCF9eW4FheP9/AADHQXwAAAAAXcIIAMzMzMzMzMzMzMzMzFWL7IPsCFOLXQwPV8BWV4t9EIsTi/KLQwSLyGYPE0X4AzcTTwQ78nUGO8h1BOsYO8h3D3IEO/JzCbgBAAAAM9LrC2YPE0X4i0X4i1X8i30IiU8Ei00QiTeLcQgDcwiLSQwTSwwD8BPKO3MIdQU7Swx0IDtLDHcQcgU7cwhzCbgBAAAAM9LrC2YPE0X4i1X8i0X4iU8Mi00QiXcIi3EQA3MQi0kUE0sUA/ATyjtzEHUFO0sUdCA7SxR3EHIFO3MQcwm4AQAAADPS6wtmDxNF+ItV/ItF+IlPFIl3EItLGItbHIlNDItNEItxGAN1DItJHBPLA/ATyjt1DHUEO8t0LDvLdx1yBTt1DHMWiXcYuAEAAACJTxwz0l9eW4vlXcIMAGYPE0X4i1X8i0X4iXcYiU8cX15bi+VdwgwAzMzMzMzMVYvsi0UIxwAAAAAAx0AEAAAAAMdACAAAAADHQAwAAAAAx0AQAAAAAMdAFAAAAADHQBgAAAAAx0AcAAAAAF3CBADMzMzMzMzMzMzMzMz" & _
                                                    "MzMxVi+yLTQy6AwAAAFOLXQhWK9mNQRhXiV0IDx+AAAAAAIs0A4tcAwSLeASLCDvfdy5yIjvxdyg733IadwQ78XIUi10Ig+gIg+oBedVfXjPAW13CCABfXoPI/1tdwggAX164AQAAAFtdwggAzMzMzMzMVYvsi1UIM8APH4QAAAAAAIsMwgtMwgR1D0CD+ARy8bgBAAAAXcIEADPAXcIEAMzMVYvsg+wQU4tdELlAAAAAVot1CCvLV4t9DGYPbsOJTRCLB4tXBIlF+IlV/PMPfk34Zg/zyGYP1g7oI7L//4tNEIlF8ItHCIlV9ItXDIlF+IlV/PMPfk34Zg9uw2YP88jzD35F8GYP68hmD9ZOCOjusf//i00QiUXwi0cQiVX0i1cUiUX4iVX88w9+TfhmD27DZg/zyPMPfkXwZg/ryGYP1k4Q6Lmx//+LTRCJRfCLRxiJVfSLVxyJRfiJVfzzD35N+GYPbsNmD/PI8w9+RfBmD+vIZg/WThjohLH//19eW4vlXcIMAMzMzMzMzMzMzMzMVYvsg+woU4tdCFZXi30MV1PoKgkAAItHLA9XwIlF5ItHMIlF6ItHNIlF7ItHOIlF8ItHPIlF9I1F2GoBUFBmDxNF2MdF4AAAAADowf7//4vwjUXYUFNT6IT8//+LTzgD8ItHMItXPIlF5DPAC0c0iUXojUXYagFQUMdF4AAAAACJTeyJVfDHRfQAAAAA6H7+//8D8I1F2FBTU+hB/P//A/DHReQAAAAAi0cgD1fAiUXYi0ckiUXci0coiUXgi0c4iUXwi0c8iUX0jUXYUFNTZg8TRejoB/z//4tPJAPwM8CJTdgLRyiJRdyLRzCLVzSLyolF+DPAC0csiUXgi0c4iUXoi0c8iUXsM8AL" & _
                                                    "RyCJRfSNRdhQU1OJTeSJVfDov/v//4tPLAPwi1c0M8ALRzAPV8CJRdyLRyCJRfCNRdiJTdgzyQtPKFBTU4lV4MdF5AAAAABmDxNF6IlN9OjBCQAAi1ckK/CLRzAPV8CJRdixIItHNIlF3ItHOIlF4ItHPIlF5ItHIGYPE0Xo6OKv//8LVyyJRfCNRdhQU1OJVfTofgkAAItVDCvwi080M8ALRziLXySJRdwzwAtHPIlN2ItPIDP/iUXgi0Ioi1IsiU3ksSDoe6///wvYx0XwAAAAAIld6Av6i10MiX3si30Ii0MwiUX0jUXYUFdX6CMJAAAr8MdF4AAAAACLQziJRdiLQzyJRdyLQySJReSLQyiJReiLQyyJReyLQzSJRfSNRdhQV1fHRfAAAAAA6OQIAAAr8Hkg6Muo//9QV1fok/r//wPweO9fXluL5V3CCABmDx9EAACF9nURV+imqP//UOgA/P//g/gBdNzolqj//1BXV+ieCAAAK/Dr2szMzMzMzMzMzMxVi+xW/3UQi3UI/3UMVug9+v//C8J1Df91FFbowPv//4XAeAr/dRRWVuhiCAAAXl3CEADMzMzMzMzMzMzMzMzMVYvsgeyIAAAAVot1DFbo/fv//4XAdA//dQjoMfv//16L5V3CDABXVo2FeP///1DoPAYAAIt9EI1FmFdQ6C8GAACNRdhQ6Ab7//+NRbjHRdgBAAAAUMdF3AAAAADo7/r//41FmFCNhXj///9Q6C/7//+L0IXSD4SwAQAAUw8fQACLjXj///8PV8CD4QFmDxNF+IPJAHUvjYV4////UOhuBQAAi0XYg+ABg8gAD4S2AAAAV41F2FBQ6FT5//+L8Iva6agAAACLRZiD4AGDyAB1LI1FmFDoNwUAA" & _
                                                    "ItFuIPgAYPIAA+ECAEAAFeNRbhQUOgd+f//i/CL2un6AAAAhdIPjowAAACNRZhQjYV4////UFDoOwcAAI2FeP///1Do7wQAAI1FuFCNRdhQ6HL6//+FwHkLV41F2FBQ6NP4//+NRbhQjUXYUFDoBQcAAItF2IPgAYPIAHQRV41F2FBQ6K/4//+L8Iva6waLXfyLdfiNRdhQ6JoEAAAL8w+EkgAAAItF8IFN9AAAAICJRfDpgAAAAI2FeP///1CNRZhQUOivBgAAjUWYUOhmBAAAjUXYUI1FuFDo6fn//4XAeQtXjUW4UFDoSvj//41F2FCNRbhQUOh8BgAAi0W4g+ABg8gAdBFXjUW4UFDoJvj//4vwi9rrBotd/It1+I1FuFDoEQQAAAvzdA2LRdCBTdQAAACAiUXQjUWYUI2FeP///1DogPn//4vQhdIPhVb+//9bjUXYUP91COg5BAAAX16L5V3CDADMVYvsgeyAAAAAU1aLdRRXVug7AwAA/3UQi9iNRYD/dQxQ6MoBAACNRaBQ6CEDAACL+IX/dAiBxwABAADrC41FgFDoCgMAAIv4O/tzFY1FgFD/dQjo2AMAAF9eW4vlXcIQAI1FwFDopvj//41F4FDonfj//4vHK8OL2MHrBoPgP3QYUI1FwFaNBNhQ6HH5//+JRN3giVTd5OsNjUXAVo0E2FDoigMAAItdCFPoYfj//8cDAQAAAMdDBAAAAAAPH0AAgf8AAQAAdw5WjUXAUOiO+P//hcB4c41FoFCNReBQ6H34//+FwHgTdTyNRYBQjUXAUOhq+P//hcB/K41FwFCNRYBQUOgIBQAAC8J0C1ONRaBQUOj5BAAAjUXgUI1FoFBQ6OsEAACLdeCNReBQweYf6JwCAACNRc" & _
                                                    "BQ6JMCAAAJddxPi3UU6Xf///+NRYBQU+jdAgAAX15bi+VdwhAAzMzMzFWL7IPsQI1FwP91EP91DFDoewAAAI1FwFD/dQjob/n//4vlXcIMAMzMzMzMzMzMzFWL7IPsQI1FwP91DFDozgIAAI1FwFD/dQjoQvn//4vlXcIIAMzMzMzMzMzMzMzMzFWL7Fb/dRCLdQj/dQxW6D0EAAALwnQK/3UUVlbo7/X//15dwhAAzMzMzMzMzMzMzFWL7IPsYFMPV8BWZg8TRdiLRdxXZg8TRdAz/4td1IlF/DP2jUf9g/8ED1fAZg8TRfSLVfQPQ/A79w+H0gAAAItNEIvHDxBF0CvGDxFFwI0cwYtF+IlF8IlV+GYPH0QAAIP+BA+DowAAAP9zBItFDP8z/3TwBP808I1FsFDof+H//4PsEIvMg+wQDxAADxAIi8QPEQEPEEXADxFN4A8RAI1FoFDoiKr//2YPc9kMDxAQZg9+yA8owmYPc9gMZg9+wQ8RVcCJTfwPEVXQO8h3E3IIi0XYO0Xocwm4AQAAADPJ6w4PV8BmDxNF6ItN7ItF6ItV+APQi0XwiVX4E8FGg+sIiUXwO/cPhlT///+LXdTrA4tF+ItNCIt10Ik0+Yvxi8qL0IlV3Ilc/gRHi3XYi138iXXQiV3UiU3YiVX8g/8HD4Lb/v//i0UIX4lwOF6JWDxbi+VdwgwAzMzMzMzMzMxVi+xWV4t9CFfoQgAAAIvwhfZ1Bl9eXcIEAItU9/iLyotE9/wz/wvIdBNmDx9EAAAPrMIBR9Hoi8oLyHXzweYGjUbAA8dfXl3CBADMzMzMzFWL7ItVCLgDAAAADx9EAACLDMILTMIEdQWD6AF58kBdwgQAzMzMzMzMzMzMzMzMzFWL7IP" & _
                                                    "sCItFCA9XwFOL2GYPE0X4g8AgO8N2OItN+FZXi338iU0Ii3D4g+gIi86LUAQPrNEBC00I0eoL14kIi/6JUATB5x/HRQgAAAAAO8N31V9eW4vlXcIEAMzMzMzMzFWL7ItVDItNCIsCiQGLQgSJQQSLQgiJQQiLQgyJQQyLQhCJQRCLQhSJQRSLQhiJQRiLQhyJQRxdwggAzMzMzMxVi+yD7GBTD1fAM8lWZg8TRdiLRdxXZg8TRdCLfdSJTeiJRfAz9o1B/YP5BA9XwGYPE0X4i138D0PwO/EPhxkBAACLVQyLwQ8QRdArxold9A8RRcCNBMKLVfiJReyJVfyL+Sv+O/cPh+oAAAD/cAT/MItFDP908AT/NPCNRbBQ6Pze//8PEAAPEUXQO/dzQ4tN3IvBi1XUi/rB6B8BRfyLRdiD0wDB7x8PpMEBiV30M9sDwAvZC/iJXdyLRdAPpMIBiX3YA8CJVdSJRdAPEEXQ6waLXdyLfdiD7BCLxIPsEA8RAIvEDxBFwA8RAI1FoFDou6f//w8QCA8owWYPc9gMZg9+wA8RTcCJRfAPEU3QO8N3EHIFOX3Ycwm4AQAAADPJ6w4PV8BmDxNF4ItN5ItF4ItV/Itd9APQi0XsE9mJVfyLTehGg+gIiV30iUXsO/EPhgr///+LfdTrA4tV+It1CItF0IkEzotF2Il8zgRBi33wiVXYi9OJRdCJfdSJVfCJVdyJTeiD+QcPgpX+//+JfjxfiUY4XluL5V3CCADMzFWL7IPsDFOLXQwPV8BWV4t9EIsTi/KLQwSLyGYPE0X0KzcbTwQ78nUGO8h1BOsYO8hyD3cEO/J2CbgBAAAAM9LrC2YPE0X0i0X0i1X4i30IiU8Ei00QiTeLcwiJdfgrcQiL" & _
                                                    "SwyLXRAbSwwr8ItdDBvKO3X4dQU7Swx0IDtLDHIQdwU7cwh2CbgBAAAAM9LrC2YPE0X0i1X4i0X0iU8Mi00QiXcIi3MQiXX8K3EQi0sUi10QG0sUK/CLXQwbyjt1/HUFO0sUdCA7SxRyEHcFO3MQdgm4AQAAADPS6wtmDxNF9ItV+ItF9IlPFIl3EItLGIvxi30Qi1sciU0Mi00QK3EYi8sbTxwr8It9CBvKO3UMdQQ7y3QsO8tyHXcFO3UMdhaJdxi4AQAAAIlPHDPSX15bi+VdwgwAZg8TRfSLVfiLRfSJdxiJTxxfXluL5V3CDADMzMzMzMzMzMzMzMzMzMxVi+yLTQgz0lZXi30MM/aLx4PgPw+rxoP4IA9D1jPyg/hAD0PWwe8GIzT5I1T5BIvGX15dwggAzMzMzMzMzMzMVYvsi1UUg+wQM8mF0g+EwgAAAFOLXRBWi3UIV4t9DIP6IA+CiwAAAI1D/wPCO/B3CY1G/wPCO8NzeY1H/wPCO/B3CY1G/wPCO8dzZ4vCi9cr04Pg4IlV/IvWK9OJRfCJVfiLw4td+IvXi338K9aJVfSNVhAPEACLdfSDwSCNQCCNUiAPEEwH4GYP78gPEUwD4A8QTBbgi3UIDxBA8GYP78gPEUrgO03wcsqLVRSLfQyLXRA7ynMbK/uNBBkr8yvRigw4jUABMkj/iEww/4PqAXXuX15bi+VdwhAAAAA=" ' 25325, 22.4.2020 14:28:07
Private Const STR_LIBSODIUM_SHA384_STATE As String = "2J4FwV2du8sH1Xw2KimaYhfdcDBaAVmROVkO99jsLxUxC8D/ZyYzZxEVWGiHSrSOp4/5ZA0uDNukT/q+HUi1Rw=="
'--- numeric
Private Const LNG_SHA256_HASHSZ         As Long = 32
Private Const LNG_SHA256_BLOCKSZ        As Long = 64
Private Const LNG_SHA384_HASHSZ         As Long = 48
Private Const LNG_SHA384_BLOCKSZ        As Long = 128
Private Const LNG_SHA384_CONTEXTSZ      As Long = 200
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

Private m_uData                    As UcsCryptoThunkData

Private Enum UcsThunkPfnIndexEnum
    ucsPfnSecp256r1MakeKey = 1
    ucsPfnSecp256r1SharedSecret
    ucsPfnSecp256r1UncompressKey
    ucsPfnSecp256r1Sign
    ucsPfnSecp256r1Verify
    ucsPfnCurve25519ScalarMultiply
    ucsPfnCurve25519ScalarMultBase
    ucsPfnSha256Init
    ucsPfnSha256Update
    ucsPfnSha256Final
    ucsPfnSha384Init
    ucsPfnSha384Update
    ucsPfnSha384Final
    ucsPfnChacha20Poly1305Encrypt
    ucsPfnChacha20Poly1305Decrypt
    ucsPfnAesGcmEncrypt
    ucsPfnAesGcmDecrypt
    [_ucsPfnMax]
End Enum

Private Type UcsCryptoThunkData
    Thunk               As Long
    Glob()              As Byte
    Pfn(1 To [_ucsPfnMax] - 1) As Long
    EccKeySize          As Long
#If ImplUseLibSodium Then
    HashCtx(0 To LNG_LIBSODIUM_SHA512_CONTEXTSZ - 1) As Byte
#Else
    HashCtx(0 To LNG_SHA384_CONTEXTSZ - 1) As Byte
#End If
    HashPad(0 To LNG_SHA384_BLOCKSZ - 1 + 1000) As Byte
    HashFinal(0 To LNG_SHA384_HASHSZ - 1 + 1000) As Byte
    hRandomProv         As Long
#If ImplUseBCrypt Then
    hEcdhP256Prov       As Long
#End If
End Type

Public Type UcsRsaContextType
    hProv               As Long
    hPrivKey            As Long
    hPubKey             As Long
    HashAlgId           As Long
End Type

'=========================================================================
' Functions
'=========================================================================

Public Function CryptoInit() As Boolean
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
        #If ImplUseBCrypt Then
            If .hEcdhP256Prov = 0 Then
                hResult = BCryptOpenAlgorithmProvider(.hEcdhP256Prov, StrPtr("ECDH_P256"), StrPtr("Microsoft Primitive Provider"), 0)
                If hResult < 0 Then
                    sApiSource = "BCryptOpenAlgorithmProvider"
                    GoTo QH
                End If
            End If
        #End If
        If m_uData.Thunk = 0 Then
            .EccKeySize = 32
            '--- prepare thunk/context in executable memory
            .Thunk = pvThunkAllocate(STR_THUNK1 & STR_THUNK2 & STR_THUNK3)
            If .Thunk = 0 Then
                hResult = ERR_OUT_OF_MEMORY
                sApiSource = "VirtualAlloc"
                GoTo QH
            End If
            ReDim .Glob(0 To (Len(STR_GLOB) \ 4) * 3 - 1) As Byte
            pvThunkAllocate STR_GLOB, VarPtr(.Glob(0))
            '--- init pfns from thunk addr + offsets stored at beginning of it
            For lIdx = LBound(.Pfn) To UBound(.Pfn)
                Call CopyMemory(lOffset, ByVal UnsignedAdd(.Thunk, 4 * lIdx), 4)
                .Pfn(lIdx) = UnsignedAdd(.Thunk, lOffset)
            Next
            '--- init pfns trampolines
            Call pvPatchTrampoline(AddressOf pvCryptoCallSecp256r1MakeKey)
            Call pvPatchTrampoline(AddressOf pvCryptoCallSecp256r1SharedSecret)
            Call pvPatchTrampoline(AddressOf pvCryptoCallSecp256r1UncompressKey)
            Call pvPatchTrampoline(AddressOf pvCryptoCallSecp256r1Sign)
            Call pvPatchTrampoline(AddressOf pvCryptoCallSecp256r1Verify)
            Call pvPatchTrampoline(AddressOf pvCryptoCallCurve25519Multiply)
            Call pvPatchTrampoline(AddressOf pvCryptoCallCurve25519MulBase)
            Call pvPatchTrampoline(AddressOf pvCryptoCallSha256Init)
            Call pvPatchTrampoline(AddressOf pvCryptoCallSha256Update)
            Call pvPatchTrampoline(AddressOf pvCryptoCallSha256Final)
            Call pvPatchTrampoline(AddressOf pvCryptoCallSha384Init)
            Call pvPatchTrampoline(AddressOf pvCryptoCallSha384Update)
            Call pvPatchTrampoline(AddressOf pvCryptoCallSha384Final)
            Call pvPatchTrampoline(AddressOf pvCryptoCallChacha20Poly1305Encrypt)
            Call pvPatchTrampoline(AddressOf pvCryptoCallChacha20Poly1305Decrypt)
            Call pvPatchTrampoline(AddressOf pvCryptoCallAesGcmEncrypt)
            Call pvPatchTrampoline(AddressOf pvCryptoCallAesGcmDecrypt)
            '--- init thunk's first 4 bytes -> global data in C/C++
            Call CopyMemory(ByVal .Thunk, VarPtr(.Glob(0)), 4)
        End If
    End With
    '--- success
    CryptoInit = True
QH:
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), sApiSource
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
        #If ImplUseBCrypt Then
            If .hEcdhP256Prov <> 0 Then
                Call BCryptCloseAlgorithmProvider(.hEcdhP256Prov, 0)
                .hEcdhP256Prov = 0
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
    Case Else
        CryptoIsSupported = True
    End Select
End Function

Public Function CryptoEccSecp256r1MakeKey(baPrivate() As Byte, baPublic() As Byte) As Boolean
    Const MAX_RETRIES   As Long = 16
    Dim lIdx            As Long
    
    #If ImplUseBCrypt Then
        CryptoEccSecp256r1MakeKey = pvBCryptEcdhP256KeyPair(baPrivate, baPublic)
    #Else
        ReDim baPrivate(0 To m_uData.EccKeySize - 1) As Byte
        ReDim baPublic(0 To m_uData.EccKeySize) As Byte
        For lIdx = 1 To MAX_RETRIES
            CryptoRandomBytes VarPtr(baPrivate(0)), m_uData.EccKeySize
            If pvCryptoCallSecp256r1MakeKey(m_uData.Pfn(ucsPfnSecp256r1MakeKey), baPublic(0), baPrivate(0)) = 1 Then
                Exit For
            End If
        Next
        '--- success (or failure)
        CryptoEccSecp256r1MakeKey = (lIdx <= MAX_RETRIES)
    #End If
End Function

Public Function CryptoEccSecp256r1SharedSecret(baPrivate() As Byte, baPublic() As Byte) As Byte()
    Dim baRetVal()      As Byte
    
    #If ImplUseBCrypt Then
        Debug.Assert pvArraySize(baPrivate) = BCRYPT_SECP256R1_PRIVATE_KEYSZ
        Debug.Assert pvArraySize(baPublic) = BCRYPT_SECP256R1_COMPRESSED_PUBLIC_KEYSZ Or pvArraySize(baPublic) = BCRYPT_SECP256R1_UNCOMPRESSED_PUBLIC_KEYSZ
        baRetVal = pvBCryptEcdhP256AgreedSecret(baPrivate, baPublic)
    #Else
        Debug.Assert UBound(baPrivate) >= m_uData.EccKeySize - 1
        Debug.Assert UBound(baPublic) >= m_uData.EccKeySize
        ReDim baRetVal(0 To m_uData.EccKeySize - 1) As Byte
        If pvCryptoCallSecp256r1SharedSecret(m_uData.Pfn(ucsPfnSecp256r1SharedSecret), baPublic(0), baPrivate(0), baRetVal(0)) = 0 Then
            GoTo QH
        End If
    #End If
    CryptoEccSecp256r1SharedSecret = baRetVal
QH:
End Function

Public Function CryptoEccSecp256r1UncompressKey(baPublic() As Byte) As Byte()
    Dim baRetVal()      As Byte
    
    ReDim baRetVal(0 To 2 * m_uData.EccKeySize) As Byte
    If pvCryptoCallSecp256r1UncompressKey(m_uData.Pfn(ucsPfnSecp256r1UncompressKey), baPublic(0), baRetVal(0)) = 0 Then
        GoTo QH
    End If
    CryptoEccSecp256r1UncompressKey = baRetVal
QH:
End Function

Public Function CryptoEccSecp256r1Sign(baPrivate() As Byte, baHash() As Byte) As Byte()
    Const MAX_RETRIES   As Long = 16
    Dim baRandom()      As Byte
    Dim baRetVal()      As Byte
    Dim lIdx            As Long
    
    ReDim baRandom(0 To m_uData.EccKeySize - 1) As Byte
    ReDim baRetVal(0 To 2 * m_uData.EccKeySize - 1) As Byte
    For lIdx = 1 To MAX_RETRIES
        CryptoRandomBytes VarPtr(baRandom(0)), m_uData.EccKeySize
        If pvCryptoCallSecp256r1Sign(m_uData.Pfn(ucsPfnSecp256r1Sign), baPrivate(0), baHash(0), baRandom(0), baRetVal(0)) <> 0 Then
            Exit For
        End If
    Next
    If lIdx < MAX_RETRIES Then
        '--- success
        CryptoEccSecp256r1Sign = baRetVal
    End If
End Function

Public Function CryptoEccSecp256r1Verify(baPublic() As Byte, baHash() As Byte, baSignature() As Byte) As Boolean
    CryptoEccSecp256r1Verify = (pvCryptoCallSecp256r1Verify(m_uData.Pfn(ucsPfnSecp256r1Verify), baPublic(0), baHash(0), baSignature(0)) <> 0)
End Function

Public Function CryptoEccCurve25519MakeKey(baPrivate() As Byte, baPublic() As Byte) As Boolean
    ReDim baPrivate(0 To m_uData.EccKeySize - 1) As Byte
    ReDim baPublic(0 To m_uData.EccKeySize - 1) As Byte
    CryptoRandomBytes VarPtr(baPrivate(0)), m_uData.EccKeySize
    '--- fix issues w/ specific privkeys
    baPrivate(0) = baPrivate(0) And 248
    baPrivate(UBound(baPrivate)) = (baPrivate(UBound(baPrivate)) And 127) Or 64
    #If ImplUseLibSodium Then
        Call crypto_scalarmult_curve25519_base(baPublic(0), baPrivate(0))
    #Else
        pvCryptoCallCurve25519MulBase m_uData.Pfn(ucsPfnCurve25519ScalarMultBase), baPublic(0), baPrivate(0)
    #End If
    '--- success
    CryptoEccCurve25519MakeKey = True
End Function

Public Function CryptoEccCurve25519SharedSecret(baPrivate() As Byte, baPublic() As Byte) As Byte()
    Dim baRetVal()      As Byte
    
    Debug.Assert UBound(baPrivate) >= m_uData.EccKeySize - 1
    Debug.Assert UBound(baPublic) >= m_uData.EccKeySize - 1
    ReDim baRetVal(0 To m_uData.EccKeySize - 1) As Byte
    #If ImplUseLibSodium Then
        Call crypto_scalarmult_curve25519(baRetVal(0), baPrivate(0), baPublic(0))
    #Else
        pvCryptoCallCurve25519Multiply m_uData.Pfn(ucsPfnCurve25519ScalarMultiply), baRetVal(0), baPrivate(0), baPublic(0)
    #End If
    CryptoEccCurve25519SharedSecret = baRetVal
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
    #If ImplUseLibSodium Then
        Call crypto_hash_sha256(baRetVal(0), ByVal lPtr, Size)
    #Else
        With m_uData
            lCtxPtr = VarPtr(.HashCtx(0))
            pvCryptoCallSha256Init .Pfn(ucsPfnSha256Init), lCtxPtr
            pvCryptoCallSha256Update .Pfn(ucsPfnSha256Update), lCtxPtr, lPtr, Size
            pvCryptoCallSha256Final .Pfn(ucsPfnSha256Final), lCtxPtr, baRetVal(0)
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
    With m_uData
        lCtxPtr = VarPtr(.HashCtx(0))
        #If ImplUseLibSodium Then
            Call crypto_hash_sha384_init(.HashCtx)
            Call crypto_hash_sha512_update(ByVal lCtxPtr, ByVal lPtr, Size)
            Call crypto_hash_sha512_final(ByVal lCtxPtr, .HashFinal(0))
            Call CopyMemory(baRetVal(0), .HashFinal(0), LNG_SHA384_HASHSZ)
        #Else
            pvCryptoCallSha384Init .Pfn(ucsPfnSha384Init), lCtxPtr
            pvCryptoCallSha384Update .Pfn(ucsPfnSha384Update), lCtxPtr, lPtr, Size
            pvCryptoCallSha384Final .Pfn(ucsPfnSha384Final), lCtxPtr, baRetVal(0)
        #End If
    End With
    CryptoHashSha384 = baRetVal
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
            '-- inner hash
            pvCryptoCallSha256Init .Pfn(ucsPfnSha256Init), lCtxPtr
            Call FillMemory(.HashPad(0), LNG_SHA256_BLOCKSZ, LNG_HMAC_INNER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_INNER_PAD
            Next
            pvCryptoCallSha256Update .Pfn(ucsPfnSha256Update), lCtxPtr, VarPtr(.HashPad(0)), LNG_SHA256_BLOCKSZ
            pvCryptoCallSha256Update .Pfn(ucsPfnSha256Update), lCtxPtr, lPtr, Size
            pvCryptoCallSha256Final .Pfn(ucsPfnSha256Final), lCtxPtr, .HashFinal(0)
            '-- outer hash
            pvCryptoCallSha256Init .Pfn(ucsPfnSha256Init), lCtxPtr
            Call FillMemory(.HashPad(0), LNG_SHA256_BLOCKSZ, LNG_HMAC_OUTER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_OUTER_PAD
            Next
            pvCryptoCallSha256Update .Pfn(ucsPfnSha256Update), lCtxPtr, VarPtr(.HashPad(0)), LNG_SHA256_BLOCKSZ
            pvCryptoCallSha256Update .Pfn(ucsPfnSha256Update), lCtxPtr, VarPtr(.HashFinal(0)), LNG_SHA256_HASHSZ
            pvCryptoCallSha256Final .Pfn(ucsPfnSha256Final), lCtxPtr, baRetVal(0)
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
            '-- inner hash
            pvCryptoCallSha384Init .Pfn(ucsPfnSha384Init), lCtxPtr
            Call FillMemory(.HashPad(0), LNG_SHA384_BLOCKSZ, LNG_HMAC_INNER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_INNER_PAD
            Next
            pvCryptoCallSha384Update .Pfn(ucsPfnSha384Update), lCtxPtr, VarPtr(.HashPad(0)), LNG_SHA384_BLOCKSZ
            pvCryptoCallSha384Update .Pfn(ucsPfnSha384Update), lCtxPtr, lPtr, Size
            pvCryptoCallSha384Final .Pfn(ucsPfnSha384Final), lCtxPtr, .HashFinal(0)
            '-- outer hash
            pvCryptoCallSha384Init .Pfn(ucsPfnSha384Init), lCtxPtr
            Call FillMemory(.HashPad(0), LNG_SHA384_BLOCKSZ, LNG_HMAC_OUTER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_OUTER_PAD
            Next
            pvCryptoCallSha384Update .Pfn(ucsPfnSha384Update), lCtxPtr, VarPtr(.HashPad(0)), LNG_SHA384_BLOCKSZ
            pvCryptoCallSha384Update .Pfn(ucsPfnSha384Update), lCtxPtr, VarPtr(.HashFinal(0)), LNG_SHA384_HASHSZ
            pvCryptoCallSha384Final .Pfn(ucsPfnSha384Final), lCtxPtr, baRetVal(0)
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
            Call pvCryptoCallChacha20Poly1305Encrypt(m_uData.Pfn(ucsPfnChacha20Poly1305Encrypt), _
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
        If pvCryptoCallChacha20Poly1305Decrypt(m_uData.Pfn(ucsPfnChacha20Poly1305Decrypt), _
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
            Call pvCryptoCallAesGcmEncrypt(m_uData.Pfn(ucsPfnAesGcmEncrypt), _
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
        If pvCryptoCallAesGcmDecrypt(m_uData.Pfn(ucsPfnAesGcmDecrypt), _
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

'= RSA helpers ===========================================================

Public Function CryptoRsaInitContext(uCtx As UcsRsaContextType, baPrivKey() As Byte, baCert() As Byte, baPubKey() As Byte, Optional ByVal SignatureType As Long) As Boolean
    Dim lHashAlgId      As Long
    Dim hProv           As Long
    Dim lPkiPtr         As Long
    Dim lKeyPtr         As Long
    Dim lKeySize        As Long
    Dim uKeyBlob        As CRYPT_DER_BLOB
    Dim hPrivKey        As Long
    Dim pContext        As Long
    Dim lPtr            As Long
    Dim hPubKey         As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    
    If SignatureType = TLS_SIGNATURE_RSA_PKCS1_SHA1 Then
        lHashAlgId = CALG_SHA1
    ElseIf SignatureType = TLS_SIGNATURE_RSA_PKCS1_SHA256 Then
        lHashAlgId = CALG_SHA_256
    ElseIf SignatureType <> 0 Then
        GoTo QH
    End If
    If CryptAcquireContext(hProv, 0, 0, IIf(lHashAlgId = CALG_SHA_256, PROV_RSA_AES, PROV_RSA_FULL), CRYPT_VERIFYCONTEXT) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptAcquireContext"
        GoTo QH
    End If
    If pvArraySize(baPrivKey) > 0 Then
        If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_PRIVATE_KEY_INFO, baPrivKey(0), UBound(baPrivKey) + 1, CRYPT_DECODE_ALLOC_FLAG, 0, lPkiPtr, 0) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptDecodeObjectEx(PKCS_PRIVATE_KEY_INFO)"
            GoTo QH
        End If
        Call CopyMemory(uKeyBlob, ByVal UnsignedAdd(lPkiPtr, 16), Len(uKeyBlob)) '--- dereference PCRYPT_PRIVATE_KEY_INFO->PrivateKey
        If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_RSA_PRIVATE_KEY, ByVal uKeyBlob.pbData, uKeyBlob.cbData, CRYPT_DECODE_ALLOC_FLAG, 0, lKeyPtr, lKeySize) = 0 Then
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
        pContext = CertCreateCertificateContext(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, baCert(0), UBound(baCert) + 1)
        If pContext = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CertCreateCertificateContext"
            GoTo QH
        End If
        Call CopyMemory(lPtr, ByVal UnsignedAdd(pContext, 12), 4)       '--- dereference pContext->pCertInfo
        lPtr = UnsignedAdd(lPtr, 56)                                    '--- &pContext->pCertInfo->SubjectPublicKeyInfo
        If CryptImportPublicKeyInfo(hProv, X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, ByVal lPtr, hPubKey) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptImportPublicKeyInfo#1"
            GoTo QH
        End If
    ElseIf pvArraySize(baPubKey) > 0 Then
        If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, X509_PUBLIC_KEY_INFO, baPubKey(0), UBound(baPubKey) + 1, CRYPT_DECODE_ALLOC_FLAG, 0, lKeyPtr, 0) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptDecodeObjectEx(PKCS_PRIVATE_KEY_INFO)"
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
    If pContext <> 0 Then
        Call CertFreeCertificateContext(pContext)
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
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), sApiSource
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

Public Function CryptoRsaSign(uCtx As UcsRsaContextType, baPlainText() As Byte) As Byte()
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
    lSize = pvArraySize(baPlainText)
    If lSize > 0 Then
        If CryptHashData(hHash, baPlainText(0), lSize, 0) = 0 Then
            hResult = Err.LastDllError
            sApiSource = "CryptHashData"
            GoTo QH
        End If
    End If
    ReDim baRetVal(0 To MAX_SIG_SIZE - 1) As Byte
    lSize = UBound(baRetVal) + 1
    If CryptSignHash(hHash, AT_KEYEXCHANGE, 0, 0, baRetVal(0), lSize) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptSignHash"
        GoTo QH
    End If
    If UBound(baRetVal) <> lSize - 1 Then
        ReDim Preserve baRetVal(0 To lSize - 1) As Byte
    End If
    pvArrayReverse baRetVal
    CryptoRsaSign = baRetVal
QH:
    If hHash <> 0 Then
        Call CryptDestroyHash(hHash)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), sApiSource
    End If
End Function

Public Function CryptoRsaVerify(uCtx As UcsRsaContextType, baPlainText() As Byte, baSignature() As Byte) As Boolean
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
    lSize = pvArraySize(baPlainText)
    If lSize > 0 Then
        If CryptHashData(hHash, baPlainText(0), lSize, 0) = 0 Then
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
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), sApiSource
    End If
End Function

Public Function CryptoRsaExtractPublicKey(baCert() As Byte) As Byte()
    Dim pContext        As Long
    Dim hProv           As Long
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    Dim lPtr            As Long
    Dim hResult         As Long
    Dim sApiSource      As String

    If CryptAcquireContext(hProv, 0, 0, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptAcquireContext"
        GoTo QH
    End If
    pContext = CertCreateCertificateContext(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, baCert(0), UBound(baCert) + 1)
    If pContext = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CertCreateCertificateContext"
        GoTo QH
    End If
    Call CopyMemory(lPtr, ByVal UnsignedAdd(pContext, 12), 4)       '--- dereference pContext->pCertInfo
    lPtr = UnsignedAdd(lPtr, 56)                                    '--- &pContext->pCertInfo->SubjectPublicKeyInfo
    '--- get required size first
    If CryptEncodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, X509_PUBLIC_KEY_INFO, ByVal lPtr, 0, 0, ByVal 0, lSize) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptEncodeObjectEx"
        GoTo QH
    End If
    ReDim baRetVal(0 To lSize - 1) As Byte
    If CryptEncodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, X509_PUBLIC_KEY_INFO, ByVal lPtr, 0, 0, baRetVal(0), lSize) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptEncodeObjectEx"
        GoTo QH
    End If
    If UBound(baRetVal) <> lSize - 1 Then
        ReDim Preserve baRetVal(0 To lSize - 1) As Byte
    End If
    CryptoRsaExtractPublicKey = baRetVal
QH:
    If hProv <> 0 Then
        Call CryptReleaseContext(hProv, 0)
    End If
    If pContext <> 0 Then
        Call CertFreeCertificateContext(pContext)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), sApiSource
    End If
End Function

Public Function CryptoRsaEncrypt(ByVal hKey As Long, baPlainText() As Byte) As Byte()
    Const MAX_RSA_BYTES As Long = MAX_RSA_KEY / 8
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    Dim lAlignedSize    As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    
    lSize = pvArraySize(baPlainText)
    lAlignedSize = lSize + MAX_RSA_BYTES - 1 And -MAX_RSA_BYTES
    ReDim baRetVal(0 To lAlignedSize - 1) As Byte
    Call CopyMemory(baRetVal(0), baPlainText(0), lSize)
    If CryptEncrypt(hKey, 0, 1, 0, baRetVal(0), lSize, lAlignedSize) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptEncrypt"
        GoTo QH
    End If
    ReDim Preserve baRetVal(0 To lSize - 1) As Byte
    pvArrayReverse baRetVal
    CryptoRsaEncrypt = baRetVal
QH:
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), sApiSource
    End If
End Function

Public Function CryptoRsaPssSign(baPrivKey() As Byte, baMessage() As Byte, ByVal lSignatureType As Long) As Byte()
    Dim baRetVal()      As Byte
    Dim lPkiPtr         As Long
    Dim lKeyPtr         As Long
    Dim lKeySize        As Long
    Dim uKeyBlob        As CRYPT_DER_BLOB
    Dim hAlgRSA         As Long
    Dim hKey            As Long
    Dim uPadInfo        As BCRYPT_PSS_PADDING_INFO
    Dim lSize           As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    
    If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_PRIVATE_KEY_INFO, baPrivKey(0), UBound(baPrivKey) + 1, CRYPT_DECODE_ALLOC_FLAG, 0, lPkiPtr, 0) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptDecodeObjectEx(PKCS_PRIVATE_KEY_INFO)"
        GoTo QH
    End If
    Call CopyMemory(uKeyBlob, ByVal UnsignedAdd(lPkiPtr, 16), Len(uKeyBlob)) '--- dereference PCRYPT_PRIVATE_KEY_INFO->PrivateKey
    If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, PKCS_RSA_PRIVATE_KEY, ByVal uKeyBlob.pbData, uKeyBlob.cbData, CRYPT_DECODE_ALLOC_FLAG, 0, lKeyPtr, lKeySize) = 0 Then
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
    If lSignatureType = TLS_SIGNATURE_RSA_PSS_RSAE_SHA256 Then
        uPadInfo.pszAlgId = StrPtr("SHA256")
        uPadInfo.cbSalt = 32
    ElseIf lSignatureType = TLS_SIGNATURE_RSA_PSS_RSAE_SHA384 Then
        uPadInfo.pszAlgId = StrPtr("SHA384")
        uPadInfo.cbSalt = 48
    End If
    ReDim baRetVal(0 To 1023) As Byte
    hResult = BCryptSignHash(hKey, uPadInfo, baMessage(0), UBound(baMessage) + 1, baRetVal(0), UBound(baRetVal) + 1, lSize, BCRYPT_PAD_PSS)
    If hResult < 0 Then
        sApiSource = "BCryptSignHash"
        GoTo QH
    End If
    ReDim Preserve baRetVal(0 To lSize - 1) As Byte
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
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), sApiSource
    End If
End Function

Public Function CryptoRsaPssVerify(baPubKey() As Byte, baMessage() As Byte, baSignature() As Byte, ByVal lSignatureType As Long) As Boolean
    Dim lKeyPtr         As Long
    Dim hKey            As Long
    Dim uPadInfo        As BCRYPT_PSS_PADDING_INFO
    Dim hResult         As Long
    Dim sApiSource      As String
    
    If CryptDecodeObjectEx(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, X509_PUBLIC_KEY_INFO, baPubKey(0), UBound(baPubKey) + 1, CRYPT_DECODE_ALLOC_FLAG, 0, lKeyPtr, 0) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptDecodeObjectEx(PKCS_PRIVATE_KEY_INFO)"
        GoTo QH
    End If
    If CryptImportPublicKeyInfoEx2(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, lKeyPtr, 0, 0, hKey) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptImportPublicKeyInfoEx2"
        GoTo QH
    End If
    If lSignatureType = TLS_SIGNATURE_RSA_PSS_RSAE_SHA256 Then
        uPadInfo.pszAlgId = StrPtr("SHA256")
        uPadInfo.cbSalt = 32
    ElseIf lSignatureType = TLS_SIGNATURE_RSA_PSS_RSAE_SHA384 Then
        uPadInfo.pszAlgId = StrPtr("SHA384")
        uPadInfo.cbSalt = 48
    End If
    hResult = BCryptVerifySignature(hKey, uPadInfo, baMessage(0), UBound(baMessage) + 1, baSignature(0), UBound(baSignature) + 1, BCRYPT_PAD_PSS)
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
    If lKeyPtr <> 0 Then
        Call LocalFree(lKeyPtr)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), sApiSource
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

'--- PEM = privacy-enhanced mail
Public Function CryptoPemTextPortions(sContents As String, sBoundary As String, Optional RetVal As Collection) As Collection
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
    Set CryptoPemTextPortions = RetVal
End Function

Public Function FromBase64Array(sText As String) As Byte()
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    lSize = (Len(sText) \ 4) * 3
    ReDim baRetVal(0 To lSize - 1) As Byte
    pvThunkAllocate sText, VarPtr(baRetVal(0))
    If Right$(sText, 2) = "==" Then
        ReDim Preserve baRetVal(0 To lSize - 3)
    ElseIf Right$(sText, 1) = "=" Then
        ReDim Preserve baRetVal(0 To lSize - 2)
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

Private Sub pvPatchTrampoline(ByVal Pfn As Long)
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
End Sub

Private Sub pvPatchMethodTrampoline(ByVal Pfn As Long, ByVal lMethodIdx As Long)
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
End Sub

Private Function pvSetTrue(bValue As Boolean) As Boolean
    bValue = True
    pvSetTrue = True
End Function

Private Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function

#If ImplUseLibSodium Then
    Private Sub crypto_hash_sha384_init(baCtx() As Byte)
        Static baSha384State() As Byte
        
        If pvArraySize(baSha384State) = 0 Then
            ReDim baSha384State(0 To (Len(STR_LIBSODIUM_SHA384_STATE) \ 4) * 3 - 1) As Byte
            pvThunkAllocate STR_LIBSODIUM_SHA384_STATE, VarPtr(baSha384State(0))
            ReDim Preserve baSha384State(0 To 63) As Byte
        End If
        Debug.Assert pvArraySize(baCtx) >= crypto_hash_sha512_statebytes()
        Call crypto_hash_sha512_init(baCtx(0))
        Call CopyMemory(baCtx(0), baSha384State(0), UBound(baSha384State) + 1)
    End Sub
#End If

'= BCrypt helpers ========================================================

#If ImplUseBCrypt Then
Private Function pvBCryptEcdhP256KeyPair(baPriv() As Byte, baPub() As Byte) As Boolean
    Dim hProv           As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    Dim hKeyPair        As Long
    Dim baBlob()        As Byte
    Dim cbResult        As Long
    
    hProv = m_uData.hEcdhP256Prov
    hResult = BCryptGenerateKeyPair(hProv, hKeyPair, 256, 0)
    If hResult < 0 Then
        sApiSource = "BCryptGenerateKeyPair"
        GoTo QH
    End If
    hResult = BCryptFinalizeKeyPair(hKeyPair, 0)
    If hResult < 0 Then
        sApiSource = "BCryptFinalizeKeyPair"
        GoTo QH
    End If
    ReDim baBlob(0 To 1023) As Byte
    hResult = BCryptExportKey(hKeyPair, 0, StrPtr("ECCPRIVATEBLOB"), VarPtr(baBlob(0)), UBound(baBlob) + 1, cbResult, 0)
    If hResult < 0 Then
        sApiSource = "BCryptExportKey(ECCPRIVATEBLOB)"
        GoTo QH
    End If
    baPriv = pvBCryptFromKeyBlob(baBlob, cbResult)
    hResult = BCryptExportKey(hKeyPair, 0, StrPtr("ECCPUBLICBLOB"), VarPtr(baBlob(0)), UBound(baBlob) + 1, cbResult, 0)
    If hResult < 0 Then
        sApiSource = "BCryptExportKey(ECCPUBLICBLOB)"
        GoTo QH
    End If
    baPub = pvBCryptFromKeyBlob(baBlob, cbResult)
    '--- success
    pvBCryptEcdhP256KeyPair = True
QH:
    If hKeyPair <> 0 Then
        Call BCryptDestroyKey(hKeyPair)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), sApiSource
    End If
End Function

Private Function pvBCryptEcdhP256AgreedSecret(baPriv() As Byte, baPub() As Byte) As Byte()
    Dim baRetVal()      As Byte
    Dim hProv           As Long
    Dim hPrivKey        As Long
    Dim hPubKey         As Long
    Dim hAgreedSecret   As Long
    Dim cbAgreedSecret  As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    Dim baBlob()        As Byte
    
    hProv = m_uData.hEcdhP256Prov
    baBlob = pvBCryptToKeyBlob(baPriv)
    hResult = BCryptImportKeyPair(hProv, 0, StrPtr("ECCPRIVATEBLOB"), hPrivKey, VarPtr(baBlob(0)), UBound(baBlob) + 1, 0)
    If hResult < 0 Then
        sApiSource = "BCryptImportKeyPair(ECCPRIVATEBLOB)"
        GoTo QH
    End If
    baBlob = pvBCryptToKeyBlob(baPub)
    hResult = BCryptImportKeyPair(hProv, 0, StrPtr("ECCPUBLICBLOB"), hPubKey, VarPtr(baBlob(0)), UBound(baBlob) + 1, 0)
    If hResult < 0 Then
        sApiSource = "BCryptImportKeyPair(ECCPUBLICBLOB)"
        GoTo QH
    End If
    hResult = BCryptSecretAgreement(hPrivKey, hPubKey, hAgreedSecret, 0)
    If hResult < 0 Then
        sApiSource = "BCryptSecretAgreement"
        GoTo QH
    End If
    ReDim baRetVal(0 To 1023) As Byte
    hResult = BCryptDeriveKey(hAgreedSecret, StrPtr("TRUNCATE"), 0, VarPtr(baRetVal(0)), UBound(baRetVal) + 1, cbAgreedSecret, 0)
    If hResult < 0 Then
        sApiSource = "BCryptDeriveKey"
        GoTo QH
    End If
    ReDim Preserve baRetVal(0 To cbAgreedSecret - 1) As Byte
    pvArrayReverse baRetVal
    pvBCryptEcdhP256AgreedSecret = baRetVal
QH:
    If hAgreedSecret <> 0 Then
        Call BCryptDestroySecret(hAgreedSecret)
    End If
    If hPrivKey <> 0 Then
        Call BCryptDestroyKey(hPrivKey)
    End If
    If hPubKey <> 0 Then
        Call BCryptDestroyKey(hPubKey)
    End If
    If LenB(sApiSource) <> 0 Then
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), sApiSource
    End If
End Function

Private Function pvBCryptToKeyBlob(baKey() As Byte, Optional ByVal lSize As Long = -1) As Byte()
    Dim baRetVal()      As Byte
    Dim lMagic          As Long
    Dim baUncompr()     As Byte
    Dim lKeyPtr         As Long
    
    If lSize < 0 Then
        lSize = pvArraySize(baKey)
    End If
    If lSize = BCRYPT_SECP256R1_COMPRESSED_PUBLIC_KEYSZ Then
        Debug.Assert baKey(0) = BCRYPT_SECP256R1_TAG_COMPRESSED_POS Or baKey(0) = BCRYPT_SECP256R1_TAG_COMPRESSED_NEG
        lMagic = BCRYPT_ECDH_PUBLIC_P256_MAGIC
        lSize = BCRYPT_SECP256R1_UNCOMPRESSED_PUBLIC_KEYSZ
        ReDim baUncompr(0 To lSize - 1) As Byte
        Call pvCryptoCallSecp256r1UncompressKey(m_uData.Pfn(ucsPfnSecp256r1UncompressKey), baKey(0), baUncompr(0))
        lKeyPtr = VarPtr(baUncompr(1))
        lSize = lSize - 1
    ElseIf lSize = BCRYPT_SECP256R1_UNCOMPRESSED_PUBLIC_KEYSZ Then
        Debug.Assert baKey(0) = BCRYPT_SECP256R1_TAG_UNCOMPRESSED
        lMagic = BCRYPT_ECDH_PUBLIC_P256_MAGIC
        lKeyPtr = VarPtr(baKey(1))
        lSize = lSize - 1
    ElseIf lSize = BCRYPT_SECP256R1_PRIVATE_KEYSZ Then
        lMagic = BCRYPT_ECDH_PRIVATE_P256_MAGIC
        lKeyPtr = VarPtr(baKey(0))
    Else
        Err.Raise vbObjectError, "pvBCryptToKeyBlob", "Unrecognized key size"
    End If
    ReDim baRetVal(0 To 8 + lSize - 1) As Byte
    Call CopyMemory(baRetVal(0), lMagic, 4)
    Call CopyMemory(baRetVal(4), BCRYPT_SECP256R1_PARTSZ, 4)
    Call CopyMemory(baRetVal(8), ByVal lKeyPtr, lSize)
    pvBCryptToKeyBlob = baRetVal
End Function

Private Function pvBCryptFromKeyBlob(baBlob() As Byte, Optional ByVal lSize As Long = -1) As Byte()
    Dim baRetVal()      As Byte
    Dim lMagic          As Long
    Dim lPartSize       As Long
    
    If lSize < 0 Then
        lSize = pvArraySize(baBlob)
    End If
    Call CopyMemory(lMagic, baBlob(0), 4)
    Select Case lMagic
    Case BCRYPT_ECDH_PUBLIC_P256_MAGIC
        Call CopyMemory(lPartSize, baBlob(4), 4)
        Debug.Assert lPartSize = 32
        ReDim baRetVal(0 To BCRYPT_SECP256R1_UNCOMPRESSED_PUBLIC_KEYSZ - 1) As Byte
        Debug.Assert lSize >= 8 + 2 * lPartSize
        baRetVal(0) = BCRYPT_SECP256R1_TAG_UNCOMPRESSED
        Call CopyMemory(baRetVal(1), baBlob(8), 2 * lPartSize)
    Case BCRYPT_ECDH_PRIVATE_P256_MAGIC
        Call CopyMemory(lPartSize, baBlob(4), 4)
        Debug.Assert lPartSize = 32
        ReDim baRetVal(0 To BCRYPT_SECP256R1_PRIVATE_KEYSZ - 1) As Byte
        Debug.Assert lSize >= 8 + 3 * lPartSize
        Call CopyMemory(baRetVal(0), baBlob(8), 3 * lPartSize)
    Case Else
        Err.Raise vbObjectError, "pvBCryptFromKeyBlob", "Unknown BCrypt magic"
    End Select
    pvBCryptFromKeyBlob = baRetVal
End Function
#End If

'= trampolines ===========================================================

Private Function pvCryptoCallSecp256r1MakeKey(ByVal Pfn As Long, pPubKeyPtr As Byte, pPrivKeyPtr As Byte) As Long
    ' int ecc_make_key(uint8_t p_publicKey[ECC_BYTES+1], uint8_t p_privateKey[ECC_BYTES]);
End Function

Private Function pvCryptoCallSecp256r1SharedSecret(ByVal Pfn As Long, pPubKeyPtr As Byte, pPrivKeyPtr As Byte, pSecretPtr As Byte) As Long
    ' int ecdh_shared_secret(const uint8_t p_publicKey[ECC_BYTES+1], const uint8_t p_privateKey[ECC_BYTES], uint8_t p_secret[ECC_BYTES]);
End Function

Private Function pvCryptoCallSecp256r1UncompressKey(ByVal Pfn As Long, pPubKeyPtr As Byte, pUncompressedKeyPtr As Byte) As Long
    ' int ecdh_uncompress_key(const uint8_t p_publicKey[ECC_BYTES + 1], uint8_t p_uncompressedKey[2 * ECC_BYTES + 1])
End Function

Private Function pvCryptoCallSecp256r1Sign(ByVal Pfn As Long, pPrivKeyPtr As Byte, pHashPtr As Byte, pRandomPtr As Byte, pSignaturePtr As Byte) As Long
    ' int ecdsa_sign(const uint8_t p_privateKey[ECC_BYTES], const uint8_t p_hash[ECC_BYTES], uint64_t k[NUM_ECC_DIGITS], uint8_t p_signature[ECC_BYTES*2])
End Function

Private Function pvCryptoCallSecp256r1Verify(ByVal Pfn As Long, pPubKeyPtr As Byte, pHashPtr As Byte, pSignaturePtr As Byte) As Long
    ' int ecdsa_verify(const uint8_t p_publicKey[ECC_BYTES+1], const uint8_t p_hash[ECC_BYTES], const uint8_t p_signature[ECC_BYTES*2])
End Function

Private Function pvCryptoCallCurve25519Multiply(ByVal Pfn As Long, pSecretPtr As Byte, pPubKeyPtr As Byte, pPrivKeyPtr As Byte) As Long
    ' void cf_curve25519_mul(uint8_t out[32], const uint8_t priv[32], const uint8_t pub[32])
End Function

Private Function pvCryptoCallCurve25519MulBase(ByVal Pfn As Long, pPubKeyPtr As Byte, pPrivKeyPtr As Byte) As Long
    ' void cf_curve25519_mul_base(uint8_t out[32], const uint8_t priv[32])
End Function

Private Function pvCryptoCallSha256Init(ByVal Pfn As Long, ByVal lCtxPtr As Long) As Long
    ' void cf_sha256_init(cf_sha256_context *ctx)
End Function

Private Function pvCryptoCallSha256Update(ByVal Pfn As Long, ByVal lCtxPtr As Long, ByVal lDataPtr As Long, ByVal lSize As Long) As Long
    ' void cf_sha256_update(cf_sha256_context *ctx, const void *data, size_t nbytes)
End Function

Private Function pvCryptoCallSha256Final(ByVal Pfn As Long, ByVal lCtxPtr As Long, pHashPtr As Byte) As Long
    ' void cf_sha256_digest_final(cf_sha256_context *ctx, uint8_t hash[LNG_SHA256_HASHSZ])
End Function

Private Function pvCryptoCallSha384Init(ByVal Pfn As Long, ByVal lCtxPtr As Long) As Long
    ' void cf_sha384_init(cf_sha384_context *ctx)
End Function

Private Function pvCryptoCallSha384Update(ByVal Pfn As Long, ByVal lCtxPtr As Long, ByVal lDataPtr As Long, ByVal lSize As Long) As Long
    ' void cf_sha384_update(cf_sha384_context *ctx, const void *data, size_t nbytes)
End Function

Private Function pvCryptoCallSha384Final(ByVal Pfn As Long, ByVal lCtxPtr As Long, pHashPtr As Byte) As Long
    ' void cf_sha384_digest_final(cf_sha384_context *ctx, uint8_t hash[LNG_SHA384_HASHSZ])
End Function

Private Function pvCryptoCallChacha20Poly1305Encrypt( _
            ByVal Pfn As Long, pKeyPtr As Byte, pNoncePtr As Byte, _
            ByVal lHeaderPtr As Long, ByVal lHeaderSize As Long, _
            pPlaintTextPtr As Byte, ByVal lPlaintTextSize As Long, _
            pCipherTextPtr As Byte, pTagPtr As Byte) As Long
    ' void cf_chacha20poly1305_encrypt(const uint8_t key[32], const uint8_t nonce[12], const uint8_t *header, size_t nheader,
    '                                  const uint8_t *plaintext, size_t nbytes, uint8_t *ciphertext, uint8_t tag[16])
End Function

Private Function pvCryptoCallChacha20Poly1305Decrypt( _
            ByVal Pfn As Long, pKeyPtr As Byte, pNoncePtr As Byte, _
            pHeaderPtr As Byte, ByVal lHeaderSize As Long, _
            pCipherTextPtr As Byte, ByVal lCipherTextSize As Long, _
            pTagPtr As Byte, pPlaintTextPtr As Byte) As Long
    ' int cf_chacha20poly1305_decrypt(const uint8_t key[32], const uint8_t nonce[12], const uint8_t *header, size_t nheader,
    '                                 const uint8_t *ciphertext, size_t nbytes, const uint8_t tag[16], uint8_t *plaintext)
End Function

Private Function pvCryptoCallAesGcmEncrypt( _
            ByVal Pfn As Long, pCipherTextPtr As Byte, pTagPtr As Byte, pPlaintTextPtr As Byte, ByVal lPlaintTextSize As Long, _
            ByVal lHeaderPtr As Long, ByVal lHeaderSize As Long, pNoncePtr As Byte, pKeyPtr As Byte, ByVal lKeySize As Long) As Long
    ' void cf_aesgcm_encrypt(uint8_t *c, uint8_t *mac, const uint8_t *m, const size_t mlen, const uint8_t *ad, const size_t adlen,
    '                        const uint8_t *npub, const uint8_t *k, size_t klen)
End Function

Private Function pvCryptoCallAesGcmDecrypt( _
            ByVal Pfn As Long, pPlaintTextPtr As Byte, pCipherTextPtr As Byte, ByVal lCipherTextSize As Long, pTagPtr As Byte, _
            pHeaderPtr As Byte, ByVal lHeaderSize As Long, pNoncePtr As Byte, pKeyPtr As Byte, ByVal lKeySize As Long) As Long
    ' void cf_aesgcm_decrypt(uint8_t *m, const uint8_t *c, const size_t clen, const uint8_t *mac, const uint8_t *ad, const size_t adlen,
    '                        const uint8_t *npub, const uint8_t *k, const size_t klen)
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
