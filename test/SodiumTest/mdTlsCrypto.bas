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
Private Const CRYPT_VERIFYCONTEXT                       As Long = &HF0000000
'--- for CryptDecodeObjectEx
Private Const X509_ASN_ENCODING                         As Long = 1
Private Const PKCS_7_ASN_ENCODING                       As Long = &H10000
Private Const X509_PUBLIC_KEY_INFO                      As Long = 8
Private Const PKCS_RSA_PRIVATE_KEY                      As Long = 43
Private Const PKCS_PRIVATE_KEY_INFO                     As Long = 44
Private Const CRYPT_DECODE_ALLOC_FLAG                   As Long = &H8000
'--- for CryptCreateHash
Private Const CALG_MD5                                  As Long = &H8003&
Private Const CALG_SHA1                                 As Long = &H8004&
Private Const CALG_SHA_256                              As Long = &H800C&
Private Const CALG_SHA_384                              As Long = &H800D&
Private Const CALG_SHA_512                              As Long = &H800E&
'--- for CryptSignHash
Private Const AT_KEYEXCHANGE                            As Long = 1
Private Const MAX_RSA_KEY                               As Long = 8192     '--- in bits
'--- for CryptVerifySignature
Private Const NTE_BAD_SIGNATURE                         As Long = &H80090006
Private Const NTE_PROV_TYPE_NOT_DEF                     As Long = &H80090017
'--- for BCryptSignHash
Private Const BCRYPT_PAD_PSS                            As Long = 8
'--- for BCryptVerifySignature
Private Const STATUS_INVALID_SIGNATURE                  As Long = &HC000A000
Private Const ERROR_INVALID_DATA                        As Long = &HC000000D
'--- for thunks
Private Const MEM_COMMIT                                As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE                    As Long = &H40

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
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
Private Declare Function CryptDecrypt Lib "advapi32" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, pbData As Any, pdwDataLen As Long) As Long
Private Declare Function CryptImportPublicKeyInfo Lib "crypt32" (ByVal hCryptProv As Long, ByVal dwCertEncodingType As Long, pInfo As Any, phKey As Long) As Long
Private Declare Function CryptDecodeObjectEx Lib "crypt32" (ByVal dwCertEncodingType As Long, ByVal lpszStructType As Long, pbEncoded As Any, ByVal cbEncoded As Long, ByVal dwFlags As Long, ByVal pDecodePara As Long, pvStructInfo As Any, pcbStructInfo As Long) As Long
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

Private Type CRYPT_DER_BLOB
    cbData              As Long
    pbData              As Long
End Type

Private Type BCRYPT_PSS_PADDING_INFO
    pszAlgId            As Long
    cbSalt              As Long
End Type

Private Type CERT_PUBLIC_KEY_INFO
    AlgObjIdPtr         As Long
    AlgParamSize        As Long
    AlgParamPtr         As Long
    PubKeySize          As Long
    PubKeyPtr           As Long
    PubKeyUnusedBits    As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_GLOB                  As String = "////////////////AAAAAAAAAAAAAAAAAQAAAP////9LYNInPjzOO/awU8ywBh1lvIaYdlW967Pnkzqq2DXGWpbCmNhFOaH0oDPrLYF9A3fyQKRj5ea8+EdCLOHy0Rdr9VG/N2hAtsvOXjFrVzPOKxaeD3xK6+eOm38a/uJC409RJWP8wsq584SeF6et+ua8//////////8AAAAA//////////8AAAAAAAAAAP/////+/////////////////////////////////////////+8q7NPtyIUqndEuio05VsZahxNQjwgUAxJBgf5unB0YGS3442sFjpjk5z7ipy8xs7cKdnI4XlQ6bClVv13yAlU4KlSC4EH3WZibp4tiOx1udK0g8x7HsY43BYu+IsqHql8O6pB8HUN6nYF+Hc6xYArAuPC1EzHa6XwUmii9HfT4KdySkr+Ynl1vLCaWSt4XNnMpxcxqGezseqewSLINGljfLTf0gU1jx////////////////////////////////5gvikKRRDdxz/vAtaXbtelbwlY58RHxWaSCP5LVXhyrmKoH2AFbgxK+hTEkw30MVXRdvnL+sd6Apwbcm3Txm8HBaZvkhke+78adwQ/MoQwkbyzpLaqEdErcqbBc2oj5dlJRPphtxjGoyCcDsMd/Wb/zC+DGR5Gn1VFjygZnKSkUhQq3JzghGy78bSxNEw04U1RzCmW7Cmp2LsnCgYUscpKh6L+iS2YaqHCLS8KjUWzHGeiS0SQGmdaFNQ70cKBqEBbBpBkIbDceTHdIJ7W8sDSzDBw5SqrYTk/KnFvzby5o7oKPdG9jpXgUeMiECALHjPr/vpDrbFCk96P5vvJ4ccYirijXmC+KQs1l7yORRDdxLztN" & _
                                                    "7M/7wLW824mBpdu16Ti1SPNbwlY5GdAFtvER8VmbTxmvpII/khiBbdrVXhyrQgIDo5iqB9i+b3BFAVuDEoyy5E6+hTEk4rT/1cN9DFVviXvydF2+crGWFjv+sd6ANRLHJacG3JuUJmnPdPGbwdJK8Z7BaZvk4yVPOIZHvu+11YyLxp3BD2WcrHfMoQwkdQIrWW8s6S2D5KZuqoR0StT7Qb3cqbBctVMRg9qI+Xar32buUlE+mBAytC1txjGoPyH7mMgnA7DkDu++x39Zv8KPqD3zC+DGJacKk0eRp9VvggPgUWPKBnBuDgpnKSkU/C/SRoUKtycmySZcOCEbLu0qxFr8bSxN37OVnRMNOFPeY6+LVHMKZaiydzy7Cmp25q7tRy7JwoE7NYIUhSxykmQD8Uyh6L+iATBCvEtmGqiRl/jQcItLwjC+VAajUWzHGFLv1hnoktEQqWVVJAaZ1iogcVeFNQ70uNG7MnCgahDI0NK4FsGkGVOrQVEIbDcemeuO30x3SCeoSJvhtbywNGNaycWzDBw5y4pB40qq2E5z42N3T8qcW6O4stbzby5o/LLvXe6Cj3RgLxdDb2OleHKr8KEUeMiE7DlkGggCx4woHmMj+v++kOm9gt7rbFCkFXnGsvej+b4rU3Lj8nhxxpxhJurOPifKB8LAIce4htEe6+DN1n3a6njRbu5/T331um8Xcqpn8AammMiixX1jCq4N+b4EmD8RG0ccEzULcRuEfQQj9XfbKJMkx0B7q8oyvL7JFQq+njxMDRCcxGcdQ7ZCPsu+1MVMKn5l/Jwpf1ns+tY6q2/LXxdYR0qMGURsZXhwYW5kIDE2LWJ5dGUgawBleHBhbmQgMzItYnl0ZSBrAAAABQAA" & _
                                                    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAPwAAABjfHd78mtvxTABZyv+16t2yoLJffpZR/Ct1KKvnKRywLf9kyY2P/fMNKXl8XHYMRUExyPDGJYFmgcSgOLrJ7J1CYMsGhtuWqBSO9azKeMvhFPRAO0g/LFbasu+OUpMWM/Q76r7Q00zhUX5An9QPJ+oUaNAj5KdOPW8ttohEP/z0s0ME+xfl0QXxKd+PWRdGXNggU/cIiqQiEbuuBTeXgvb4DI6CkkGJFzC06xikZXkeefIN22N1U6pbFb06mV6rgi6eCUuHKa0xujddB9LvYuKcD61ZkgD9g5hNVe5hsEdnuH4mBFp2Y6Umx6H6c5VKN+MoYkNv+ZCaEGZLQ+wVLsWjQECBAgQIECAGzZSCWrVMDalOL9Ao56B89f7fOM5gpsv/4c0jkNExN7py1R7lDKmwiM97kyVC0L6w04ILqFmKNkksnZboklti9Elcvj2ZIZomBbUpFzMXWW2kmxwSFD97bnaXhVGV6eNnYSQ2KsAjLzTCvfkWAW4s0UG0Cwej8o/DwLBr70DAROKazqREUFPZ9zql/LPzvC05nOWrHQi5601heL5N+gcdd9uR/EacR0pxYlvt2IOqhi+G/xWPkvG0nkgmtvA/njNWvQf3agziAfHMbESEFkngOxfYFF/qRm1Sg0t5Xqfk8mc76DgO02uKvWwyOu7PINTmWEXKwR+unfWJuFpFGNVIQx9AAAAAAA=" ' 1928, 24.4.2020 12:48:15
Private Const STR_THUNK1                As String = "IAIlAdAgAADwIwAAMDkAAOA/AACAQAAAEEIAAEBHAADAQAAAQEMAAHAtAADALQAAoCsAADAuAADALgAAAC4AAFAyAADgMgAA0C4AAPAfAACwHwAAABUAAIAUAADMzMzM6AAAAABYLWVAJAEFAEAkAYsAw8zMzMzMzMzMzMzMzMzoAAAAAFgthUAkAQUAQCQBw8zMzMzMzMzMzMzMzMzMzFWL7IPsaFOLXRBT6EBsAACFwA+FWwEAAFaLdQyNRchXVlDoiXwAAIt9CI1FyFBXjUWYUOi4egAAjUXIUFDobnwAAFNWVuimegAAU1PoX3wAAOhq////BaAAAABQU1dX6OxzAADoV////wWgAAAAUFNTU+jZcwAA6ET///8FoAAAAFBTV1PohnwAAFNXV+heegAA6Cn///8FoAAAAFBXV1Poq3MAAOgW////BaAAAABQU1dX6JhzAABqAFfogIgAAAvCdCXo9/7//wWgAAAAUFdX6GpmAABXi/DoUoAAAMHmHwl3LIt1DOsGV+hBgAAAV1PounsAAOjF/v//BaAAAABQjUWYUFNT6AR8AADor/7//wWgAAAAUI1FmFBTU+juewAA6Jn+//8FoAAAAFBTjUWYUFDo2HsAAI1FmFBXV+iteQAA6Hj+//8FoAAAAFCNRchQV1Dot3sAAFNX6JCAAABWU+iJgAAAjUXIUFbof4AAAF9eW4vlXcIMAMzMzMzMzFWL7IPsSFOLXRBT6PBqAACFwA+FKQEAAFaLdQyNRdhXVlDoOXsAAIt9CI1F2FBXjUW4UOjIegAAjUXYUFDoHnsAAFNWVui2egAAU1PoD3sAAOjq/f//UFNXV+ixcgAA6Nz9//9QU1NT6KNyAADozv3//1BTV1PoRXsAAFNXV+h9egAA" & _
                                                    "6Lj9//9QV1dT6H9yAADoqv3//1BTV1focXIAAGoAV+gZhwAAC8J0IOiQ/f//UFdX6EhnAABXi/DoUH8AAMHmHwl3HIt1DOsGV+g/fwAAV1PoiHoAAOhj/f//UI1FuFBTU+jXegAA6FL9//9QjUW4UFNT6MZ6AADoQf3//1BTjUW4UFDotXoAAI1FuFBXV+jqeQAA6CX9//9QjUXYUFdQ6Jl6AABTV+iifwAAVlPom38AAI1F2FBW6JF/AABfXluL5V3CDADMzMzMzMzMzFWL7FaLdQhW6HNpAACFwHQXjUYwUOhmaQAAhcB0CrgBAAAAXl3CBAAzwF5dwgQAzFWL7FaLdQhW6HNpAACFwHQXjUYgUOhmaQAAhcB0CrgBAAAAXl3CBAAzwF5dwgQAzFWL7IHs+AAAAFOLXQyNRZhWV1NQ6Kd+AACNQzBQiUX4jYU4////UOiUfgAA/3UUjYUI////UI2FaP///1CNhTj///9QjUWYUOjDBwAAi10QU+iqfAAAjXD+hfZ+YA8fAFZT6KmFAAALwnUHuAEAAADrAjPAjQRAweAEjY0I////A8iNlWj///8D0IlNFFH32IlV/I29OP///1ID+I1dmAPYV1PoiAQAAFdT/3UU/3X86JsCAACLXRBOhfZ/o2oAU+hLhQAAC8J1B7gBAAAA6wIzwI0EQMHgBI2dCP///wPYjY1o////UwPIjb04////USv4iU0QjXWYK/BXVugvBAAA6Ir7//8FoAAAAFCNhWj///9QjUWYUI1FyFDowHgAAFeNRchQUOiVdgAA/3UMjUXIUFDoiHYAAOhT+///BaAAAABQjUXIUFDoU3AAAP91+I1FyFBQ6GZ2AABWjUXIUFDoW3YAAFdW" & _
                                                    "U/91EOjgAQAAjUXIUI2FCP///1CNhWj///9Q6EkKAACLdQiNhWj///9QVugpfQAAjYUI////UI1GMFDoGX0AAF9eW4vlXcIQAFWL7IHsqAAAAFOLXQyNRbhWV1NQ6Fd9AACNQyBQiUX4jYV4////UOhEfQAA/3UUjYVY////UI1FmFCNhXj///9QjUW4UOiWBgAAi10QU+hNewAAg+gCiUUUhcB+Ww8fAFBT6PmDAAALwnUHuAEAAADrAjPAweAFjZ1Y////A9iNTZgDyI21eP///1P32IlN/FED8I19uAP4VlfocQQAAFZXU/91/Oj2AQAAi0UUi10QSIlFFIXAf6hqAFPooIMAAAvCdQWNSAHrAjPJweEFjZ1Y////A9mJTRBTjUWYA8GNvXj///9QK/mNdbgr8VdW6BwEAADo5/n//1CNRZhQjUW4UI1F2FDoVXcAAFeNRdhQUOiKdgAA/3UMjUXYUFDofXYAAOi4+f//UI1F2FBQ6B1xAAD/dfiNRdhQUOhgdgAAVo1F2FBQ6FV2AABXVo1FmANFEFNQ6EYBAACNRdhQjYVY////UI1FmFDoAgkAAIt1CI1FmFBW6PV7AACNhVj///9QjUYgUOjlewAAX15bi+VdwhAAzMzMzMzMzMzMzMzMVYvsg+wwU1ZX6DL5//+LXQgFoAAAAIt1EFBTVo1F0FDoa3YAAI1F0FBQ6AF2AACNRdBQU1PoNnQAAI1F0FBWVugrdAAA6Pb4//+LdQwFoAAAAIt9FFBWV1foMnYAAFeNRdBQ6Mh1AADo0/j//wWgAAAAUFONRdBQUOgSdgAA6L34//8FoAAAAFCLRRBQjUXQUFDo+XUAAOik+P//BaAAAABQi0UQU1BQ6ON1" & _
                                                    "AACLRRBQVlbouHMAAOiD+P//BaAAAABQjUXQUFOLXRBT6L91AABTV1fol3MAAOhi+P//BaAAAABQVldX6KR1AACNRdBQU+h6egAAX15bi+VdwhAAzFWL7IPsIFNWV+gy+P//i10Ii3UQUFNWjUXgUOigdQAAjUXgUFDoNnUAAI1F4FBTU+jLdAAAjUXgUFZW6MB0AADo+/f//4t1DIt9FFBWV1fobHUAAFeNReBQ6AJ1AADo3ff//1BTjUXgUFDoUXUAAOjM9///UItFEFCNReBQUOg9dQAA6Lj3//9Qi0UQU1BQ6Cx1AACLRRBQVlboYXQAAOic9///UI1F4FBTi10QU+gNdQAAU1dX6EV0AADogPf//1BWV1fo93QAAI1F4FBT6P15AABfXluL5V3CEADMzMzMVYvsgeyQAAAAU1ZX6E/3//+LXQgFoAAAAIt9EFBTV41FoFDoiHQAAI1FoFBQ6B50AACNRaBQU1PoU3IAAI1FoFBXV+hIcgAA6BP3//+LXQwFoAAAAIt1FFBTVo1FoFDojGsAAOj39v//BaAAAABQU1ZW6Dl0AADo5Pb//wWgAAAAUP91CI1F0FdQ6CF0AACNRdBQU1Po9nEAAOjB9v//BaAAAABQV/91CI1F0FDoPmsAAFZX6JdzAADoovb//wWgAAAAUI1F0FBXV+jhcwAA6Iz2//8FoAAAAFBXi30IjYVw////V1DoxXMAAI2FcP///1BWVuiXcQAA6GL2//8FoAAAAFBTVlbopHMAAI1FoFCNhXD///9Q6DRzAADoP/b//wWgAAAAUI1F0FCNhXD///9QUOh4cwAA6CP2//8FoAAAAFBXjYVw////UI1F0FDoXHMAAI1FoFCNRdBQUOgu" & _
                                                    "cQAA6Pn1//8FoAAAAFBTjUXQUFPoOHMAAI2FcP///1BX6At4AABfXluL5V3CEADMzFWL7IPsYFNWV+jC9f//i10Ii30QUFNXjUXAUOgwcwAAjUXAUFDoxnIAAI1FwFBTU+hbcgAAjUXAUFdX6FByAADoi/X//4tdDIt1FFBTVo1FwFDoSWoAAOh09f//UFNWVujrcgAA6Gb1//9Q/3UIjUXgV1Do2HIAAI1F4FBTU+gNcgAA6Ej1//9QV/91CI1F4FDoCmoAAFZX6FNyAADoLvX//1CNReBQV1foonIAAOgd9f//UFeLfQiNRaBXUOiOcgAAjUWgUFZW6MNxAADo/vT//1BTVlbodXIAAI1FwFCNRaBQ6AhyAADo4/T//1CNReBQjUWgUFDoVHIAAOjP9P//UFeNRaBQjUXgUOhAcgAAjUXAUI1F4FBQ6HJxAADorfT//1BTjUXgUFPoIXIAAI1FoFBX6Cd3AABfXluL5V3CEADMzMzMzMzMzMzMzMzMzFWL7IPsMFaLdQhXVv91EOicdgAAi30MV/91FOiQdgAAjUXQUOhXXwAAi0UYx0XQAQAAAMdF1AAAAACFwHQKUI1F0FDoaHYAAI1F0FBXVuhtAwAAjUXQUFdW6GL0//+NRdBQ/3UU/3UQ6FMDAABfXovlXcIUAMzMzMzMzMzMzMzMVYvsg+wgVot1CFdW/3UQ6Hx2AACLfQxX/3UU6HB2AACNReBQ6DdfAACLRRjHReABAAAAx0XkAAAAAIXAdApQjUXgUOhIdgAAjUXgUFdW6D0DAACNReBQV1boYvX//41F4FD/dRT/dRDoIwMAAF9ei+VdwhQAzMzMzMzMzMzMzMxTi0QkDItMJBD34YvYi0QkCPdk" & _
                                                    "JBQD2ItEJAj34QPTW8IQAMzMzMzMzMzMzMzMzMyA+UBzFYD5IHMGD6XC0+DDi9AzwIDhH9PiwzPAM9LDzID5QHMVgPkgcwYPrdDT6sOLwjPSgOEf0+jDM8Az0sPMVYvsi0UQU1aLdQiNSHhXi30MjVZ4O/F3BDvQcwuNT3g78XcwO9dyLCv4uxAAAAAr8IsUOAMQi0w4BBNIBI1ACIlUMPiJTDD8g+sBdeRfXltdwgwAi9eNSBCL3ivQK9gr/rgEAAAAjXYgjUkgDxBB0A8QTDfgZg/UyA8RTuAPEEwK4A8QQeBmD9TIDxFMC+CD6AF10l9eW13CDADMzMzMzFWL7ItVHIPsCItFIFaLdQhXi30MA9cTRRCJFolGBDtFEHcPcgQ713MJuAEAAAAzyesOD1fAZg8TRfiLTfyLRfgDRSRfE00oA0UUiUYIi8YTTRiJTgxei+VdwiQAzMzMzFWL7ItVDItNCIsCMQGLQgQxQQSLQggxQQiLQgwxQQxdwggAzMzMzMzMzMzMzMzMzFWL7IPsCItNCItVEFNWiwGNWQTB6gIz9olVEIld+I0EhQQAAACJRfxXhdJ0QotVDIt9EIPCAmZmDx+EAAAAAAAPtkr+jVIED7ZC+8HhCAvID7ZC/MHhCAvID7ZC/cHhCAvIiQyzRjv3ctaLRfyL17kBAAAAM/+JTQw78A+DjQAAAIvGK8KNBIOJRQgPH0QAAItcs/w7+nUIQTP/iU0M6wSF/3Ut6Dfx//8FeAUAAMHDCFBT6IhXAACL2Ogh8f//i00MD7aECHgGAADB4Bgz2Osdg/oGdh6D/wR1GegA8f//BXgFAABQU+hUVwAAi9iLRQiLVRCLCEczy4PABItd+IlFCIkMs0aL" & _
                                                    "TQw7dfxygl9eW4vlXcIMAMzMzMzMzMzMzFWL7IPsMI1F0P91EFDonm0AAI1F0FCLRQhQUOjQawAA/3UQjUXQUFDow2sAAI1F0FCLRQxQUOi1awAAi+VdwgwAzMzMzMzMzMzMzMzMzMzMVYvsg+wgjUXg/3UQUOh+bQAAjUXgUItFCFBQ6BBtAAD/dRCNReBQUOgDbQAAjUXgUItFDFBQ6PVsAACL5V3CDADMzMzMzMzMzMzMzMzMzMxVi+yD7CBTVot1CDPJV4lN7IEEzgAAAQCLBM6DVM4EAItczgQPrNgQwfsQiUXog/kPdRXHRfwBAAAAi9DHRfAAAAAAiV346yIPV8BmDxNF9ItF+IlF8ItF9GYPE0Xgi1XgiUX8i0XkiUX4g/kPjXkBagAbwPfYD6/HK1X8aiWNNMaLRfgbRfBQUugS/P//i03oA8ET04PoAYPaAAEGi0XsEVYEi3UID6TLEMHhECkMxovPiU3sGVzGBIP5EA+CT////19eW4vlXcIEAMzMzMzMVYvsg+wQi1UMVlcPtgoPtkIBweEIC8gPtkICweEIC8gPtkIDweEIC8gPtkIFiU3wD7ZKBMHhCAvID7ZCBsHhCAvID7ZCB8HhCAvID7ZCCYlN9A+2SgjB4QgLyA+2QgrB4QgLyA+2QgvB4QgLyA+2QgyJTfgPtkoNweAIC8gPtkIOweEIC8gPtkIPweEIC8iJTfyLTQiLOY1xBIvHweAEA/CNRfBWUOiV/P//g+4Qg8f/dC2NRfBQ6KQ8AACNRfBQ6Ds9AABWjUXwUOhx/P//jUXwUOhIPAAAg+4Qg+8BddONRfBQ6Hc8AACNRfBQ6A49AABWjUXwUOhE/P//i3UQi1Xwi8KLTfTB6BiI" & _
                                                    "BovCwegQiEYBi8LB6AiIRgKLwcHoGIhWA4hGBIvBwegQiEYFi8HB6AiIRgaITgeLTfiLwcHoGIhGCIvBwegQiEYJi8HB6AiIRgqITguLTfyLwcHoGIhGDIvBwegQiEYNi8HB6AiIRg5fiE4PXovlXcIMAMzMVYvsg+wQU1ZXi1UMi10ID7YKD7ZCAcHhCI1zBAvID7ZCAsHhCAvID7ZCA8HhCAvID7ZCBYlN8A+2SgTB4QgLyA+2QgbB4QgLyA+2QgfB4QgLyA+2QgmJTfQPtkoIweEIC8gPtkIKweEIC8gPtkILweEIC8gPtkIMiU34D7ZKDcHgCAvID7ZCDsHhCAvID7ZCD8HhCAvIjUXwVlCJTfzoHfv//78BAAAAg8YQOTt2LpCNRfBQ6AdTAACNRfBQ6J5RAACNRfBQ6LU8AABWjUXwUOjr+v//R4PGEDs7ctONRfBQ6NpSAACNRfBQ6HFRAABWjUXwUOjH+v//i3UQi1Xwi8KLTfTB6BiIBovCwegQiEYBi8LB6AiIRgKLwcHoGIhWA4hGBIvBwegQiEYFi8HB6AiIRgaITgeLTfiLwcHoGIhGCIvBwegQiEYJi8HB6AiIRgqITguLTfyLwcHoGIhGDIvBwegQiEYNi8HB6AiIRg5fiE4PXluL5V3CDADMzMzMVYvsVot1CGj0AAAAagBW6Iw7AACLRRCDxAyD+BB0NYP4GHQag/ggdTxQ/3UMxwYOAAAAVug3+v//Xl3CDABqGP91DMcGDAAAAFboIfr//15dwgwAahD/dQzHBgoAAABW6Av6//9eXcIMAMzMzMzMzFWL7IHsAAEAAFbo8ev//76QUiQBge4AQCQBA/Do3+v///91KLkQUSQBx0X0EAAA" & _
                                                    "AP91JIHpAEAkAYl1+APBiUX8jYUA////UOhD/////3UIjYUA////ahD/dRRqDP91IP91HP91GP91EP91DFCNRfRQ6DoPAABei+VdwiQAzMzMVYvsgewAAQAAVuhx6///vpBSJAGB7gBAJAED8Ohf6////3UouRBRJAHHRfQQAAAA/3UkgekAQCQBiXX4A8GJRfyNhQD///9Q6MP+//9qEP91DI2FAP////91CGoM/3Ug/3Uc/3UY/3UU/3UQUI1F9FDoehAAAF6L5V3CJADMzMxVi+xRU4tdGDPAiUX8hdt0cYtVEItNDFbHRRgBAAAAV4s5i/Ir9zveD0LzhcB1HQ+2RRRWUItFCAPHUOjwOQAAi00Mg8QMi0X8i1UQhf91CTvyD0RFGIlF/I0EPjvCdRf/dQj/dSD/VRyLTQyLVRDHAQAAAADrAgExi0X8K951oF9eW4vlXcIcAMzMzMzMzMxVi+xWi3Ugi8aD6AB0YIPoAQ+ErAAAAFOD6AFXjUUUdG2LfSiLXSRXU2oBUP91EP91DP91COi2AAAAi00YV1M4TRx0L41G/ot1EFBRVv91DP91COgY////V1NqAY1FHFBW/3UM/3UI6IQAAABfW15dwiQAjUb/i3UQUFFW/3UM/3UI6On+//9fW15dwiQA/3Uoi10Q/3Uki30Mi3UIagFQU1dW6EgAAAD/dSiNRRz/dSRqAVBTV1boNAAAAF9bXl3CJAD/dSiKRRz/dSQwRRSNRRRqAVD/dRD/dQz/dQjoDQAAAF5dwiQAzMzMzMzMzMxVi+z/dSCLRRxQUP91GP91FP91EP91DP91COgRAAAAXcIcAMzMzMzMzMzMzMzMzMxVi+yLTQyLRSRTi10UixFWi3UY" & _
                                                    "V4XSdFmF9nRVi0UQi/4rwjvGD0L4i8IDRQhXU1DoGzgAAItFDAPfK/eDxAwBOIt9EDk4i0UkdSn/dQhQhfZ1Df9VIItNDItFJIkx6xT/VRyLTQyLRSTHAQAAAADrA4t9EDv3chlTUDv3dQX/VSDrA/9VHItFJCv3A98793PnhfZ0LotFDIsIi8crwYv+O8YPQviLRQhXA8FTUOieNwAAi0UMA9+DxAwBOCv3i30QddVfXltdwiAAzMzMzMzMVYvsi00cg+wIV4t9GIXJdHZTi10MVoM7AHUR/3UI/3Uk/1Ugi0UQi00ciQOLA4vxi1UQK9A7wYlVGA9C8DPAiXX8hfZ0L4tdFCvfiV34ZpCLdfyNFDiKDBOLVRgDVQiLXfgyDAKNFDhAiAo7xnLhi10Mi00cKTMrzgF1FAP+iU0chcl1kV5bX4vlXcIgAMzMVYvs6Ojn//+5cF8kAYHpAEAkAQPBi00IUVD/dRSNQXT/dRD/dQxqQFCNQTRQ6D7///9dwhAAzMzMzMzMzMzMzFWL7IPsbItNFFNWVw+2WQMPtkECD7ZRB8HiCMHjCAvYD7ZBAcHjCAvYD7YBweMIC9gPtkEGC9CJXdjB4ggPtkEFC9APtkEEweIIC9APtkEKiVX0iVXUD7ZRC8HiCAvQD7ZBCcHiCAvQD7ZBCMHiCAvQD7ZBDolV8IlV0A+2UQ/B4ggL0A+2QQ3B4ggL0A+2QQyLTQjB4ggL0IlV+A+2QQKJVcwPtlEDweIIC9APtkEBweIIC9APtgHB4ggL0A+2QQaJVeyJVcgPtlEHweIIC9APtkEFweIIC9APtkEEweIIC9APtkEKiVXoiVXED7ZRC8HiCAvQweIID7ZBCQvQD7ZBCMHiCAvQ" & _
                                                    "D7ZBDolV5IlVwA+2UQ/B4ggL0A+2QQ3B4ggL0A+2QQyLTQzB4ggL0IlV4A+2QQKJVbwPtlEDweIIC9APtkEBweIIC9APtgHB4ggL0A+2QQaJVQiJVbgPtlEHweIIC9APtkEFweIIC9APtkEEweIIC9APtkEKiVUUiVW0D7ZRC8HiCAvQD7ZBCcHiCAvQD7ZBCMHiCAvQD7ZBDolVDIlVsA+2UQ/B4ggL0A+2QQ3B4ggL0A+2QQzB4ggL0IlV/IlVrItVEA+2SgMPtkICweEIC8gPtkIBweEIC8gPtgLB4QgLyIlN3IlNqA+2cgcPtkIGD7Z6Cw+2Sg7B5ggL8MHnCA+2QgXB5ggL8MdFmAoAAAAPtkIEweYIC/APtkIKC/iJdaQPtkIJwecIC/gPtkIIwecIC/gPtkIPweAIC8GJfaAPtkoNweAIC8EPtkoMi1XcweAIC8GLTeyJRZzrA4tdEAPZi00IM9OJXRDBwhADyolNCDNN7MHBDAPZM9OJXRCLXQjBwggD2olV3ItV9ANV6DPyiV0IM9nBxhCLTRQDzsHDB4lNFDNN6MHBDAPRM/KJVfSLVRTBxggD1ol17It18AN15DP+iVUUM9HBxxCLTQwDz8HCB4lNDDNN5MHBDAPxM/6JdfCLdQzBxwgD94l9lIt9+AN94DPHiXUMM/HBwBCLTfwDyMHGB4lN/DNN4MHBDAP5M8eJffiLffzBwAgD+Il9/DP5i00QA8rBxwczwYlNEItNDMHAEAPIiU0MM8qLVRDBwQwD0TPCiVUQi1UMwcAIA9CJVQwz0YtN9APOwcIHiU30iVXoi1XcM9GLTfzBwhADyolN/DPOi3X0wcEMA/Ez1ol19It1/MHCCAPyiXX8M/GL" & _
                                                    "TfADz8HGB4lN8Il15It17DPxi00IwcYQA86JTQgzz4t98MHBDAP5M/eJffCLfQjBxggD/ol9CDP5i034A8vBxweJfeCLfZQz+YlN+ItNFMHHEAPPiU0UM8uLXfjBwQwD2TP7iV34wccIAX0Ui10UM9mLy4ld7MHBB4NtmAGLXfiJTewPhUD+//8BRZwBXcyLTdgDTRABVaiLVRiJTdiLXdiLw4tN1ANN9IgaiU3Ui03QA03wwegIiEIBi8OJTdCLTewBTciLTcQDTejB6BCIQgLB6xiIWgOLXdSLw4haBMHoCIhCBYvDiU3Ei03AA03kwegQiEIGiU3Ai028A03gAXWkAX2gwesYiFoHi13Qi8OIWgiJTbyLTbgDTQjB6AiIQgmLw4lNuItNtANNFMHoEIhCCsHrGIhaC4tdzIvDiU20i02wA00MiFoMwegIiEINi8OJTbCLTawDTfzB6BCIQg7B6xiIWg+LXciLw4lNrIhaEMHoCIhCEYvDwegQiEISwesYiFoTi13Ei8OIWhTB6AiIQhWLw8HoEIhCFsHrGIhaF4tdwIvDiFoYwegIiEIZi8PB6BCIQhrB6xiIWhuLXbyLw4haHMHoCIhCHYvDwegQiEIewesYiFofi124i8OIWiDB6AiIQiGLw8HoEIhCIsHrGIhaI4tdtIvDiFokwegIiEIli8PB6BCIQibB6xiIWieLXbCLw4haKMHoCIhCKYvDwegQiEIqwesYiFori9mIWiyLw8HoCIhCLYvDwegQiEIuwesYiFovi12oi8OIWjDB6AiIQjGNSjyLw8HrGMHoEIhCMohaM4tdpIvDiFo0wegIiEI1i8PB6BCIQjbB6xiIWjeLXaCLw4haOMHoCIhCOYvD" & _
                                                    "wegQiEI6wesYiFo7i1Wci8LB6AiIEYhBAYvCX8HoEMHqGF6IQQKIUQNbi+VdwhQAzFWL7Fb/dRCLdQj/dQxW6G0/AABqEP91FI1GIFDoLzAAAItFGIPEDMdGdAAAAACJRnheXcIUAMzMzMzMzMzMzMxVi+xWi3UIV/91DP92MI1+IFeNRhBQVuhE+f//i1Z4M8CABwF1C0A7wnQGgAQ4AXT1X15dwggAzMzMzMzMzMzMVYvsg+wQjUXwahD/dSBQ6LwvAACDxAyNRfBQagD/dST/dRz/dRj/dRT/dRD/dQz/dQjoGToAAIvlXcIgAMzMzFWL7P91JGoB/3Ug/3Uc/3UY/3UU/3UQ/3UM/3UI6O45AABdwiAAzMzMzMzMzMzMzFWL7OhY4P//uSBzJAGB6QBAJAEDwYtNCFFQ/3UUiwH/dRD/dQz/MI1BKFCNQRhQ6Kz3//9dwhAAzMzMzMzMzMxVi+yLTQiLRQyJQSyLRRCJQTBdwgwAzMzMzMzMzMzMzFWL7FaLdQhqNGoAVugfLwAAi00Mx0YsAAAAAIsBiUYwi0UQiUYEjUYIiQ7HRigAAAAA/zH/dRRQ6MMuAACDxBheXcIQAMzMzMzMzMzMzMzMVYvsgewgBAAAU1ZXanCNhXD9///HhWD9//9B2wAAagBQx4Vk/f//AAAAAMeFaP3//wEAAADHhWz9//8AAAAA6JwuAACLdQyNhWD///9qH1ZQ6FouAACKRh+DxBiApWD////4JD8MQIiFf////42F4Pv///91EFDoxEUAAA9XwI21YP7//2YPE4Vg/v//jb1o/v//uR4AAABmDxNFgPOluR4AAABmDxOF4P7//411gMeFYP7//wEAAACNfYjHhWT+//8A" & _
                                                    "AAAA86W5HgAAAMdFgAEAAACNteD+///HRYQAAAAAjb3o/v//u/4AAADzpbkgAAAAjbXg+///jb3g/f//86WLww+2y8H4A4PhBw+2tAVg////jYXg/f//0+6D5gFWUI1FgFDoljsAAFaNhWD+//9QjYXg/v//UOiCOwAAjYXg/v//UI1FgFCNheD8//9Q6Cvr//+NheD+//9QjUWAUFDoekMAAI2FYP7//1CNheD9//9QjYXg/v//UOgA6///jYVg/v//UI2F4P3//1BQ6ExDAACNheD8//9QjYVg/v//UOgZQwAAjUWAUI2FYPz//1DoCUMAAI1FgFCNheD+//9QjUWAUOgFLwAAjYXg/P//UI2F4P3//1CNheD+//9Q6OsuAACNheD+//9QjUWAUI2F4Pz//1DohOr//42F4P7//1CNRYBQUOjTQgAAjUWAUI2F4P3//1Doo0IAAI2FYPz//1CNhWD+//9QjYXg/v//UOipQgAAjYVg/f//UI2F4P7//1CNRYBQ6IIuAACNhWD+//9QjUWAUFDoIer//41FgFCNheD+//9QUOhgLgAAjYVg/P//UI2FYP7//1CNRYBQ6EkuAACNheD7//9QjYXg/f//UI2FYP7//1DoLy4AAI2F4Pz//1CNheD9//9Q6AxCAABWjYXg/f//UI1FgFDo+zkAAFaNhWD+//9QjYXg/v//UOjnOQAAg+sBD4kf/v//jYXg/v//UFDowSkAAI2F4P7//1CNRYBQUOjQLQAAjUWAUP91COjkMAAAX15bi+VdwgwAzMzMzMzMzMzMzMxVi+yD7CCNReDGReAJUP91DA9XwMdF+QAAAAD/dQgPEUXhZsdF/QAAZg/WRfHGRf8A6Kr8//+L"
Private Const STR_THUNK2                As String = "5V3CCADMzMzMVYvsgewUAQAAU4tdCI1F8FZXi30MD1fAUFCLQwRXxkXwAGYP1kXxx0X5AAAAAGbHRf0AAMZF/wD/0It1JIP+DHUgVv91II1F0FDoASsAAIPEDGbHRd0AAMZF3ADGRd8B6zCNRfBQjYXs/v//UOiuKAAAVv91II2F7P7//1DoziYAAI1F0FCNhez+//9Q6H4nAACNRfBQjYU8////UOh+KAAA/3UcjYU8/////3UYUOh8JgAAjUXQxkXgAFBXU41FjMdF6QAAAAAPV8Bmx0XtAABQZg/WReHGRe8A6HD7//9qBGoMjUWMUOhD+///ahCNReBQUI1FjFDo8/r///91FI2FPP////91EFDoQSYAAI1FwFCNhTz///9Q6PEmAACLdSyNReBWUI1FwFBQ6L9kAAAy0o1FwLsBAAAAhfZ0Got9KIvIK/mKDAeNQAEySP8K0SvzdfGE0nUU/3UUjUWM/3Uw/3UQUOiF+v//M9sPV8APEUXwikXwDxFF0IpF0A8RReCKReAPEUXAikXAalCNhTz///9qAFDo5CkAAIqNPP///41FjGo0agBQ6NEpAACKTYyDxBiLw19eW4vlXcIsAFWL7IHsFAEAAFOLXQiNRfBWV4t9DA9XwFBQi0MEV8ZF8ABmD9ZF8cdF+QAAAABmx0X9AADGRf8A/9CLdSSD/gx1IFb/dSCNRdBQ6EEpAACDxAxmx0XdAADGRdwAxkXfAeswjUXwUI2F7P7//1Do7iYAAFb/dSCNhez+//9Q6A4lAACNRdBQjYXs/v//UOi+JQAAjUXwUI2FPP///1DoviYAAP91HI2FPP////91GFDovCQAAI1F0MZF4ABQV1ONRYzHRekAAAAAD1fAZsdF7QAAUGYP1kXhxkXv" & _
                                                    "AOiw+f//agRqDI1FjFDog/n//2oQjUXgUFCNRYxQ6DP5//+LfRSNRYyLdShXVv91EFDoH/n//1dWjYU8////UOhxJAAAjUXAxkXAAFCNhTz////HRckAAAAAD1fAZsdFzQAAUGYP1kXBxkXPAOgEJQAA/3UwjUXgUI1FwFD/dSzo0WIAAA9XwA8RRfCKRfAPEUXQikXQDxFF4IpF4A8RRcCKRcBqUI2FPP///2oAUOgyKAAAioU8////ajSNRYxqAFDoHygAAIpFjIPEGF9eW4vlXcIsAFWL7ItVDItNEFaLdQiLBjMCiQGLRgQzQgSJQQSLRggzQgiJQQiLRgwzQgyJQQxeXcIMAMzMzMzMzMzMzMzMzMxVi+xRU4tdDFZXi30IZsdF/ADhiw+LwdHog+EBiQOLVwSLwtHog+IBweEfC8jB4h+JSwSLdwiLxtHog+YBC9DB5h+JUwiLTwyLwdHog+EBC/BfiXMMD7ZEDfzB4BgxA15bi+VdwggAzMzMzMzMzMzMVYvsi1UMVot1CA+2Dg+2RgHB4QgLyA+2RgLB4QgLyA+2RgPB4QgLyIkKD7ZOBA+2RgXB4QgLyA+2RgbB4QgLyA+2RgfB4QgLyIlKBA+2TggPtkYJweEIC8gPtkYKweEIC8gPtkYLweEIC8iJSggPtk4MD7ZGDcHhCAvID7ZGDsHhCAvID7ZGD8HhCAvIiUoMXl3CCADMzMzMzMzMzMzMzFWL7IPsIFZXahCNReBqAFDoqyYAAGoQ/3UMjUXwUOhtJgAAi30Ig8QYDxBN4DP2kIvGuR8AAACD4B8ryIvGwfgFiwSH0+ioAXQMDxBF8GYP78gPEU3gjUXwUFDokP7//0aB/oAAAAB8x2oQjUXg" & _
                                                    "UP91EOgZJgAAg8QMX16L5V3CDADMzMzMzMzMzMzMzMzMzFWL7FaLdQxXi30IixeLwsHoGIgGi8LB6BCIRgGLwsHoCIhGAohWA4tPBIvBwegYiEYEi8HB6BCIRgWLwcHoCIhGBohOB4tPCIvBwegYiEYIi8HB6BCIRgmLwcHoCIhGCohOC4tPDIvBwegYiEYMi8HB6BCIRg2LwcHoCIhGDl+ITg9eXcIIAMzMzMzMzMzMzFWL7IPsRFaLdQiDvqgAAAAAdAZW6IctAAAzyQ8fRAAAD7aEDogAAACJRI28QYP5EHLuVsdF/AAAAADoYSwAAI1FvFBW6PcrAACLVQwzyWaQigSOiAQRQYP5EHL0aKwAAABqAFboNyUAAIoGg8QMXovlXcIIAMzMzMzMzMzMzMzMVYvsVot1CGisAAAAagBW6AwlAACLTQxqEP91EA+2AYlGRA+2QQGJRkgPtkECiUZMD7ZBA4PgD4lGUA+2QQQl/AAAAIlGVA+2QQWJRlgPtkEGiUZcD7ZBB4PgD4lGYA+2QQgl/AAAAIlGZA+2QQmJRmgPtkEKiUZsD7ZBC4PgD4lGcA+2QQwl/AAAAIlGdA+2QQ2JRngPtkEOiUZ8D7ZBD4PgD8eGhAAAAAAAAACJhoAAAACNhogAAABQ6DEkAACDxBheXcIMAMzMzMzMzMzMzFWL7OgY1f//ucCZJAGB6QBAJAEDwYtNCFFQ/3UQjYGoAAAA/3UMahBQjYGYAAAAUOhr6///XcIMAMzMzMzMzMxVi+yD7BhTVlfo0tT///91CL4wnyQBuUAAAACB7gBAJAED8ItFCFaNeGSLQGD34QMHi9iD0gCDwAiD4D8ryFFqAGoAaIAAAABqQFeLfQgPpNoD" & _
                                                    "iVX8jUcgweMDUIlV+OgM6v//i1X8i8uLwohd78HoGIhF6IvCwegQiEXpi8LB6AiIReqKRfiIReuLwg+swRhqCMHoGIhN7IvCi8sPrMEQwegQi8OITe0PrNAIiEXujUXoUMHqCFfoZAEAAIsXi8KLdQzB6BiIBovCwegQiEYBi8LB6AiIRgKIVgOLTwSLwcHoGIhGBIvBwegQiEYFi8HB6AiIRgaITgeLTwiLwcHoGIhGCIvBwegQiEYJi8HB6AiIRgqITguLTwyLwcHoGIhGDIvBwegQiEYNi8HB6AiIRg6ITg+LTxCLwcHoGIhGEIvBwegQiEYRi8HB6AiIRhKIThOLTxSLwcHoGIhGFIvBwegQiEYVi8HB6AiIRhaITheLTxiLwcHoGIhGGIvBwegQiEYZi8HB6AiIRhqIThuLTxyLwcHoGIhGHIvBwegQiEYdi8FqaMHoCGoAiEYeV4hOH+hZIgAAg8QMX15bi+VdwggAzMzMzMzMzMzMzMzMzFWL7FaLdQhqaGoAVugvIgAAg8QMxwZn5glqx0YEha5nu8dGCHLzbjzHRgw69U+lx0YQf1IOUcdGFIxoBZvHRhir2YMfx0YcGc3gW15dwgQAVYvs6LjS//+5MJ8kAYHpAEAkAQPBi00IUVD/dRCNQWT/dQxqQFCNQSBQ6BHp//9dwgwAzMzMzMzMzMzMzMzMzFWL7IPsQI1FwFD/dQjovgAAAGowjUXAUP91DOhgIQAAg8QMi+VdwggAzMzMzMzMzFWL7FaLdQhoyAAAAGoAVuhsIQAAg8QMxwbYngXBx0YEXZ27y8dGCAfVfDbHRgwqKZpix0YQF91wMMdGFFoBWZHHRhg5WQ73x0Yc2OwvFcdGIDELwP/H" & _
                                                    "RiRnJjNnx0YoERVYaMdGLIdKtI7HRjCnj/lkx0Y0DS4M28dGOKRP+r7HRjwdSLVHXl3CBADMzMzMzOkbBAAAzMzMzMzMzMzMzMxVi+yD7ByLRQhTjZjEAAAAVouAwAAAAFe/gAAAAPfni/ADM4vGg9IAD6TCA8HgA4lV/IlF+IlV9Ohz0f///3UIuQChJAGB6QBAJAEDwVCNRhCLdQiD4H8r+FdqAGoAaIAAAABogAAAAFONRkBQ6M7m//9qCI1F5MdF5AAAAABQVsdF6AAAAADohAMAAItd/IvDi1X4i8rB6BiIReSLw8HoEIhF5YvDwegIiEXmikX0iEXni8MPrMEYagjB6BiITeiLw4vKiFXrD6zBEMHoEIvCiE3pD6zYCIhF6o1F5FBWwesI6CkDAACLXgSLw4sOiU38wegYi30MiAeLw8HoEIhHAYvDwegIiEcCi8MPrMEYiF8DwegYiE8Ei8OLTfwPrMEQwegQiE8Fi038i8EPrNgIiEcGi8aITwfB6wiLWAiLy4tQDIvCwegYiEcIi8LB6BCIRwmLwsHoCIhHCovCD6zBGIhXC8HoGIhPDIvCi8sPrMEQwegQiE8Ni8MPrNAIiEcOi8aIXw/B6giLWBCLy4tQFIvCwegYiEcQi8LB6BCIRxGLwsHoCIhHEovCD6zBGIhXE8HoGIhPFIvCi8sPrMEQwegQi8OITxUPrNAIiEcWi8bB6giIXxeLWBiLy4tQHIvCwegYiEcYi8LB6BCIRxmLwsHoCIhHGovCD6zBGIhXG8HoGIhPHIvCi8sPrMEQwegQiE8di8MPrNAIiEcei8aIXx/B6giLWCCLy4tQJIvCwegYiEcgi8LB6BCIRyGLwsHoCIhHIovCD6zB" & _
                                                    "GIhXI8HoGIhPJIvCi8sPrMEQwegQiE8li8MPrNAIiEcmi8aIXyfB6giLWCiLy4tQLIvCwegYiEcoi8LB6BCIRymLwsHoCIhHKovCD6zBGIhXK8HoGIhPLIvCi8sPrMEQwegQi8OITy0PrNAIweoIiEcui8aIXy+NdzhoyAAAAGoAi1gwi8uLUDSLwsHoGIhHMIvCwegQiEcxi8LB6AiIRzKLwg+swRiIVzPB6BiITzSLwovLD6zBEMHoEIhPNYvDD6zQCIhHNohfN4t9CMHqCFeLVzyLwotfOIvLwegYiAaLwsHoEIhGAYvCwegIiEYCi8IPrMEYiFYDwegYiE4Ei8KLyw+swRDB6BCLw4hOBQ+s0AiIRgbB6giIXgfodR0AAIPEDF9eW4vlXcIIAMzMzMzMzMzMzFWL7FaLdQhoyAAAAGoAVuhMHQAAg8QMxwYIybzzx0YEZ+YJasdGCDunyoTHRgyFrme7x0YQK/iU/sdGFHLzbjzHRhjxNh1fx0YcOvVPpcdGINGC5q3HRiR/Ug5Rx0YoH2w+K8dGLIxoBZvHRjBrvUH7x0Y0q9mDH8dGOHkhfhPHRjwZzeBbXl3CBADMzMzMzFWL7OiYzf//uQChJAGB6QBAJAEDwYtNCFFQ/3UQjYHEAAAA/3UMaIAAAABQjUFAUOjr4///XcIMAMzMzMzMzMxVi+xWi3UI/3UMiw6NRghQ/3YEi0EE/9CLViyLRjAD1khegEQCCAF1Ew8fgAAAAACFwHQISIBEAggBdPRdwggAVYvsU4tdDFZXi30ID7ZDKJmLyIvyD6TOCA+2QynB4QiZC8gL8g+kzggPtkMqweEImQvIC/IPpM4ID7ZDK8HhCJkLyAvyD7ZDLA+kzgiZ" & _
                                                    "weEIC/ILyA+2Qy0PpM4ImcHhCAvyC8gPtkMuD6TOCJnB4QgL8gvID7ZDLw+kzgiZweEIC/ILyIl3BIkPD7ZDIJmLyIvyD7ZDIQ+kzgiZweEIC/ILyA+2QyIPpM4ImcHhCAvyC8gPtkMjD6TOCJnB4QgL8gvID7ZDJA+kzgiZweEIC8gL8g+kzggPtkMlweEImQvIC/IPpM4ID7ZDJsHhCJkLyAvyD6TOCA+2QyfB4QiZC8gL8olPCIl3DA+2QxiZi8iL8g+kzggPtkMZweEImQvIC/IPtkMaD6TOCJnB4QgL8gvID7ZDGw+kzgiZweEIC/ILyA+2QxwPpM4ImcHhCAvyC8gPtkMdD6TOCJnB4QgL8gvID7ZDHg+kzgiZweEIC/ILyA+2Qx8PpM4ImcHhCAvyC8iJdxSJTxAPtkMQmYvIi/IPtkMRD6TOCJnB4QgL8gvID7ZDEg+kzgjB4QiZC8gL8g+kzggPtkMTweEImQvIC/IPpM4ID7ZDFMHhCJkLyAvyD6TOCA+2QxXB4QiZC8gL8g+kzggPtkMWweEImQvIC/IPtkMXD6TOCJnB4QgL8gvIiXcciU8YD7ZDCJmLyIvyD7ZDCQ+kzgiZweEIC/ILyA+2QwoPpM4ImcHhCAvyC8gPtkMLD6TOCJnB4QgL8gvID7ZDDA+kzgiZweEIC/ILyA+2Qw0PpM4ImcHhCAvyC8gPtkMOD6TOCJnB4QgL8gvID7ZDDw+kzgiZweEIC8gL8olPIIl3JA+2A5mLyIvyD7ZDAQ+kzgiZweEIC/ILyA+2QwIPpM4ImcHhCAvyC8gPtkMDD6TOCJnB4QgL8gvID7ZDBA+kzgiZweEIC/ILyA+2QwUPpM4ImcHhCAvyC8gPtkMG" & _
                                                    "D6TOCJnB4QgL8gvID7ZDBw+kzgiZweEIC8gL8ol3LIlPKF9eW13CCADMzMzMzFWL7FOLXQxWV4t9CA+2QxiZi8iL8g+kzggPtkMZweEImQvIC/IPpM4ID7ZDGsHhCJkLyAvyD6TOCA+2QxvB4QiZC8gL8g+2QxwPpM4ImcHhCAvyC8gPtkMdD6TOCJnB4QgL8gvID7ZDHg+kzgiZweEIC/ILyA+2Qx8PpM4ImcHhCAvyC8iJdwSJDw+2QxCZi8iL8g+2QxEPpM4ImcHhCAvyC8gPtkMSD6TOCJnB4QgL8gvID7ZDEw+kzgiZweEIC/ILyA+2QxQPpM4ImcHhCAvIC/IPpM4ID7ZDFcHhCJkLyAvyD6TOCA+2QxbB4QiZC8gL8g+kzggPtkMXweEImQvIC/KJTwiJdwwPtkMImYvIi/IPpM4ID7ZDCcHhCJkLyAvyD7ZDCg+kzgiZweEIC/ILyA+2QwsPpM4ImcHhCAvyC8gPtkMMD6TOCJnB4QgL8gvID7ZDDQ+kzgiZweEIC/ILyA+2Qw4PpM4ImcHhCAvyC8gPtkMPD6TOCJnB4QgL8gvIiXcUiU8QD7YDmYvIi/IPtkMBD6TOCJnB4QgL8gvID7ZDAg+kzgjB4QiZC8gL8g+2QwMPpM4ImcHhCAvyC8gPtkMED6TOCJnB4QgL8gvID7ZDBQ+kzgiZweEIC/ILyA+2QwYPpM4ImcHhCAvyC8gPtkMHD6TOCJnB4QgLyAvyiXcciU8YX15bXcIIAMzMzFWL7IHskAAAAI1F0P91DFDoy/r//41F0FDoUjQAAIXAdAgzwIvlXcIIAI1F0FDorcf//wVgAQAAUOhSMwAAg/gBdBXomMf//wVgAQAAUI1F0FBQ6JhN" & _
                                                    "AABqAI1F0FDofcf//wUAAQAAUI2FcP///1Do28r//42FcP///1Dob8r//4XAdZ2KRaCLTQgkAQQCiAGNhXD///9QjUEBUOivAAAAuAEAAACL5V3CCADMzMzMVYvsg+xgjUXg/3UMUOgu/f//jUXgUOjVMwAAhcB0CDPAi+VdwggAjUXgUOgAx///g+iAUOgXMwAAg/gBdBPo7cb//4PogFCNReBQUOj/TgAAagCNReBQ6NTG//+DwEBQjUWgUOjny///jUWgUOj+yf//hcB1qYpFwItNCCQBBAKIAY1FoFCNQQFQ6IECAAC4AQAAAIvlXcIIAMzMzMzMzFWL7FaLdQixKFeLfQwPtkcHiEYoD7ZHBohGKYsHi1cE6DvT//+IRiqxIIsHi1cE6CzT//+IRiuLD4tHBA+swRiITiyLD8HoGItHBA+swRCITi2LD8HoEItHBA+swQiITi6xKMHoCA+2B4hGLw+2Rw+IRiAPtkcOiEYhi0cIi1cM6NvS//+IRiKxIItHCItXDOjL0v//iEYji08Ii0cMD6zBGIhOJItPCMHoGItHDA+swRCITiWLTwjB6BCLRwwPrMEIiE4msSjB6AgPtkcIiEYnD7ZHF4hGGA+2RxaIRhmLRxCLVxTodtL//4hGGrEgi0cQi1cU6GbS//+IRhuLTxCLRxQPrMEYiE4ci08QwegYi0cUD6zBEIhOHYtPEMHoEItHFA+swQiITh6xKMHoCA+2RxCIRh8PtkcfiEYQD7ZHHohGEYtHGItXHOgR0v//iEYSsSCLRxiLVxzoAdL//4hGE4tPGItHHA+swRiIThSLTxjB6BiLRxwPrMEQiE4Vi08YwegQi0ccD6zBCIhOFrEowegID7ZHGIhG" & _
                                                    "Fw+2RyeIRggPtkcmiEYJi0cgi1ck6KzR//+IRgqxIItHIItXJOic0f//iEYLi08gi0ckD6zBGMHoGIhODItPIItHJA+swRDB6BCITg2LTyCLRyQPrMEIwegIiE4OD7ZHIIhGDw+2Ry+IBg+2Ry6IRgGxKItHKItXLOhI0f//iEYCsSCLRyiLVyzoONH//4hGA4tPKItHLA+swRjB6BiITgSLTyiLRywPrMEQwegQiE4Fi08oi0csD6zBCMHoCIhOBg+2RyhfiEYHXl3CCADMzMzMzMzMzFWL7FaLdQixKFeLfQwPtkcHiEYYD7ZHBohGGYsHi1cE6MvQ//+IRhqxIIsHi1cE6LzQ//+IRhuLD4tHBA+swRiIThyLD8HoGItHBA+swRCITh2LD8HoEItHBA+swQiITh6xKMHoCA+2B4hGHw+2Rw+IRhAPtkcOiEYRi0cIi1cM6GvQ//+IRhKxIItHCItXDOhb0P//iEYTi08Ii0cMD6zBGIhOFItPCMHoGItHDA+swRCIThWLTwjB6BCLRwwPrMEIiE4WsSjB6AgPtkcIiEYXD7ZHF4hGCA+2RxaIRgmLRxCLVxToBtD//4hGCrEgi0cQi1cU6PbP//+IRguLTxCLRxQPrMEYiE4Mi08QwegYi0cUD6zBEIhODYtPEMHoEItHFA+swQiITg6xKMHoCA+2RxCIRg8PtkcfiAYPtkceiEYBi0cYi1cc6KLP//+IRgKxIItHGItXHOiSz///iEYDi08Yi0ccD6zBGMHoGIhOBItPGItHHA+swRDB6BCITgWLTxiLRxwPrMEIwegIiE4GD7ZHGF+IRgdeXcIIAMzMVYvsg+wwU4tdCA9XwFaLdQzHRdADAAAAx0XUAAAA" & _
                                                    "AA8RRdiNRgFmD9ZF+FBTDxFF6OhK9f//gD4EdRWNRjFQjUMwUOg49f//XluL5V3CCABXU417MFfoFT8AAOggwv//BaAAAABQjUXQUFdX6F8/AABTV1foNz0AAOgCwv//BaAAAABQ6PfB//8F0AAAAFBXV+h6NgAAV+jUEQAAigYz9osPJAEPtsCD4QGZO8h1BDvydBJX6MfB//8FoAAAAFBX6MtHAABfXluL5V3CCADMzFWL7IPsIFOLXQgPV8BXi30Mx0XgAwAAAMdF5AAAAAAPEUXojUcBZg/WRfhQU+iO9///gD8EdRWNRyFQjUMgUOh89///X1uL5V3CCABWU41zIFboeT4AAOhUwf//UI1F4FBWVujIPgAAU1ZW6AA+AADoO8H//1DoNcH//4PAIFBWVuj6NQAAVujEEQAAigcz/4sOJAEPtsCD4QGZO8h1BDv6dA1W6AfB//9QVuggSQAAXl9bi+VdwggAzMzMzMzMzFWL7IHs8AAAAI2FEP////91CFDoWP7///91DI1F0FDozPP//2oAjUXQUI2FEP///1CNhXD///9Q6CPE//+NhXD///9Q/3UQ6BT6//+NhXD///9Q6KjD///32BvAQIvlXcIMAMzMzMzMzMzMzMzMzMxVi+yB7KAAAACNhWD/////dQhQ6Lj+////dQyNReBQ6Gz2//9qAI1F4FCNhWD///9QjUWgUOhmxf//jUWgUP91EOga/P//jUWgUOhxw///99gbwECL5V3CDADMzMzMzMxVi+yD7GCNRaBW/3UIUOiN/f//i3UMjUWgUI1GAcYGBFDoavn//41F0FCNRjFQ6F35//+4AQAAAF6L5V3CCADMVYvsg+xAjUXAVv91CFDoHf7/" & _
                                                    "/4t1DI1FwFCNRgHGBgRQ6Jr7//+NReBQjUYhUOiN+///uAEAAABei+VdwggAzFWL7IHswAAAAFeLfRBX6B0sAACFwHQJM8Bfi+VdwhAAV+h6v///BWABAABQ6B8rAACD+AF0Euhlv///BWABAABQV1foaEUAAGoAV+hQv///BQABAABQjYVA////UOiuwv//jYVA////UOgyv///BWABAABQ6NcqAACD+AF0GOgdv///BWABAABQjYVA////UFDoGkUAAI2FQP///1DojisAAIXAD4Vt////Vot1FI2FQP///1BW6FX4////dQiNRaBQ6Nnx///o1L7//wVgAQAAUI1FoFCNhUD///9QjUXQUOhqOAAA/3UMjUWgUOiu8f//6Km+//8FYAEAAFCNRdBQjUWgUI1F0FDoIjMAAOiNvv//BWABAABQV1fokDMAAOh7vv//BWABAABQV41F0FBQ6Bo4AACNRdBQjUYwUOjN9///XrgBAAAAX4vlXcIQAFWL7IHsgAAAAFeLfRBX6P0qAACFwHQJM8Bfi+VdwhAAV+gqvv//g+iAUOhBKgAAg/gBdBDoF77//4PogFBXV+gsRgAAagBX6AS+//+DwEBQjUWAUOgXw///jUWAUOjuvf//g+iAUOgFKgAAg/gBdBPo273//4PogFCNRYBQUOjtRQAAjUWAUOiEKgAAhcB1h1aLdRSNRYBQVuiS+f///3UIjUXAUOi28///6KG9//+D6IBQjUXAUI1FgFCNReBQ6Ow4AAD/dQyNRcBQ6JDz///oe73//4PogFCNReBQjUXAUI1F4FDoNjIAAOhhvf//g+iAUFdX6MY0AADoUb3//4PogFBXjUXgUFDoojgAAI1F4FCNRiBQ" & _
                                                    "6BX5//9euAEAAABfi+VdwhAAzMzMzMzMzMxVi+yB7IACAACNhYD9//9W/3UIUOiH+v//i3UQjYXQ/v//VlDo9+///41GMFCNhWD///9Q6Ofv//+NhdD+//9Q6GspAACFwA+FnAMAAI2FYP///1DoVykAAIXAD4WIAwAAjYXQ/v//UOizvP//BWABAABQ6FgoAACD+AEPhWgDAACNhWD///9Q6JO8//8FYAEAAFDoOCgAAIP4AQ+FSAMAAFNX6Hi8//8FYAEAAFCNhWD///9QjUXAUOhyMQAA/3UMjYUA////UOhT7///6E68//8FYAEAAFCNRcBQjYUA////UFDo5zUAAOgyvP//BWABAABQjUXAUI2F0P7//1CNhaD+//9Q6MU1AACNhYD9//9QjYUQ/v//UOgyPgAAjYWw/f//UI2FQP7//1DoHz4AAOjqu///BQABAABQjYUw////UOgIPgAA6NO7//8FMAEAAFCNhWD///9Q6PE9AADovLv//wWgAAAAUI2FMP///1CNhRD+//9QjUXAUOjvOAAAjYVA/v//UI2FEP7//1CNhWD///9QjYUw////UOg+wv//6Hm7//8FoAAAAFCNRcBQUOh5MAAAjUXAUI2FQP7//1CNhRD+//9Q6JLK///HRfAAAAAA6Ea7//8FAAEAAIlF9I2FgP3//4lF+I2FEP7//4lF/I2FoP7//1DokDsAAIvYjYUA////UOiCOwAAO8MPR9iNhQD///+Nc/9WUOh9RAAAC8J0B78BAAAA6wIz/1aNhaD+//9Q6GNEAAALwnQHvgIAAADrAjP2C/eNRZCLdLXwVlDo9jwAAI1GMFCNhXD+//9Q6OY8AACNRcBQ6K0lAACNc/7HRcAB" & _
                                                    "AAAAx0XEAAAAAIX2D4joAAAADx9AAI1FwFCNhXD+//9QjUWQUOi8uv//Vo2FAP///1Do70MAAAvCdAe/AQAAAOsCM/9WjYWg/v//UOjVQwAAC8J0B7gCAAAA6wIzwAvHi3yF8IX/D4SFAAAAV42FMP///1DoXTwAAI1HMFCNhWD///9Q6E08AACNRcBQjYVg////UI2FMP///1DoRsn//+gBuv//BaAAAABQjYUw////UI1FkFCNheD9//9Q6DQ3AACNhXD+//9QjUWQUI2FYP///1CNhTD///9Q6IbA//+NheD9//9QjUXAUFDo5TQAAIPuAQ+JHP///+inuf//BaAAAABQjUXAUFDopy4AAI1FwFCNhXD+//9QjUWQUOjDyP//jUWQUOh6uf//BWABAABQ6B8lAABfW4P4AXQV6GO5//8FYAEAAFCNRZBQUOhjPwAAjYXQ/v//UI1FkFDo8yQAAPfYXhvAQIvlXcIMADPAXovlXcIMAMzMzMzMzMzMzMzMzMzMVYvsgeywAQAAjYVQ/v//Vv91CFDoV/f//4t1EI2FMP///1ZQ6Afv//+NRiBQjUWQUOj67v//jYUw////UOieJQAAhcAPhVQDAACNRZBQ6I0lAACFwA+FQwMAAI2FMP///1Doubj//4PogFDo0CQAAIP4AQ+FJQMAAI1FkFDonrj//4PogFDotSQAAIP4AQ+FCgMAAFNX6IW4//+D6IBQjUWQUI1F4FDo5C8AAP91DI2FUP///1Dode7//+hguP//g+iAUI1F4FCNhVD///9QUOirMwAA6Ea4//+D6IBQjUXgUI2FMP///1CNhRD///9Q6IszAACNhVD+//9QjYWw/v//UOioOgAAjYVw/v//"
Private Const STR_THUNK3                As String = "UI2F0P7//1DolToAAOgAuP//g8BAUI2FcP///1DogDoAAOjrt///g8BgUI1FkFDobjoAAOjZt///UI2FcP///1CNhbD+//9QjUXgUOhBNQAAjYXQ/v//UI2FsP7//1CNRZBQjYVw////UOhjv///6J63//9QjUXgUFDoAy8AAI1F4FCNhdD+//9QjYWw/v//UOgMx///x0XQAAAAAOhwt///g8BAiUXUjYVQ/v//iUXYjYWw/v//iUXcjYUQ////UOgMOAAAi/iNhVD///9Q6P43AAA7xw9H+I2FUP///41f/1NQ6KlAAAALwnQHvgEAAADrAjP2U42FEP///1Doj0AAAAvCdAe4AgAAAOsCM8AL8I1FsIt0tdBWUOiCOQAAjUYgUI2F8P7//1DocjkAAI1F4FDoOSIAAI13/sdF4AEAAADHReQAAAAAhfYPiNIAAACNReBQjYXw/v//UI1FsFDobLj//1aNhVD///9Q6B9AAAALwnQHvwEAAADrAjP/Vo2FEP///1DoBUAAAAvCdAe4AgAAAOsCM8ALx4t8hdCF/3R3V42FcP///1Do8TgAAI1HIFCNRZBQ6OQ4AACNReBQjUWQUI2FcP///1Do0MX//+g7tv//UI2FcP///1CNRbBQjYWQ/v//UOijMwAAjYXw/v//UI1FsFCNRZBQjYVw////UOjIvf//jYWQ/v//UI1F4FBQ6LcyAACD7gEPiS7////o6bX//1CNReBQUOhOLQAAjUXgUI2F8P7//1CNRbBQ6FrF//+NRbBQ6MG1//+D6IBQ6NghAABfW4P4AXQT6Ky1//+D6IBQjUWwUFDovj0AAI2FMP///1CNRbBQ6K4hAAD32F4bwECL5V3CDAAzwF6L5V3CDADMzMzMzMzMzMxV" & _
                                                    "i+yLTQiLwcHoB4Hhf39//yUBAQEBA8lrwBszwV3CBADMzMzMzMzMzMzMzMzMzMxVi+zoWLX//7kAjCQBgekAQCQBA8GLTQhRUP91EI1BMP91DGoQUI1BIFDoscv//13CDADMzMzMzMzMzMzMzMzMVYvsi00Ii0UQAUE4g1E8AIlFEIlNCF3ppP///8zMzMxVi+xWi3UIg35IAXUNVugtAAAAx0ZIAgAAAItFEAFGQFD/dQyDVkQAVuhy////Xl3CDADMzMzMzMzMzMzMzMzMVYvsVot1CItOMIXJdCm4EAAAACvBUI1GIAPBagBQ6M0DAACDxAyNRiBQVugQAAAAx0YwAAAAAF5dwgQAzMzMzFWL7IPsEI1F8FZXUP91DOg83P//i3UIjUXwjX4QV1dQ6Hvb//9XVlfow9z//19ei+VdwggAzMzMzMzMzMzMzMxVi+yD7BRTVot1CItGSIP4AXQFg/gCdQ1W6GL////HRkgAAAAAi144i1Y8D6TaA2oIi8LB4wPB6BiLy4hF7IvCwegQiEXti8LB6AiIRe4PtsKIRe+Lwg+swRiJVfzB6BiITfCLwovLiF3zD6zBEMHoEIvDiE3xD6zQCIhF8o1F7FDB6ghW6Fb+//+LXkCLVkQPpNoDagiLwsHjA8HoGIvLiEXsi8LB6BCIRe2LwsHoCIhF7g+2wohF74vCD6zBGIlV/MHoGIhN8IvCi8uIXfMPrMEQwegQi8OITfEPrNAIiEXyjUXsUMHqCFbo8f3///91DI1GEFDoRdz//15bi+VdwggAzMzMzMzMzMzMzMzMzFWL7FaLdQhqUGoAVuhPAgAAg8QMVv91DOjj2v//x0ZIAQAAAF5dwggAzMzMzMzMzFWL7IHs" & _
                                                    "gAAAALkgAAAAU4tdDFZXi/ONfYDzpb79AAAAjUWAUFDo5hcAAIP+AnQQg/4EdAtTjUWAUFDo4QMAAIPuAXnci30IjXWAuSAAAADzpV9eW4vlXcIIAMzMzMzMzFWL7FNWi3UIV1boAf3//4vYU+j5/P//i9BS6PH8//+L+DP+i/eLxzPDwc8IM/LBwAiLzsHJEDPBM8czxl8zwzNFCF5bXcIEAMzMzMzMzMzMVYvsVot1CP826KL/////dgSJBuiY/////3YIiUYE6I3/////dgyJRgjogv///4lGDF5dwgQAzMzMzMzMzMzMzFWL7FOLXQhWVw+2ewcPtkMCD7ZzCw+2Uw/B5wgL+A+2SwMPtkMNwecIC/jB5ggPtkMIwecIC/jB4ggPtkMGC/DB4QgPtkMBweYIC/APtkMMweYIC/APtkMKC9APtkMFweIIC9APtgPB4ggL0A+2Qw4LyIlTDA+2QwnB4QgLyIlzCA+2QwSJewTB4QhfC8heiQtbXcIEAMzMzMzMzMzMzMxVi+xW6Eex//+LdQgFgwYAAFD/NuiXFwAAiQboMLH//wWDBgAAUP92BOiCFwAAiUYE6Bqx//8FgwYAAFD/dgjobBcAAIlGCOgEsf//BYMGAABQ/3YM6FYXAACJRgxeXcIEAMzMzMzMzMzMzMzMzMzMVYvsi0UIi9BWi3UQhfZ0FVeLfQwr+IoMF41SAYhK/4PuAXXyX15dw8zMzMzMzMzMVYvsi00Qhcl0Hw+2RQxWi/FpwAEBAQFXi30IwekC86uLzoPhA/OqX16LRQhdw8zMVYvsVot1CFboA/v//4vQi84z1sHJEMHCCMHOCDPRM9Yzwl5dwgQAzMzMzMzMzMzMVYvsVot1CP82" & _
                                                    "6ML/////dgSJBui4/////3YIiUYE6K3/////dgyJRgjoov///4lGDF5dwgQAzMzMzMzMzMzMzFWL7IPsYA9XwMdFoAEAAABWV41FoMdFpAAAAABQDxFFqMdF0AEAAAAPEUW4x0XUAAAAAGYP1kXIDxFF2A8RRehmD9ZF+OjGr///BaAAAABQjUWgUOg3FwAAjUWgUOgeMAAAi30IjXD/g/4BdiwPHwCNRdBQUOiGLAAAVo1FoFDoDDkAAAvCdAtXjUXQUFDorSoAAE6D/gF3141F0FBX6J0xAABfXovlXcIEAMzMzMzMVYvsg+xAVg9XwMdFwAEAAABXjUXAx0XEAAAAAFAPEUXIx0XgAQAAAGYP1kXYx0XkAAAAAA8RRehmD9ZF+Oger///UI1FwFDo1BgAAI1FwFDoyy8AAIt9CI1w/4P+AXYpjUXgUFDoFiwAAFaNRcBQ6Gw4AAALwnQLV41F4FBQ6J0rAABOg/4Bd9eNReBQV+hdMQAAX16L5V3CBADMzMzMzFWL7IHsAAEAAItFDA9XwFNWV7k8AAAAZg8ThQD///+NtQD////HRfwQAAAAjb0I////86WLTRCNnQj///+DwRCL0yvCiU34iUUMZg8fRAAAi/nHRRAEAAAAi/MPH0QAAP90GAT/NBj/d/T/d/Dozrr//wFG+ItFDBFW/P90GAT/NBj/d/z/d/jos7r//wEGi0UMEVYE/3QYBP80GP93BP836Jq6//8BRgiLRQwRVgz/dBgE/zQY/3cM/3cI6H+6//8BRhCNfyCLRQwRVhSNdiCDbRABdYqLTfiDwwiDbfwBD4Vq////M/ZqAGom/3T1hP909YDoR7r//wGE9QD///9qABGU9QT///9qJv90" & _
                                                    "9Yz/dPWI6Ci6//8BhPUI////agARlPUM////aib/dPWU/3T1kOgJuv//AYT1EP///2oAEZT1FP///2om/3T1nP909Zjo6rn//wGE9Rj///9qABGU9Rz///9qJv909aT/dPWg6Mu5//8BhPUg////EZT1JP///4PGBYP+Dw+CWf///4tdCI21AP///7kgAAAAi/vzpVPo+bz//1Po87z//19eW4vlXcIMAMzMzMzMzMzMzMxVi+yD7BBTVot1DFeLfRhqAFZqAP91FOhkuf//agBWagBXiUXwi9roVLn//2oA/3UQiUX0i/JqAFfoQrn//2oA/3UQiUX8agD/dRSJVfjoLbn//4v4i0X0A/uD0gAD+BPWO9Z3DnIEO/hzCINF/ACDVfgBi0UIM8kLTfCJCDPJA1X8iXgEE034X16JUAiJSAxbi+VdwhQAzMzMzMzMzMzMVYvsg+wwU1aLdQhXi30MV1boei4AAGogV41F0FDoLhkAAIvYiVUMjU4Ig8Y4jUXQUFFR6KgTAAADw4kGE1UMi0UIg8AQiVYEV1BQ6JATAACLfQiJR0CNRdBQV1eJV0ToDDIAAItNDAPYE8qLVzAr04tfNBvZO180ci13BTtXMHYmgwb/iwaDVgT/I0YEg/j/dRWDRgj/jXYIg1YE/4sGI0YEg/j/dOuJXzSJVzBfXluL5V3CCADMzMzMzMzMzMzMVYvsgewIAQAAjYV4////U1ZX/3UMUOilCQAAjYV4////UOhZu///jYV4////UOhNu///jYV4////UOhBu///jb34/v//uwIAAABmDx9EAACLjXj///+LhXz///+B6e3/AACJjfj+//+D2ACJhfz+//+4CAAAAGZmDx+EAAAAAACL" & _
                                                    "dAf4i0wH/IuUBXj///+JdfgPrM4Qi4wFfP///4PmAcdEB/wAAAAAK9aD2QCB6v//AACJlAX4/v//g9kAiYwF/P7//w+3TfiJTAf4g8AIg/h4cqyLjWj///+LhWz///+LVfAPrMEQD7eFaP///4PhAYmFaP///yvRx4Vs////AAAAAItN9LgBAAAAg9kAger/fwAAiZVw////g9kAiY10////D6zKEIPiAcH5ECvCUI2F+P7//1CNhXj///9Q6I0HAACD6wEPhQT///+LdQgz0oqE1Xj///+LjNV4////iARWi4TVfP///w+swQiITFYBQsH4CIP6EHLXX15bi+VdwggAzMzMzMzMzMzMzMzMzFWL7ItFCDPSVleLfQwr+I1yEYsMB41ABANI/APRD7bKweoIiUj8g+4BdedfXl3CCADMzMzMzMzMzMzMzMzMzMxVi+xW/3UMi3UIVuiw////jUZEUFbotgEAAF5dwggAzFWL7IPsRFNWi3UIVw8QBotGQIlF/A8RRbwPEEYQDxFFzA8QRiAPEUXcDxBGMA8RRezoeqn//wU0BQAAUI1FvFDoW////4tF/I19vPfQjVXMJYAAAAAr/rkCAAAAjVj/99DB6B/B6x8j2PfbK9aLw/fQiUUIZg9uw2YPcNAAZg9uwIvGZg9w2AAPH4QAAAAAAI1AIA8QQOAPEEwH4GYP28JmD9vLZg/ryA8RSOAPEEDwDxBMAuBmD9vCZg/by2YP68gPEUjwg+kBdcaNVkCNcQGLDDqNUgQjTQiLwyNC/AvIiUr8g+4BdehfXluL5V3CBADMzMzMzMzMzMzMzMzMzMxVi+yD7ESNRbxWakRqAFDo7Pf//4t1CIPEDDPAi5aoAAAAhdJ0" & _
                                                    "G2ZmDx+EAAAAAAAPtowGmAAAAIlMhbxAO8Jy741FvMdElbwBAAAAUFbojf7//16L5V3CBADMzMzMzMxVi+xWi3UIM8Az0g8fRAAAAwSWD7bIiQyWQsHoCIP6EHzuA0ZAi8jB6AKD4QMz0olOQI0MgAMMlg+2wYkElkLB6QiD+hB87gFOQF5dwgQAzFWL7IPsVItFDI1NrFNWi3UIM9srwcdF+BAAAABXiUXwM9Iz/zPAiVUIiVX8hdt4UY1LAYP5Anwwi03wjVWsjQyZA9GLDIaNUvgPr0oIAU0Ii0yGBIPAAg+vSgQBTfyNS/87wX7ei1UIO8N/Dot9DIvLK8iLPI8PrzyGi0X8A8ID+I1DATPSiVUIi8iJVfyJRfSD+BF9coN9+AJ8Q4tVDIvDK8GNFIKDwkAPH4AAAAAAiwSOjVL4D69CDI0EgMHgBgFFCItEjgSDwQIPr0IIjQSAweAGAUX8g/kQfNSLVQiD+RF9GotVDIvDK8GLRIJED68EjotVCI0EgMHgBgP4i0X8A8ID+ItF9ItN+EmJfJ2siU34i9iD+f8PjwL///+NRaxQ6In+//8PEEWsi0XsXw8RBg8QRbwPEUYQDxBFzA8RRiAPEEXcDxFGMIlGQF5bi+VdwggAzMzMzMzMzMzMzMxVi+yLVQyD7EQzwA8fRAAAD7YMEIlMhbxAg/gQfPKNRbzHRfwBAAAAUP91COif/P//i+VdwggAzMzMzMzMzMzMVYvsgex8AQAAU1ZXagz/dQyNReDGRdwAD1fAx0XlAAAAAFBmD9ZF3WbHRekAAMZF6wDoSfX//4PEDMZFvACNRdzHRdUAAAAAD1fAZsdF2QAADxFFvWoEUGog/3UIjYUw////Zg/WRc1Q" & _
                                                    "xkXbAOi+xP//aiCNRbxQUI2FMP///1DoC77//41FzFCNRbxQjYWE/v//UOj3z///D1fADxFFvIpFvGogjUW8UFCNhTD///9QDxFFzOjWvf//i3UUD1fAVv91EA8RRbyKRbyNhYT+///GRewAUA8RRczHRfUAAAAAZg/WRe1mx0X5AADGRfsA6GvQ//+LxvfYg+APUI1F7FCNhYT+//9Q6FPQ//+DfSQBi30gi10cU3UUV/91GI2FMP///1DoZr3//1NX6wP/dRiNhYT+//9Q6CPQ//+Lw/fYg+APUI1F7FCNhYT+//9Q6AvQ//8z0ohd9IvGiVXoiEXsi8iLwg+swQhqEMHoCIhN7YvCi84PrMEQwegQiE3ui8KLzg+swRjB6BgPtsKIRfCLwsHoCIhF8YvCwegQiEXyweoYiE3vi8uIVfMz0ovCiVXoD6zBCMHoCIhN9YvCi8sPrMEQwegQiE32i8KLyw+swRjB6BgPtsKIRfiLwsHoCIhF+YvCwegQiEX6jUXsUI2FhP7//8HqGFCITfeIVfvoW8///4N9JAF1M/91KI2FhP7//1Do9s3//2p8jYUw////agBQ6Ibz//+KhTD///+DxAwzwF9eW4vlXcIkAI1FrFCNhYT+//9Q6MLN//+LdSiNTayLwTLbuhAAAAAr8JCKBA6NSQEyQf8K2IPqAXXwi0UchNt1P1BX/3UYjYUw////UOgIvP//anyNhTD///9qAFDoGPP//4qFMP///4PEDA9XwA8RRayKRaxfXjPAW4vlXcIkAIXAdA5QagBX6O3y//+KB4PEDGp8jYUw////agBQ6Njy//+KhTD///+DxAwPV8APEUWsikWsX164AQAAAFuL5V3CJADMzMzM" & _
                                                    "zMzMVYvsVleLfQgPtgeZi8iL8g+2RwEPpM4ImcHhCAvyC8gPtkcCD6TOCJnB4QgL8gvID7ZHAw+kzgiZweEIC/ILyA+2RwQPpM4ImcHhCAvyC8gPtkcFD6TOCJnB4QgL8gvID7ZHBg+kzgiZweEIC/ILyA+2RwcPpM4ImcHhCAvBC9ZfXl3CBADMzMzMzMzMzMzMVYvsg+wIi0UQSPfQmVOLXQiJRfiLRQyJVfzzD35d+I1LeFYz9mYPbNuNUHg7wXdLO9NyRyvYx0UQEAAAAFdmkIs8GI1ACIt0GPyLSPiLUPwzzyNN+DPWI1X8M/kz8ol8GPiJdBj8MUj4MVD8g20QAXXOX15bi+VdwgwAi9ONSBAr0A8QDPONSSAPEFHQZg/v0WYP29MPKMJmD+/BDxEE84PGBA8QQdBmD+/QDxFR0A8QTArgDxBR4GYP79FmD9vTDyjCZg/vwQ8RRArgDxBB4GYP78IPEUHgg/4QcqVeW4vlXcIMAMzMzMzMzMzMzMzMVYvsi1UMi0UIK9BWvhAAAACLDAKNQAiJSPiLTAL8iUj8g+4BdeteXcIIAMzMzMzMVYvsi0UQVleD+BB0OYP4IHVfi3UMi30IahBWV+iv8P//ahCNRhBQjUcQUOig8P//g8QY6Hih//8FIQUAAIlHMF9eXcIMAIt1DIt9CGoQVlfoe/D//2oQjUcQVlDob/D//4PEGOhHof//BRAFAACJRzBfXl3CDADMzMzMzMzMzMxVi+yD7GyLRQiNVZRTVruQAQAAM/aLSASJTfiLSAiJTfSLSAyJTeiLSBCJTfyLSBSJTfCLSBiJTeyLTQyDwQKJddxXizgr04tAHIl94IlF5IlN2IldDIlV1A8fgAAAAACD" & _
                                                    "/hBzKQ+2cf4PtkH/weYIC/APtgHB5ggL8A+2QQHB5ggL8IPBBIk0GolN2OtUjV4Bg+YPjUP9g+APjX2UjTy3i0yFlIvDg+APi/HBxg+LVIWUi8HBwA0z8MHpCjPxi8KLysHIB8HBDjPIweoDjUP4M8qLXQyD4A8D8QN0hZQDN4k36Emg//+LffyL18HKC4vPwcEHM9GLz8HJBvfXI33sM9GLDBiDwwSLRfADyiNF/APOi3XgM/iL1oldDMHKDYvGwcAKA/kDfeQz0IvGwcgCM9CLRfiLyCPGM84jTfQzyItF7IlF5APRi0Xwi034iUXsi0X8iUXwi0XoA8eJdfiLddwD+otV1EaJRfyLRfSJTfSLTdiJReiJfeCJddyB+5ACAAAPgtf+//+LRQiLTfiLVfwBSASLTfQBSAgBUBABOItN6ItV8AFIDAFQFItV7ItN5AFQGAFIHP9AYF9eW4vlXcIIAMzMzMzMzMzMzMzMzFWL7IHs4AAAAFNWi3UIu5ACAABXiV24iwaJReyLRgSJRfCLRgyLfgiJReCLRhCJRdSLRhSJRdCLRhiJRbSLRhyJRbCLRiCJReiLRiSJRfSLRiiJRcyLRiyJRciLRjCJRcSLRjSJRcCLRjiJRayLRjyLdQyJfdiNvSD///+JRagzwCv7iUXciX2gDx+AAAAAAIP4EHMfVuhl+///i8iDxgiLwolNDIlF5IkMH4lEHwTpEwEAAI1QAcdFDAAAAACNQv2D4A+LjMUg////i4TFJP///4lF+IvCg+APiU38jY0g////i5TFIP///4v6i5zFJP///4tF3IPgD4lVvMHnGI0EwYvLiUWki8IPrMgICUUMi0W8wekIC/mLyw+syAGJfeSL+tHp" & _
                                                    "M9IL0MHnHzFVDAv5i0W8i03kD6zYBzPPMUUMi0X8wesHM8sz24lN5ItN+IvRD6TBA8HqHcHgAwvZi034C9CLRfyL+A+syBOJVbwz0gvQwekTi0W8M8LB5w2LVfwL+YtN+DPfD6zKBjPCwekGi1UMM9mLTeQD0ItF3BPLg8D5g+APA5TFIP///xOMxST///+LRaQDEIlVDBNIBIkQiU3kiUgE6JSd//+LVfQz/4tN6IvaD6TKF8HrCQv6weEXi1X0C9mLTeiJXfyL2Q+s0RKJffgz/wv5weoSMX38M/+LTejB4w4L2otV9DFd+IvZD6zRDsHjEgv5weoOMX38C9qLTfiLVbgzy4td/It96PfXAxwQE0wQBCN9xItV9ItFyPfSI0X0I1XAM9CJTfiLTcwjTeiLRfgz+YtN8APfE8IDXQwTReQDXayJXfwTRagz24lF+ItF7IvQD6zIHMHiBMHpHAvYi0XsC9GLTfCL+Q+kwR6JVQwz0sHvAgvRweAeC/gz3zFVDDPSi03wi/mLRewPpMEZwe8HC9HB4BkxVQwL+ItN2DPfi1Xgi/kzfewjfdQjTewzVfAz+SNV0ItF4CNF8ItNxDPQi0UMA9+LffgTwolNrItNwItV/ANVtIlNqBN9sItNzANd/IlNxItNyIlNwItN6IlNzItN9Il99It91Il9tIt90Il9sIt92Il91It94Il90It97IlNyIvIE034i0XciV3sQItduIl92IPDCIt98Il94It9oIlV6IlN8IlF3IlduIH7EAUAAA+CG/3//4t1CItF7It92AEGi0XgEU4Ei8oBfgiLfbQRRgyLRdQBRhCLRdARRhQBfhiLRbARRhwBTiCLRfQRRiSLRcwBRiiLRcgR" & _
                                                    "RiyLRcQBRjCLRcARRjSLTawBTjiLTagRTjz/hsAAAABfXluL5V3CCADMzMzMzMzMzMzMzMzMzFWL7FOLXQhWVw+2ewcPtkMKD7ZzCw+2Uw/B5wgL+A+2SwMPtkMNwecIC/jB5ggPtgPB5wgL+MHiCA+2Qw4L8MHhCA+2QwHB5ggL8A+2QwTB5ggL8A+2QwIL0A+2QwXB4ggL0A+2QwjB4ggL0A+2QwYLyIl7BA+2QwnB4QgLyIlzCA+2QwzB4QhfC8iJUwxeiQtbXcIEAMzMzMzMzMzMzMxVi+yLRQxQUP91COgA7P//XcIIAMzMzMzMzMzMzMzMzFWL7ItFEFNWi3UIjUh4V4t9DI1WeDvxdwQ70HMLjU94O/F3MDvXciwr+LsQAAAAK/CLFDgrEItMOAQbSASNQAiJVDD4iUww/IPrAXXkX15bXcIMAIvXjUgQi94r0CvYK/64BAAAAI12II1JIA8QQdAPEEw34GYP+8gPEU7gDxBMCuAPEEHgZg/7yA8RTAvgg+gBddJfXltdwgwAzMzMzMxVi+xW6Aea//+LdQgFeAUAAFD/NuhXAAAAiQbo8Jn//wV4BQAAUP92BOhCAAAAiUYE6NqZ//8FeAUAAFD/dgjoLAAAAIlGCOjEmf//BXgFAABQ/3YM6BYAAACJRgxeXcIEAMzMzMzMzMzMzMzMzMzMVYvsi1UMU4tdCIvDwegYi8tWwekID7bJD7Y0EIvDwegQD7bAD7YMEcHmCA+2BBALxsHgCAvBD7bLweAIXlsPtgwRC8FdwggAzMzMzMzMzMxVi+yLTQxTi10IVoPDEMdFDAQAAABXg8EDDx+AAAAAAA+2Qf6NWyCZjUkIi/CL+g+2QfUPpPcImcHmCAPw" & _
                                                    "iXPQE/qJe9QPtkH3mYvwi/oPtkH4mQ+kwgjB4AgD8Ilz2BP6iXvcD7ZB+pmL8Iv6D7ZB+Q+k9wiZweYIA/CJc+AT+ol75A+2QfyZi/CL+g+2QfsPpPcImcHmCAPwiXPoE/qDbQwBiXvsD4V0////i00IX15bgWF4/38AAMdBfAAAAABdwggAzMzMzMzMzMzMzMzMVYvsg+wMU4tdDA9XwFZXi30QixOL8otDBIvIZg8TRfQDNxNPBDvydQY7yHUE6xg7yHcPcgQ78nMJuAEAAAAz0usLZg8TRfSLRfSLVfiLfQiJTwSJN4tLCIt1EIlN/ItLDIlN+ItOCANN/IlNCItODBNN+ItdCAPYiV0IE8o7XfyLXQx1BTtLDHQjO0sMdxNyCItDCDlFCHMJuAEAAAAz0usLZg8TRfSLVfiLRfSLdQiJTwyJdwiLSxCLdRCJTfyLSxSJTfiLThADTfyJTQiLThQTTfiLXQgD2IldCBPKO138i10MdQU7SxR0IztLFHcTcgiLQxA5RQhzCbgBAAAAM9LrC2YPE0X0i1X4i0X0i3UIiU8UiXcQi0sYi3UQiU38i0sciU34i04YA038iU0Ii04cE034i10IA9iJXQgTyjtd/ItdDHUFO0scdCM7Sxx3E3IIi0MYOUUIcwm4AQAAADPS6wtmDxNF9ItV+ItF9It1CIlPHIl3GItLIIt1EIlN/ItLJIlN+ItOIANN/IlNCItOJBNN+ItdCAPYiV0IE8o7XfyLXQx1BTtLJHQjO0skdxNyCItDIDlFCHMJuAEAAAAz0usLZg8TRfSLVfiLRfSLdQiJdyCLdRCJTySLSyiLWyyLdigD8YlNDItNEItJLBPLA/ATyjt1DHUEO8t0LDvL" & _
                                                    "dx1yBTt1DHMWiXcouAEAAACJTywz0l9eW4vlXcIMAGYPE0X0i1X4i0X0iXcoiU8sX15bi+VdwgwAzMzMzMzMVYvsg+wIU4tdDA9XwFZXi30QixOL8otDBIvIZg8TRfgDNxNPBDvydQY7yHUE6xg7yHcPcgQ78nMJuAEAAAAz0usLZg8TRfiLRfiLVfyLfQiJTwSLTRCJN4txCANzCItJDBNLDAPwE8o7cwh1BTtLDHQgO0sMdxByBTtzCHMJuAEAAAAz0usLZg8TRfiLVfyLRfiJTwyLTRCJdwiLcRADcxCLSRQTSxQD8BPKO3MQdQU7SxR0IDtLFHcQcgU7cxBzCbgBAAAAM9LrC2YPE0X4i1X8i0X4iU8UiXcQi0sYi1sciU0Mi00Qi3EYA3UMi0kcE8sD8BPKO3UMdQQ7y3QsO8t3HXIFO3UMcxaJdxi4AQAAAIlPHDPSX15bi+VdwgwAZg8TRfiLVfyLRfiJdxiJTxxfXluL5V3CDADMzMzMzMxVi+yLTQjHAQAAAADHQQQAAAAAiwGJQQiLQQSJQQyLQQiJQRCLQQyJQRSLQRCJQRiLQRSJQRyLQRiJQSCLQRyJQSSLQSCJQSiLQSSJQSxdwgQAzMzMzMzMzMzMzMzMzMxVi+yLRQjHAAAAAADHQAQAAAAAx0AIAAAAAMdADAAAAADHQBAAAAAAx0AUAAAAAMdAGAAAAADHQBwAAAAAXcIEAMzMzMzMzMzMzMzMzMzMzFWL7ItNDLoFAAAAU4tdCFYr2Y1BKFeJXQgPH4AAAAAAizQDi1wDBIt4BIsIO993LnIiO/F3KDvfchp3BDvxchSLXQiD6AiD6gF51V9eM8BbXcIIAF9eg8j/W13CCABfXrgBAAAA"
Private Const STR_THUNK4                As String = "W13CCADMzMzMzMxVi+yLTQy6AwAAAFOLXQhWK9mNQRhXiV0IDx+AAAAAAIs0A4tcAwSLeASLCDvfdy5yIjvxdyg733IadwQ78XIUi10Ig+gIg+oBedVfXjPAW13CCABfXoPI/1tdwggAX164AQAAAFtdwggAzMzMzMzMVYvsi1UIM8APH4QAAAAAAIsMwgtMwgR1D0CD+AZy8bgBAAAAXcIEADPAXcIEAMzMVYvsi1UIM8APH4QAAAAAAIsMwgtMwgR1D0CD+ARy8bgBAAAAXcIEADPAXcIEAMzMVYvsg+wQU4tdELlAAAAAVot1CCvLV4t9DGYPbsOJTRCLB4tXBIlF+IlV/PMPfk34Zg/zyGYP1g7oo5///4tNEIlF8ItHCIlV9ItXDIlF+IlV/PMPfk34Zg9uw2YP88jzD35F8GYP68hmD9ZOCOhun///i00QiUXwi0cQiVX0i1cUiUX4iVX88w9+TfhmD27DZg/zyPMPfkXwZg/ryGYP1k4Q6Dmf//+LTRCJRfCLRxiJVfSLVxyJRfiJVfzzD35N+GYPbsNmD/PI8w9+RfBmD+vIZg/WThjoBJ///4tNEIlF8ItHIIlV9ItXJIlF+IlV/PMPfk34Zg9uw2YP88jzD35F8GYP68hmD9ZOIOjPnv//iUXwi0coiVX0i1csiUX4iVX88w9+TfiLTRBmD27DZg/zyPMPfkXwZg/ryGYP1k4o6Jqe//9fXluL5V3CDADMVYvsg+wQU4tdELlAAAAAVot1CCvLV4t9DGYPbsOJTRCLB4tXBIlF+IlV/PMPfk34Zg/zyGYP1g7oU57//4tNEIlF8ItHCIlV9ItXDIlF+IlV/PMPfk34Zg9uw2YP88jzD35F8GYP68hmD9ZOCOgenv//i00QiUXw" & _
                                                    "i0cQiVX0i1cUiUX4iVX88w9+TfhmD27DZg/zyPMPfkXwZg/ryGYP1k4Q6Omd//+LTRCJRfCLRxiJVfSLVxyJRfiJVfzzD35N+GYPbsNmD/PI8w9+RfBmD+vIZg/WThjotJ3//19eW4vlXcIMAMzMzMzMzMzMzMzMVYvsg+xoU1aLdQyNXjBT6Ez9//+FwA+FZAIAAFcPHwCNRZgPV8BQZg8TRfjon/v//41FyFDolvv//1ONRZhQ6Czk//9T6Ib7//+LFov6A32Yi0YEi8gTTZw7+nUGO8h1BOsbO8h3D3IEO/pzCbgBAAAAM9LrDg9XwGYPE0X4i0X4i1X8i14MiT6LfggDfaCJTgSLyxNNpAP4E8o7fgh1BDvLdCI7y3cQcgU7fghzCbgBAAAAM9LrDg9XwGYPE0X4i1X8i0X4i14UiX4Ii34QA32oiU4Mi8sTTawD+BPKO34QdQQ7y3QiO8t3EHIFO34Qcwm4AQAAADPS6w4PV8BmDxNF+ItV/ItF+IteHIl+EIt+GAN9sIlOFIvLE020A/gTyjt+GHUEO8t0IjvLdxByBTt+GHMJuAEAAAAz0usOD1fAZg8TRfiLVfyLRfiLXiSJfhiLfiADfbiJThyLyxNNvAP4E8o7fiB1BDvLdCI7y3cQcgU7fiBzCbgBAAAAM9LrDg9XwGYPE0X4i1X8i0X4i14siX4gi34oA33AiU4ki8sTTcQD+BPKO34odQQ7y3QiO8t3EHIFO34ocwm4AQAAADPS6w4PV8BmDxNF+ItV/ItF+IteNIl+KIt+MAN9yIlOLIvLE03MA/gTyjt+MHUEO8t0IjvLdxByBTt+MHMJuAEAAAAz0usOD1fAZg8TRfiLVfyLRfiLXjyJfjCL" & _
                                                    "fjgDfdCJTjSLyxNN1AP4E8o7fjh1BDvLdCI7y3cQcgU7fjhzCbgBAAAAM9LrDg9XwGYPE0X4i1X8i0X4iU48jV4wi03YA8iJfjiLRdwTwgFOQFMRRkTo6fr//4XAD4Sh/f//X+hLjv//BaAAAABQVujv+f//hcB+J+g2jv//BaAAAABQVlboORQAAOgkjv//BaAAAABQVujI+f//hcB/2Vb/dQjoOxAAAF5bi+VdwggAzMzMVYvsg+woU4tdCFZXi30MV1PoehAAAItHLA9XwIlF5ItHMIlF6ItHNIlF7ItHOIlF8ItHPIlF9I1F2GoBUFBmDxNF2MdF4AAAAADo8fv//4vwjUXYUFNT6GT3//+LTzgD8ItHMItXPIlF5DPAC0c0iUXojUXYagFQUMdF4AAAAACJTeyJVfDHRfQAAAAA6K77//8D8I1F2FBTU+gh9///A/DHReQAAAAAi0cgD1fAiUXYi0ckiUXci0coiUXgi0c4iUXwi0c8iUX0jUXYUFNTZg8TRejo5/b//4tPJAPwM8CJTdgLRyiJRdyLRzCLVzSLyolF+DPAC0csiUXgi0c4iUXoi0c8iUXsM8ALRyCJRfSNRdhQU1OJTeSJVfDon/b//4tPLAPwi1c0M8ALRzAPV8CJRdyLRyCJRfCNRdiJTdgzyQtPKFBTU4lV4MdF5AAAAABmDxNF6IlN9OjBFAAAi1ckK/CLRzAPV8CJRdixIItHNIlF3ItHOIlF4ItHPIlF5ItHIGYPE0Xo6EKZ//8LVyyJRfCNRdhQU1OJVfTofhQAAItVDCvwi080M8ALRziLXySJRdwzwAtHPIlN2ItPIDP/iUXgi0Ioi1IsiU3ksSDo25j//wvYx0XwAAAAAIld" & _
                                                    "6Av6i10MiX3si30Ii0MwiUX0jUXYUFdX6CMUAAAr8MdF4AAAAACLQziJRdiLQzyJRdyLQySJReSLQyiJReiLQyyJReyLQzSJRfSNRdhQV1fHRfAAAAAA6OQTAAAr8Hkg6LuL//9QV1foc/X//wPweO9fXluL5V3CCABmDx9EAACF9nURV+iWi///UOiw9///g/gBdNzohov//1BXV+ieEwAAK/Dr2szMzMzMzMzMzMxVi+xW/3UQi3UI/3UMVujd8v//C8J1Df91FFboAPf//4XAeAr/dRRWVuhSEQAAXl3CEADMzMzMzMzMzMzMzMzMVYvsVv91EIt1CP91DFbo3fT//wvCdQ3/dRRW6DD3//+FwHgK/3UUVlboIhMAAF5dwhAAzMzMzMzMzMzMzMzMzFWL7IHsyAAAAFaLdQxW6G33//+FwHQP/3UI6NH1//9ei+VdwgwAV1aNhTj///9Q6OwMAACLfRCNhWj///9XUOjcDAAAjUXIUOij9f//jUWYx0XIAQAAAFDHRcwAAAAA6Iz1//+NhWj///9QjYU4////UOgp9v//i9CF0g+EvgEAAFOLjTj///8PV8CD4QFmDxNF+IPJAHUvjYU4////UOi8CwAAi0XIg+ABg8gAD4S/AAAAV41FyFBQ6LLx//+L8Iva6bEAAACLhWj///+D4AGDyAB1L42FaP///1DofwsAAItFmIPgAYPIAA+EEQEAAFeNRZhQUOh18f//i/CL2ukDAQAAhdIPjo8AAACNhWj///9QjYU4////UFDo4A8AAI2FOP///1DoNAsAAI1FmFCNRchQ6Gf1//+FwHkLV41FyFBQ6Cjx//+NRZhQjUXIUFDoqg8AAItFyIPgAYPIAHQRV41F" & _
                                                    "yFBQ6ATx//+L8Iva6waLXfyLdfiNRchQ6N8KAAAL8w+EmAAAAItF8IFN9AAAAICJRfDphgAAAI2FOP///1CNhWj///9QUOhRDwAAjYVo////UOilCgAAjUXIUI1FmFDo2PT//4XAeQtXjUWYUFDomfD//41FyFCNRZhQUOgbDwAAi0WYg+ABg8gAdBFXjUWYUFDodfD//4vwi9rrBotd/It1+I1FmFDoUAoAAAvzdA2LRcCBTcQAAACAiUXAjYVo////UI2FOP///1DobPT//4vQhdIPhUT+//9bjUXIUP91COjVCgAAX16L5V3CDADMzMzMzMzMzMzMzMzMVYvsgeyIAAAAVot1DFboPfX//4XAdA//dQjo0fP//16L5V3CDABXVo2FeP///1Do7AoAAIt9EI1FmFdQ6N8KAACNRdhQ6Kbz//+NRbjHRdgBAAAAUMdF3AAAAADoj/P//41FmFCNhXj///9Q6D/0//+L0IXSD4SwAQAAUw8fQACLjXj///8PV8CD4QFmDxNF+IPJAHUvjYV4////UOi+CQAAi0XYg+ABg8gAD4S2AAAAV41F2FBQ6JTx//+L8Iva6agAAACLRZiD4AGDyAB1LI1FmFDohwkAAItFuIPgAYPIAA+ECAEAAFeNRbhQUOhd8f//i/CL2un6AAAAhdIPjowAAACNRZhQjYV4////UFDomw8AAI2FeP///1DoPwkAAI1FuFCNRdhQ6ILz//+FwHkLV41F2FBQ6BPx//+NRbhQjUXYUFDoZQ8AAItF2IPgAYPIAHQRV41F2FBQ6O/w//+L8Iva6waLXfyLdfiNRdhQ6OoIAAAL8w+EkgAAAItF8IFN9AAAAICJRfDpgAAAAI2FeP///1CN" & _
                                                    "RZhQUOgPDwAAjUWYUOi2CAAAjUXYUI1FuFDo+fL//4XAeQtXjUW4UFDoivD//41F2FCNRbhQUOjcDgAAi0W4g+ABg8gAdBFXjUW4UFDoZvD//4vwi9rrBotd/It1+I1FuFDoYQgAAAvzdA2LRdCBTdQAAACAiUXQjUWYUI2FeP///1DokPL//4vQhdIPhVb+//9bjUXYUP91COjpCAAAX16L5V3CDADMVYvsgezAAAAAU1aLdRRXVuirBgAA/3UQi9iNhUD/////dQxQ6NcDAACNhXD///9Q6IsGAACL+IX/dAiBx4ABAADrDo2FQP///1DocQYAAIv4O/tzGI2FQP///1D/dQjoHAgAAF9eW4vlXcIQAI1FoFDo2vD//41F0FDo0fD//4vHK8OL2MHrBoPgP3QYUI1FoFaNBNhQ6KXy//+JRN3QiVTd1OsNjUWgVo0E2FDozgcAAItdCFPolfD//8cDAQAAAMdDBAAAAACB/4ABAAB3ElaNRaBQ6Cbx//+FwA+IggAAAI2FcP///1CNRdBQ6A7x//+FwHgWdUiNhUD///9QjUWgUOj48P//hcB/NI1FoFCNhUD///9QUOhDCwAAC8J0DlONhXD///9QUOgxCwAAjUXQUI2FcP///1BQ6CALAACLddCNRdBQweYf6HEGAACNRaBQ6GgGAAAJdcxPi3UU6WT///+NhUD///9QU+gPBwAAX15bi+VdwhAAzMzMzMzMVYvsg+xgjUWg/3UQ/3UMUOhrAgAAjUWgUP91COjf8///i+VdwgwAzMzMzMzMzMzMVYvsgeyAAAAAU1aLdRRXVuhLBQAA/3UQi9iNRYD/dQxQ6IoDAACNRaBQ6DEFAACL+IX/dAiBxwABAADr" & _
                                                    "C41FgFDoGgUAAIv4O/tzFY1FgFD/dQjo2AYAAF9eW4vlXcIQAI1FwFDolu///41F4FDoje///4vHK8OL2MHrBoPgP3QYUI1FwFaNBNhQ6FHy//+JRN3giVTd5OsNjUXAVo0E2FDoigYAAItdCFPoUe///8cDAQAAAMdDBAAAAAAPH0AAgf8AAQAAdw5WjUXAUOju7///hcB4c41FoFCNReBQ6N3v//+FwHgTdTyNRYBQjUXAUOjK7///hcB/K41FwFCNRYBQUOi4CwAAC8J0C1ONRaBQUOipCwAAjUXgUI1FoFBQ6JsLAACLdeCNReBQweYf6DwFAACNRcBQ6DMFAAAJddxPi3UU6Xf///+NRYBQU+jdBQAAX15bi+VdwhAAzMzMzFWL7IPsQI1FwP91EP91DFDoOwIAAI1FwFD/dQjoH/X//4vlXcIMAMzMzMzMzMzMzFWL7IPsYI1FoP91DFDozgUAAI1FoFD/dQjoIvL//4vlXcIIAMzMzMzMzMzMzMzMzFWL7IPsQI1FwP91DFDoPgcAAI1FwFD/dQjowvT//4vlXcIIAMzMzMzMzMzMzMzMzFWL7Fb/dRCLdQj/dQxW6K0IAAALwnQK/3UUVlboD+r//15dwhAAzMzMzMzMzMzMzFWL7Fb/dRCLdQj/dQxW6I0KAAALwnQK/3UUVlboH+z//15dwhAAzMzMzMzMzMzMzFWL7IPsYFMPV8BWZg8TRdiLRdxXZg8TRdAz/4td1IlF/DP2jUf7g/8GD1fAZg8TRfSLVfQPQ/A79w+H0gAAAItNEIvHDxBF0CvGDxFFwI0cwYtF+IlF8IlV+GYPH0QAAIP+Bg+DowAAAP9zBItFDP8z/3TwBP808I1FsFDoz9T/" & _
                                                    "/4PsEIvMg+wQDxAADxAIi8QPEQEPEEXADxFN4A8RAI1FoFDoOI///2YPc9kMDxAQZg9+yA8owmYPc9gMZg9+wQ8RVcCJTfwPEVXQO8h3E3IIi0XYO0Xocwm4AQAAADPJ6w4PV8BmDxNF6ItN7ItF6ItV+APQi0XwiVX4E8FGg+sIiUXwO/cPhlT///+LXdTrA4tF+ItNCIt10Ik0+Yvxi8qL0IlV3Ilc/gRHi3XYi138iXXQiV3UiU3YiVX8g/8LD4Lb/v//i0UIX4lwWF6JWFxbi+VdwgwAzMzMzMzMzMxVi+yD7GBTD1fAVmYPE0XYi0XcV2YPE0XQM/+LXdSJRfwz9o1H/YP/BA9XwGYPE0X0i1X0D0PwO/cPh9IAAACLTRCLxw8QRdArxg8RRcCNHMGLRfiJRfCJVfhmDx9EAACD/gQPg6MAAAD/cwSLRQz/M/908AT/NPCNRbBQ6G/T//+D7BCLzIPsEA8QAA8QCIvEDxEBDxBFwA8RTeAPEQCNRaBQ6NiN//9mD3PZDA8QEGYPfsgPKMJmD3PYDGYPfsEPEVXAiU38DxFV0DvIdxNyCItF2DtF6HMJuAEAAAAzyesOD1fAZg8TReiLTeyLReiLVfgD0ItF8IlV+BPBRoPrCIlF8Dv3D4ZU////i13U6wOLRfiLTQiLddCJNPmL8YvKi9CJVdyJXP4ER4t12Itd/Il10Ild1IlN2IlV/IP/Bw+C2/7//4tFCF+JcDheiVg8W4vlXcIMAMzMzMzMzMzMVYvsVleLfQhX6JIAAACL8IX2dQZfXl3CBACLVPf4i8qLRPf8M/8LyHQTZg8fRAAAD6zCAUfR6IvKC8h188HmBo1GwAPHX15dwgQAzMzMzMxVi+xW" & _
                                                    "V4t9CFfocgAAAIvwhfZ1Bl9eXcIEAItU9/iLyotE9/wz/wvIdBNmDx9EAAAPrMIBR9Hoi8oLyHXzweYGjUbAA8dfXl3CBADMzMzMzFWL7ItVCLgFAAAADx9EAACLDMILTMIEdQWD6AF58kBdwgQAzMzMzMzMzMzMzMzMzFWL7ItVCLgDAAAADx9EAACLDMILTMIEdQWD6AF58kBdwgQAzMzMzMzMzMzMzMzMzFWL7IPsCItFCA9XwFOL2GYPE0X4g8AwO8N2OItN+FZXi338iU0Ii3D4g+gIi86LUAQPrNEBC00I0eoL14kIi/6JUATB5x/HRQgAAAAAO8N31V9eW4vlXcIEAMzMzMzMzFWL7IPsCItFCA9XwFOL2GYPE0X4g8AgO8N2OItN+FZXi338iU0Ii3D4g+gIi86LUAQPrNEBC00I0eoL14kIi/6JUATB5x/HRQgAAAAAO8N31V9eW4vlXcIEAMzMzMzMzFWL7ItVDItNCIsCiQGLQgSJQQSLQgiJQQiLQgyJQQyLQhCJQRCLQhSJQRSLQhiJQRiLQhyJQRyLQiCJQSCLQiSJQSSLQiiJQSiLQiyJQSxdwggAzMzMzMzMzMzMzMzMzFWL7ItVDItNCIsCiQGLQgSJQQSLQgiJQQiLQgyJQQyLQhCJQRCLQhSJQRSLQhiJQRiLQhyJQRxdwggAzMzMzMxVi+yD7GBTD1fAM8lWZg8TRdiLRdxXZg8TRdCLfdSJTeiJRfAz9o1B+4P5Bg9XwGYPE0X4i138D0PwO/EPhxkBAACLVQyLwQ8QRdArxold9A8RRcCNBMKLVfiJReyJVfyL+Sv+O/cPh+oAAAD/cAT/MItFDP908AT/NPCNRbBQ6KzP//8PEAAP" & _
                                                    "EUXQO/dzQ4tN3IvBi1XUi/rB6B8BRfyLRdiD0wDB7x8PpMEBiV30M9sDwAvZC/iJXdyLRdAPpMIBiX3YA8CJVdSJRdAPEEXQ6waLXdyLfdiD7BCLxIPsEA8RAIvEDxBFwA8RAI1FoFDoy4n//w8QCA8owWYPc9gMZg9+wA8RTcCJRfAPEU3QO8N3EHIFOX3Ycwm4AQAAADPJ6w4PV8BmDxNF4ItN5ItF4ItV/Itd9APQi0XsE9mJVfyLTehGg+gIiV30iUXsO/EPhgr///+LfdTrA4tV+It1CItF0IkEzotF2Il8zgRBi33wiVXYi9OJRdCJfdSJVfCJVdyJTeiD+QsPgpX+//+JflxfiUZYXluL5V3CCADMzFWL7IPsYFMPV8AzyVZmDxNF2ItF3FdmDxNF0It91IlN6IlF8DP2jUH9g/kED1fAZg8TRfiLXfwPQ/A78Q+HGQEAAItVDIvBDxBF0CvGiV30DxFFwI0EwotV+IlF7IlV/Iv5K/479w+H6gAAAP9wBP8wi0UM/3TwBP808I1FsFDoDM7//w8QAA8RRdA793NDi03ci8GLVdSL+sHoHwFF/ItF2IPTAMHvHw+kwQGJXfQz2wPAC9kL+Ild3ItF0A+kwgGJfdgDwIlV1IlF0A8QRdDrBotd3It92IPsEIvEg+wQDxEAi8QPEEXADxEAjUWgUOgriP//DxAIDyjBZg9z2AxmD37ADxFNwIlF8A8RTdA7w3cQcgU5fdhzCbgBAAAAM8nrDg9XwGYPE0Xgi03ki0Xgi1X8i130A9CLRewT2YlV/ItN6EaD6AiJXfSJRew78Q+GCv///4t91OsDi1X4i3UIi0XQiQTOi0XYiXzOBEGLffCJVdiL04lF0Il9" & _
                                                    "1IlV8IlV3IlN6IP5Bw+Clf7//4l+PF+JRjheW4vlXcIIAMzMVYvsg+wMU4tdDA9XwFZXi30QixOL8otDBIvIZg8TRfQrNxtPBDvydQY7yHUE6xg7yHIPdwQ78nYJuAEAAAAz0usLZg8TRfSLRfSLVfiLfQiJTwSJN4tzCIvOiXX4i3UQK04IiU0Ii0sMG04Mi10IK9iJXQgbyjtd+ItdDHUFO0sMdCM7SwxyE3cIi0MIOUUIdgm4AQAAADPS6wtmDxNF9ItV+ItF9It1CIlPDIl3CItzEIvOiXX8i3UQK04QiU0Ii0sUG04Ui10IK9iJXQgbyjtd/ItdDHUFO0sUdCM7SxRyE3cIi0MQOUUIdgm4AQAAADPS6wtmDxNF9ItV+ItF9It1CIlPFIl3EItzGIvOiXX8i3UQK04YiU0Ii0scG04ci10IK9iJXQgbyjtd/ItdDHUFO0scdCM7SxxyE3cIi0MYOUUIdgm4AQAAADPS6wtmDxNF9ItV+ItF9It1CIl3GIt1EIlPHItLICtOIIlNDItLJBtOJIt1DCvwG8o7cyB1BTtLJHQgO0skchB3BTtzIHYJuAEAAAAz0usLZg8TRfSLVfiLRfSJdyCJTySLcyiLSyyLXRCJdQiJTQwrcygbSywr8ItdDBvKO3UIdQQ7y3QsO8tyHXcFO3UIdhaJdyi4AQAAAIlPLDPSX15bi+VdwgwAZg8TRfSLVfiLRfSJdyiJTyxfXluL5V3CDADMzMzMVYvsg+wMU4tdDA9XwFZXi30QixOL8otDBIvIZg8TRfQrNxtPBDvydQY7yHUE6xg7yHIPdwQ78nYJuAEAAAAz0usLZg8TRfSLRfSLVfiLfQiJTwSLTRCJN4tzCIl1+Ctx" & _
                                                    "CItLDItdEBtLDCvwi10MG8o7dfh1BTtLDHQgO0sMchB3BTtzCHYJuAEAAAAz0usLZg8TRfSLVfiLRfSJTwyLTRCJdwiLcxCJdfwrcRCLSxSLXRAbSxQr8ItdDBvKO3X8dQU7SxR0IDtLFHIQdwU7cxB2CbgBAAAAM9LrC2YPE0X0i1X4i0X0iU8UiXcQi0sYi/GLfRCLWxyJTQyLTRArcRiLyxtPHCvwi30IG8o7dQx1BDvLdCw7y3IddwU7dQx2Fol3GLgBAAAAiU8cM9JfXluL5V3CDABmDxNF9ItV+ItF9Il3GIlPHF9eW4vlXcIMAMzMzMzMzMzMzMzMzMzMzFWL7ItNCDPSVleLfQwz9ovHg+A/D6vGg/ggD0PWM/KD+EAPQ9bB7wYjNPkjVPkEi8ZfXl3CCADMzMzMzMzMzMxVi+yLVRSD7BAzyYXSD4TCAAAAU4tdEFaLdQhXi30Mg/ogD4KLAAAAjUP/A8I78HcJjUb/A8I7w3N5jUf/A8I78HcJjUb/A8I7x3Nni8KL1yvTg+DgiVX8i9Yr04lF8IlV+IvDi134i9eLffwr1olV9I1WEA8QAIt19IPBII1AII1SIA8QTAfgZg/vyA8RTAPgDxBMFuCLdQgPEEDwZg/vyA8RSuA7TfByyotVFIt9DItdEDvKcxsr+40EGSvzK9GKDDiNQAEySP+ITDD/g+oBde5fXluL5V3CEAAAAA==" ' 35581, 24.4.2020 12:48:15
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
            .Thunk = pvThunkAllocate(STR_THUNK1 & STR_THUNK2 & STR_THUNK3 & STR_THUNK4)
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
            Call pvPatchTrampoline(AddressOf pvCryptoCallSecpSign)
            Call pvPatchTrampoline(AddressOf pvCryptoCallSecpVerify)
            Call pvPatchTrampoline(AddressOf pvCryptoCallCurve25519Multiply)
            Call pvPatchTrampoline(AddressOf pvCryptoCallCurve25519MulBase)
            Call pvPatchTrampoline(AddressOf pvCryptoCallSha2Init)
            Call pvPatchTrampoline(AddressOf pvCryptoCallSha2Update)
            Call pvPatchTrampoline(AddressOf pvCryptoCallSha2Final)
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

Public Function CryptoEccCurve25519MakeKey(baPrivate() As Byte, baPublic() As Byte) As Boolean
    ReDim baPrivate(0 To m_uData.Ecc256KeySize - 1) As Byte
    ReDim baPublic(0 To m_uData.Ecc256KeySize - 1) As Byte
    CryptoRandomBytes VarPtr(baPrivate(0)), m_uData.Ecc256KeySize
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
    
    Debug.Assert UBound(baPrivate) >= m_uData.Ecc256KeySize - 1
    Debug.Assert UBound(baPublic) >= m_uData.Ecc256KeySize - 1
    ReDim baRetVal(0 To m_uData.Ecc256KeySize - 1) As Byte
    #If ImplUseLibSodium Then
        Call crypto_scalarmult_curve25519(baRetVal(0), baPrivate(0), baPublic(0))
    #Else
        pvCryptoCallCurve25519Multiply m_uData.Pfn(ucsPfnCurve25519ScalarMultiply), baRetVal(0), baPrivate(0), baPublic(0)
    #End If
    CryptoEccCurve25519SharedSecret = baRetVal
End Function

Public Function CryptoEccSecp256r1MakeKey(baPrivate() As Byte, baPublic() As Byte) As Boolean
    Const MAX_RETRIES   As Long = 16
    Dim lIdx            As Long
    
    ReDim baPrivate(0 To m_uData.Ecc256KeySize - 1) As Byte
    ReDim baPublic(0 To m_uData.Ecc256KeySize) As Byte
    For lIdx = 1 To MAX_RETRIES
        CryptoRandomBytes VarPtr(baPrivate(0)), m_uData.Ecc256KeySize
        If pvCryptoCallSecp256r1MakeKey(m_uData.Pfn(ucsPfnSecp256r1MakeKey), baPublic(0), baPrivate(0)) = 1 Then
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
    If pvCryptoCallSecp256r1SharedSecret(m_uData.Pfn(ucsPfnSecp256r1SharedSecret), baPublic(0), baPrivate(0), baRetVal(0)) = 0 Then
        GoTo QH
    End If
    CryptoEccSecp256r1SharedSecret = baRetVal
QH:
End Function

Public Function CryptoEccSecp256r1UncompressKey(baPublic() As Byte) As Byte()
    Dim baRetVal()      As Byte
    
    ReDim baRetVal(0 To 2 * m_uData.Ecc256KeySize) As Byte
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
    
    ReDim baRandom(0 To m_uData.Ecc256KeySize - 1) As Byte
    ReDim baRetVal(0 To 2 * m_uData.Ecc256KeySize - 1) As Byte
    For lIdx = 1 To MAX_RETRIES
        CryptoRandomBytes VarPtr(baRandom(0)), m_uData.Ecc256KeySize
        If pvCryptoCallSecpSign(m_uData.Pfn(ucsPfnSecp256r1Sign), baPrivate(0), baHash(0), baRandom(0), baRetVal(0)) <> 0 Then
            Exit For
        End If
    Next
    If lIdx < MAX_RETRIES Then
        '--- success
        CryptoEccSecp256r1Sign = baRetVal
    End If
End Function

Public Function CryptoEccSecp256r1Verify(baPublic() As Byte, baHash() As Byte, baSignature() As Byte) As Boolean
'    MsgBox "ucsPfnSecp256r1Verify = 0x" & Hex$(m_uData.Pfn(ucsPfnSecp256r1Verify))
    CryptoEccSecp256r1Verify = (pvCryptoCallSecpVerify(m_uData.Pfn(ucsPfnSecp256r1Verify), baPublic(0), baHash(0), baSignature(0)) <> 0)
End Function

Public Function CryptoEccSecp384r1Sign(baPrivate() As Byte, baHash() As Byte) As Byte()
    Const MAX_RETRIES   As Long = 16
    Dim baRandom()      As Byte
    Dim baRetVal()      As Byte
    Dim lIdx            As Long
    
    ReDim baRandom(0 To m_uData.Ecc384KeySize - 1) As Byte
    ReDim baRetVal(0 To 2 * m_uData.Ecc384KeySize - 1) As Byte
    For lIdx = 1 To MAX_RETRIES
        CryptoRandomBytes VarPtr(baRandom(0)), m_uData.Ecc384KeySize
        If pvCryptoCallSecpSign(m_uData.Pfn(ucsPfnSecp384r1Sign), baPrivate(0), baHash(0), baRandom(0), baRetVal(0)) <> 0 Then
            Exit For
        End If
    Next
    If lIdx < MAX_RETRIES Then
        '--- success
        CryptoEccSecp384r1Sign = baRetVal
    End If
End Function

Public Function CryptoEccSecp384r1Verify(baPublic() As Byte, baHash() As Byte, baSignature() As Byte) As Boolean
    CryptoEccSecp384r1Verify = (pvCryptoCallSecpVerify(m_uData.Pfn(ucsPfnSecp384r1Verify), baPublic(0), baHash(0), baSignature(0)) <> 0)
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
            pvCryptoCallSha2Init .Pfn(ucsPfnSha256Init), lCtxPtr
            pvCryptoCallSha2Update .Pfn(ucsPfnSha256Update), lCtxPtr, lPtr, Size
            pvCryptoCallSha2Final .Pfn(ucsPfnSha256Final), lCtxPtr, baRetVal(0)
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
            pvCryptoCallSha2Init .Pfn(ucsPfnSha384Init), lCtxPtr
            pvCryptoCallSha2Update .Pfn(ucsPfnSha384Update), lCtxPtr, lPtr, Size
            pvCryptoCallSha2Final .Pfn(ucsPfnSha384Final), lCtxPtr, baRetVal(0)
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
    With m_uData
        lCtxPtr = VarPtr(.HashCtx(0))
        #If ImplUseLibSodium Then
            Call crypto_hash_sha512_init(ByVal lCtxPtr)
            Call crypto_hash_sha512_update(ByVal lCtxPtr, ByVal lPtr, Size)
            Call crypto_hash_sha512_final(ByVal lCtxPtr, .HashFinal(0))
            Call CopyMemory(baRetVal(0), .HashFinal(0), LNG_SHA512_HASHSZ)
        #Else
            pvCryptoCallSha2Init .Pfn(ucsPfnSha512Init), lCtxPtr
            pvCryptoCallSha2Update .Pfn(ucsPfnSha512Update), lCtxPtr, lPtr, Size
            pvCryptoCallSha2Final .Pfn(ucsPfnSha512Final), lCtxPtr, baRetVal(0)
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
            pvCryptoCallSha2Init .Pfn(ucsPfnSha256Init), lCtxPtr
            Call FillMemory(.HashPad(0), LNG_SHA256_BLOCKSZ, LNG_HMAC_INNER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_INNER_PAD
            Next
            pvCryptoCallSha2Update .Pfn(ucsPfnSha256Update), lCtxPtr, VarPtr(.HashPad(0)), LNG_SHA256_BLOCKSZ
            pvCryptoCallSha2Update .Pfn(ucsPfnSha256Update), lCtxPtr, lPtr, Size
            pvCryptoCallSha2Final .Pfn(ucsPfnSha256Final), lCtxPtr, .HashFinal(0)
            '-- outer hash
            pvCryptoCallSha2Init .Pfn(ucsPfnSha256Init), lCtxPtr
            Call FillMemory(.HashPad(0), LNG_SHA256_BLOCKSZ, LNG_HMAC_OUTER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_OUTER_PAD
            Next
            pvCryptoCallSha2Update .Pfn(ucsPfnSha256Update), lCtxPtr, VarPtr(.HashPad(0)), LNG_SHA256_BLOCKSZ
            pvCryptoCallSha2Update .Pfn(ucsPfnSha256Update), lCtxPtr, VarPtr(.HashFinal(0)), LNG_SHA256_HASHSZ
            pvCryptoCallSha2Final .Pfn(ucsPfnSha256Final), lCtxPtr, baRetVal(0)
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
            pvCryptoCallSha2Init .Pfn(ucsPfnSha384Init), lCtxPtr
            Call FillMemory(.HashPad(0), LNG_SHA384_BLOCKSZ, LNG_HMAC_INNER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_INNER_PAD
            Next
            pvCryptoCallSha2Update .Pfn(ucsPfnSha384Update), lCtxPtr, VarPtr(.HashPad(0)), LNG_SHA384_BLOCKSZ
            pvCryptoCallSha2Update .Pfn(ucsPfnSha384Update), lCtxPtr, lPtr, Size
            pvCryptoCallSha2Final .Pfn(ucsPfnSha384Final), lCtxPtr, .HashFinal(0)
            '-- outer hash
            pvCryptoCallSha2Init .Pfn(ucsPfnSha384Init), lCtxPtr
            Call FillMemory(.HashPad(0), LNG_SHA384_BLOCKSZ, LNG_HMAC_OUTER_PAD)
            For lIdx = 0 To UBound(baKey)
                .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_OUTER_PAD
            Next
            pvCryptoCallSha2Update .Pfn(ucsPfnSha384Update), lCtxPtr, VarPtr(.HashPad(0)), LNG_SHA384_BLOCKSZ
            pvCryptoCallSha2Update .Pfn(ucsPfnSha384Update), lCtxPtr, VarPtr(.HashFinal(0)), LNG_SHA384_HASHSZ
            pvCryptoCallSha2Final .Pfn(ucsPfnSha384Final), lCtxPtr, baRetVal(0)
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
    
    Select Case SignatureType \ &H100
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

Public Function CryptoRsaSign(uCtx As UcsRsaContextType, baMessage() As Byte) As Byte()
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

Public Function CryptoRsaVerify(uCtx As UcsRsaContextType, baMessage() As Byte, baSignature() As Byte) As Boolean
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
        Err.Raise IIf(hResult < 0, hResult, hResult Or LNG_FACILITY_WIN32), sApiSource
    End If
End Function

Public Function CryptoExtractPublicKey(baCert() As Byte, baPubKey() As Byte, Optional sObjId As String) As Boolean
    Dim pContext        As Long
    Dim lPtr            As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    Dim uInfo           As CERT_PUBLIC_KEY_INFO

    pContext = CertCreateCertificateContext(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, baCert(0), UBound(baCert) + 1)
    If pContext = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CertCreateCertificateContext"
        GoTo QH
    End If
    Call CopyMemory(lPtr, ByVal UnsignedAdd(pContext, 12), 4)       '--- dereference pContext->pCertInfo
    lPtr = UnsignedAdd(lPtr, 56)                                    '--- &pContext->pCertInfo->SubjectPublicKeyInfo
    Call CopyMemory(uInfo, ByVal lPtr, Len(uInfo))
    sObjId = String$(lstrlen(uInfo.AlgObjIdPtr), 0)
    Call CopyMemory(ByVal sObjId, ByVal uInfo.AlgObjIdPtr, Len(sObjId))
    ReDim baPubKey(0 To uInfo.PubKeySize - 1) As Byte
    Call CopyMemory(baPubKey(0), ByVal uInfo.PubKeyPtr, uInfo.PubKeySize)
    '--- success
    CryptoExtractPublicKey = True
QH:
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
    lAlignedSize = (lSize + MAX_RSA_BYTES - 1 And -MAX_RSA_BYTES) + MAX_RSA_BYTES
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

Public Function CryptoRsaDecrypt(ByVal hPrivKey As Long, baCipherText() As Byte) As Byte()
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    Dim hResult         As Long
    Dim sApiSource      As String
    
    baRetVal = baCipherText
    pvArrayReverse baRetVal
    lSize = pvArraySize(baRetVal)
    If CryptDecrypt(hPrivKey, 0, 1, 0, baRetVal(0), lSize) = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CryptDecrypt"
        GoTo QH
    End If
    If UBound(baRetVal) <> lSize - 1 Then
        ReDim Preserve baRetVal(0 To lSize - 1) As Byte
    End If
    CryptoRsaDecrypt = baRetVal
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
    
    If OsVersion < ucsOsvVista Then
        GoTo QH
    End If
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

Public Function CryptoRsaPssVerify(baCert() As Byte, baMessage() As Byte, baSignature() As Byte, ByVal lSignatureType As Long) As Boolean
    Dim pContext        As Long
    Dim lPtr            As Long
    Dim hKey            As Long
    Dim uPadInfo        As BCRYPT_PSS_PADDING_INFO
    Dim hResult         As Long
    Dim sApiSource      As String
    
    If OsVersion < ucsOsvVista Then
        GoTo QH
    End If
    pContext = CertCreateCertificateContext(X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING, baCert(0), UBound(baCert) + 1)
    If pContext = 0 Then
        hResult = Err.LastDllError
        sApiSource = "CertCreateCertificateContext"
        GoTo QH
    End If
    Call CopyMemory(lPtr, ByVal UnsignedAdd(pContext, 12), 4)       '--- dereference pContext->pCertInfo
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
    If pContext <> 0 Then
        Call CertFreeCertificateContext(pContext)
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

'= trampolines ===========================================================

Private Function pvCryptoCallCurve25519Multiply(ByVal Pfn As Long, pSecretPtr As Byte, pPubKeyPtr As Byte, pPrivKeyPtr As Byte) As Long
    ' void cf_curve25519_mul(uint8_t out[32], const uint8_t priv[32], const uint8_t pub[32])
End Function

Private Function pvCryptoCallCurve25519MulBase(ByVal Pfn As Long, pPubKeyPtr As Byte, pPrivKeyPtr As Byte) As Long
    ' void cf_curve25519_mul_base(uint8_t out[32], const uint8_t priv[32])
End Function

Private Function pvCryptoCallSecp256r1MakeKey(ByVal Pfn As Long, pPubKeyPtr As Byte, pPrivKeyPtr As Byte) As Long
    ' int ecc_make_key(uint8_t p_publicKey[ECC_BYTES+1], uint8_t p_privateKey[ECC_BYTES]);
End Function

Private Function pvCryptoCallSecp256r1SharedSecret(ByVal Pfn As Long, pPubKeyPtr As Byte, pPrivKeyPtr As Byte, pSecretPtr As Byte) As Long
    ' int ecdh_shared_secret(const uint8_t p_publicKey[ECC_BYTES+1], const uint8_t p_privateKey[ECC_BYTES], uint8_t p_secret[ECC_BYTES]);
End Function

Private Function pvCryptoCallSecp256r1UncompressKey(ByVal Pfn As Long, pPubKeyPtr As Byte, pUncompressedKeyPtr As Byte) As Long
    ' int ecdh_uncompress_key(const uint8_t p_publicKey[ECC_BYTES + 1], uint8_t p_uncompressedKey[2 * ECC_BYTES + 1])
End Function

Private Function pvCryptoCallSecpSign(ByVal Pfn As Long, pPrivKeyPtr As Byte, pHashPtr As Byte, pRandomPtr As Byte, pSignaturePtr As Byte) As Long
    ' int ecdsa_sign(const uint8_t p_privateKey[ECC_BYTES], const uint8_t p_hash[ECC_BYTES], uint64_t k[NUM_ECC_DIGITS], uint8_t p_signature[ECC_BYTES*2])
    ' int ecdsa_sign384(const uint8_t p_privateKey[ECC_BYTES_384], const uint8_t p_hash[ECC_BYTES_384], uint64_t k[NUM_ECC_DIGITS_384], uint8_t p_signature[ECC_BYTES_384*2])
End Function

Private Function pvCryptoCallSecpVerify(ByVal Pfn As Long, pPubKeyPtr As Byte, pHashPtr As Byte, pSignaturePtr As Byte) As Long
    ' int ecdsa_verify(const uint8_t p_publicKey[ECC_BYTES+1], const uint8_t p_hash[ECC_BYTES], const uint8_t p_signature[ECC_BYTES*2])
    ' int ecdsa_verify384(const uint8_t p_publicKey[ECC_BYTES_384+1], const uint8_t p_hash[ECC_BYTES_384], const uint8_t p_signature[ECC_BYTES_384*2])
End Function

Private Function pvCryptoCallSha2Init(ByVal Pfn As Long, ByVal lCtxPtr As Long) As Long
    ' void cf_sha256_init(cf_sha256_context *ctx)
    ' void cf_sha384_init(cf_sha384_context *ctx)
    ' void cf_sha512_init(cf_sha512_context *ctx)
End Function

Private Function pvCryptoCallSha2Update(ByVal Pfn As Long, ByVal lCtxPtr As Long, ByVal lDataPtr As Long, ByVal lSize As Long) As Long
    ' void cf_sha256_update(cf_sha256_context *ctx, const void *data, size_t nbytes)
    ' void cf_sha384_update(cf_sha384_context *ctx, const void *data, size_t nbytes)
    ' void cf_sha512_update(cf_sha512_context *ctx, const void *data, size_t nbytes)
End Function

Private Function pvCryptoCallSha2Final(ByVal Pfn As Long, ByVal lCtxPtr As Long, pHashPtr As Byte) As Long
    ' void cf_sha256_digest_final(cf_sha256_context *ctx, uint8_t hash[LNG_SHA256_HASHSZ])
    ' void cf_sha384_digest_final(cf_sha384_context *ctx, uint8_t hash[LNG_SHA384_HASHSZ])
    ' void cf_sha512_digest_final(cf_sha512_context *ctx, uint8_t hash[LNG_SHA384_HASHSZ])
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
