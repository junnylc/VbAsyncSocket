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
                                                    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAPwAAABjfHd78mtvxTABZyv+16t2yoLJffpZR/Ct1KKvnKRywLf9kyY2P/fMNKXl8XHYMRUExyPDGJYFmgcSgOLrJ7J1CYMsGhtuWqBSO9azKeMvhFPRAO0g/LFbasu+OUpMWM/Q76r7Q00zhUX5An9QPJ+oUaNAj5KdOPW8ttohEP/z0s0ME+xfl0QXxKd+PWRdGXNggU/cIiqQiEbuuBTeXgvb4DI6CkkGJFzC06xikZXkeefIN22N1U6pbFb06mV6rgi6eCUuHKa0xujddB9LvYuKcD61ZkgD9g5hNVe5hsEdnuH4mBFp2Y6Umx6H6c5VKN+MoYkNv+ZCaEGZLQ+wVLsWjQECBAgQIECAGzZSCWrVMDalOL9Ao56B89f7fOM5gpsv/4c0jkNExN7py1R7lDKmwiM97kyVC0L6w04ILqFmKNkksnZboklti9Elcvj2ZIZomBbUpFzMXWW2kmxwSFD97bnaXhVGV6eNnYSQ2KsAjLzTCvfkWAW4s0UG0Cwej8o/DwLBr70DAROKazqREUFPZ9zql/LPzvC05nOWrHQi5601heL5N+gcdd9uR/EacR0pxYlvt2IOqhi+G/xWPkvG0nkgmtvA/njNWvQf3agziAfHMbESEFkngOxfYFF/qRm1Sg0t5Xqfk8mc76DgO02uKvWwyOu7PINTmWEXKwR+unfWJuFpFGNVIQx9AAAAAAA=" ' 1928, 24.4.2020 15:03:56
Private Const STR_THUNK1                As String = "IALYAOAgAAAAJAAAQDkAAPA/AACQQAAAIEIAAFBHAACQOAAAgD8AAFBAAADQQAAAUEMAAIAtAADQLQAAsCsAAEAuAADQLgAAEC4AAGAyAADwMgAA4C4AAAAgAADAHwAAEBUAAJAUAADMzMzMzMzMzOgAAAAAWC11QNcABQBA1wCLAMPMzMzMzMzMzMzMzMzM6AAAAABYLZVA1wAFAEDXAMPMzMzMzMzMzMzMzMzMzMxVi+yD7GhTi10QU+hAbAAAhcAPhVsBAABWi3UMjUXIV1ZQ6Il8AACLfQiNRchQV41FmFDouHoAAI1FyFBQ6G58AABTVlbopnoAAFNT6F98AADoav///wWgAAAAUFNXV+jscwAA6Ff///8FoAAAAFBTU1Po2XMAAOhE////BaAAAABQU1dT6IZ8AABTV1foXnoAAOgp////BaAAAABQV1dT6KtzAADoFv///wWgAAAAUFNXV+iYcwAAagBX6ICIAAALwnQl6Pf+//8FoAAAAFBXV+hqZgAAV4vw6FKAAADB5h8JdyyLdQzrBlfoQYAAAFdT6Lp7AADoxf7//wWgAAAAUI1FmFBTU+gEfAAA6K/+//8FoAAAAFCNRZhQU1Po7nsAAOiZ/v//BaAAAABQU41FmFBQ6Nh7AACNRZhQV1forXkAAOh4/v//BaAAAABQjUXIUFdQ6Ld7AABTV+iQgAAAVlPoiYAAAI1FyFBW6H+AAABfXluL5V3CDADMzMzMzMxVi+yD7EhTi10QU+jwagAAhcAPhSkBAABWi3UMjUXYV1ZQ6Dl7AACLfQiNRdhQV41FuFDoyHoAAI1F2FBQ6B57AABTVlbotnoAAFNT6A97AADo6v3//1BTV1fosXIAAOjc/f//UFNTU+ijcgAA6M79//9Q" & _
                                                    "U1dT6EV7AABTV1fofXoAAOi4/f//UFdXU+h/cgAA6Kr9//9QU1dX6HFyAABqAFfoGYcAAAvCdCDokP3//1BXV+hIZwAAV4vw6FB/AADB5h8JdxyLdQzrBlfoP38AAFdT6Ih6AADoY/3//1CNRbhQU1Po13oAAOhS/f//UI1FuFBTU+jGegAA6EH9//9QU41FuFBQ6LV6AACNRbhQV1fo6nkAAOgl/f//UI1F2FBXUOiZegAAU1foon8AAFZT6Jt/AACNRdhQVuiRfwAAX15bi+VdwgwAzMzMzMzMzMxVi+xWi3UIVuhzaQAAhcB0F41GMFDoZmkAAIXAdAq4AQAAAF5dwgQAM8BeXcIEAMxVi+xWi3UIVuhzaQAAhcB0F41GIFDoZmkAAIXAdAq4AQAAAF5dwgQAM8BeXcIEAMxVi+yB7PgAAABTi10MjUWYVldTUOinfgAAjUMwUIlF+I2FOP///1DolH4AAP91FI2FCP///1CNhWj///9QjYU4////UI1FmFDowwcAAItdEFPoqnwAAI1w/oX2fmAPHwBWU+iphQAAC8J1B7gBAAAA6wIzwI0EQMHgBI2NCP///wPIjZVo////A9CJTRRR99iJVfyNvTj///9SA/iNXZgD2FdT6IgEAABXU/91FP91/OibAgAAi10QToX2f6NqAFPoS4UAAAvCdQe4AQAAAOsCM8CNBEDB4ASNnQj///8D2I2NaP///1MDyI29OP///1Er+IlNEI11mCvwV1boLwQAAOiK+///BaAAAABQjYVo////UI1FmFCNRchQ6MB4AABXjUXIUFDolXYAAP91DI1FyFBQ6Ih2AADoU/v//wWgAAAAUI1FyFBQ6FNwAAD/dfiNRchQUOhm" & _
                                                    "dgAAVo1FyFBQ6Ft2AABXVlP/dRDo4AEAAI1FyFCNhQj///9QjYVo////UOhJCgAAi3UIjYVo////UFboKX0AAI2FCP///1CNRjBQ6Bl9AABfXluL5V3CEABVi+yB7KgAAABTi10MjUW4VldTUOhXfQAAjUMgUIlF+I2FeP///1DoRH0AAP91FI2FWP///1CNRZhQjYV4////UI1FuFDolgYAAItdEFPoTXsAAIPoAolFFIXAflsPHwBQU+j5gwAAC8J1B7gBAAAA6wIzwMHgBY2dWP///wPYjU2YA8iNtXj///9T99iJTfxRA/CNfbgD+FZX6HEEAABWV1P/dfzo9gEAAItFFItdEEiJRRSFwH+oagBT6KCDAAALwnUFjUgB6wIzycHhBY2dWP///wPZiU0QU41FmAPBjb14////UCv5jXW4K/FXVugcBAAA6Of5//9QjUWYUI1FuFCNRdhQ6FV3AABXjUXYUFDoinYAAP91DI1F2FBQ6H12AADouPn//1CNRdhQUOgdcQAA/3X4jUXYUFDoYHYAAFaNRdhQUOhVdgAAV1aNRZgDRRBTUOhGAQAAjUXYUI2FWP///1CNRZhQ6AIJAACLdQiNRZhQVuj1ewAAjYVY////UI1GIFDo5XsAAF9eW4vlXcIQAMzMzMzMzMzMzMzMzFWL7IPsMFNWV+gy+f//i10IBaAAAACLdRBQU1aNRdBQ6Gt2AACNRdBQUOgBdgAAjUXQUFNT6DZ0AACNRdBQVlboK3QAAOj2+P//i3UMBaAAAACLfRRQVldX6DJ2AABXjUXQUOjIdQAA6NP4//8FoAAAAFBTjUXQUFDoEnYAAOi9+P//BaAAAABQi0UQUI1F0FBQ6Pl1AADopPj/" & _
                                                    "/wWgAAAAUItFEFNQUOjjdQAAi0UQUFZW6LhzAADog/j//wWgAAAAUI1F0FBTi10QU+i/dQAAU1dX6JdzAADoYvj//wWgAAAAUFZXV+ikdQAAjUXQUFPoenoAAF9eW4vlXcIQAMxVi+yD7CBTVlfoMvj//4tdCIt1EFBTVo1F4FDooHUAAI1F4FBQ6DZ1AACNReBQU1Poy3QAAI1F4FBWVujAdAAA6Pv3//+LdQyLfRRQVldX6Gx1AABXjUXgUOgCdQAA6N33//9QU41F4FBQ6FF1AADozPf//1CLRRBQjUXgUFDoPXUAAOi49///UItFEFNQUOgsdQAAi0UQUFZW6GF0AADonPf//1CNReBQU4tdEFPoDXUAAFNXV+hFdAAA6ID3//9QVldX6Pd0AACNReBQU+j9eQAAX15bi+VdwhAAzMzMzFWL7IHskAAAAFNWV+hP9///i10IBaAAAACLfRBQU1eNRaBQ6Ih0AACNRaBQUOgedAAAjUWgUFNT6FNyAACNRaBQV1foSHIAAOgT9///i10MBaAAAACLdRRQU1aNRaBQ6IxrAADo9/b//wWgAAAAUFNWVug5dAAA6OT2//8FoAAAAFD/dQiNRdBXUOghdAAAjUXQUFNT6PZxAADowfb//wWgAAAAUFf/dQiNRdBQ6D5rAABWV+iXcwAA6KL2//8FoAAAAFCNRdBQV1fo4XMAAOiM9v//BaAAAABQV4t9CI2FcP///1dQ6MVzAACNhXD///9QVlbol3EAAOhi9v//BaAAAABQU1ZW6KRzAACNRaBQjYVw////UOg0cwAA6D/2//8FoAAAAFCNRdBQjYVw////UFDoeHMAAOgj9v//BaAAAABQV42FcP///1CNRdBQ" & _
                                                    "6FxzAACNRaBQjUXQUFDoLnEAAOj59f//BaAAAABQU41F0FBT6DhzAACNhXD///9QV+gLeAAAX15bi+VdwhAAzMxVi+yD7GBTVlfowvX//4tdCIt9EFBTV41FwFDoMHMAAI1FwFBQ6MZyAACNRcBQU1PoW3IAAI1FwFBXV+hQcgAA6Iv1//+LXQyLdRRQU1aNRcBQ6ElqAADodPX//1BTVlbo63IAAOhm9f//UP91CI1F4FdQ6NhyAACNReBQU1PoDXIAAOhI9f//UFf/dQiNReBQ6ApqAABWV+hTcgAA6C71//9QjUXgUFdX6KJyAADoHfX//1BXi30IjUWgV1DojnIAAI1FoFBWVujDcQAA6P70//9QU1ZW6HVyAACNRcBQjUWgUOgIcgAA6OP0//9QjUXgUI1FoFBQ6FRyAADoz/T//1BXjUWgUI1F4FDoQHIAAI1FwFCNReBQUOhycQAA6K30//9QU41F4FBT6CFyAACNRaBQV+gndwAAX15bi+VdwhAAzMzMzMzMzMzMzMzMzMxVi+yD7DBWi3UIV1b/dRDonHYAAIt9DFf/dRTokHYAAI1F0FDoV18AAItFGMdF0AEAAADHRdQAAAAAhcB0ClCNRdBQ6Gh2AACNRdBQV1bobQMAAI1F0FBXVuhi9P//jUXQUP91FP91EOhTAwAAX16L5V3CFADMzMzMzMzMzMzMzFWL7IPsIFaLdQhXVv91EOh8dgAAi30MV/91FOhwdgAAjUXgUOg3XwAAi0UYx0XgAQAAAMdF5AAAAACFwHQKUI1F4FDoSHYAAI1F4FBXVug9AwAAjUXgUFdW6GL1//+NReBQ/3UU/3UQ6CMDAABfXovlXcIUAMzMzMzMzMzMzMzMU4tE" & _
                                                    "JAyLTCQQ9+GL2ItEJAj3ZCQUA9iLRCQI9+ED01vCEADMzMzMzMzMzMzMzMzMgPlAcxWA+SBzBg+lwtPgw4vQM8CA4R/T4sMzwDPSw8yA+UBzFYD5IHMGD63Q0+rDi8Iz0oDhH9PowzPAM9LDzFWL7ItFEFNWi3UIjUh4V4t9DI1WeDvxdwQ70HMLjU94O/F3MDvXciwr+LsQAAAAK/CLFDgDEItMOAQTSASNQAiJVDD4iUww/IPrAXXkX15bXcIMAIvXjUgQi94r0CvYK/64BAAAAI12II1JIA8QQdAPEEw34GYP1MgPEU7gDxBMCuAPEEHgZg/UyA8RTAvgg+gBddJfXltdwgwAzMzMzMxVi+yLVRyD7AiLRSBWi3UIV4t9DAPXE0UQiRaJRgQ7RRB3D3IEO9dzCbgBAAAAM8nrDg9XwGYPE0X4i038i0X4A0UkXxNNKANFFIlGCIvGE00YiU4MXovlXcIkAMzMzMxVi+yLVQyLTQiLAjEBi0IEMUEEi0IIMUEIi0IMMUEMXcIIAMzMzMzMzMzMzMzMzMxVi+yD7AiLTQiLVRBTVosBjVkEweoCM/aJVRCJXfiNBIUEAAAAiUX8V4XSdEKLVQyLfRCDwgJmZg8fhAAAAAAAD7ZK/o1SBA+2QvvB4QgLyA+2QvzB4QgLyA+2Qv3B4QgLyIkMs0Y793LWi0X8i9e5AQAAADP/iU0MO/APg40AAACLxivCjQSDiUUIDx9EAACLXLP8O/p1CEEz/4lNDOsEhf91Leg38f//BXgFAADBwwhQU+iIVwAAi9joIfH//4tNDA+2hAh4BgAAweAYM9jrHYP6BnYeg/8EdRnoAPH//wV4BQAAUFPoVFcAAIvYi0UIi1UQiwhH" & _
                                                    "M8uDwASLXfiJRQiJDLNGi00MO3X8coJfXluL5V3CDADMzMzMzMzMzMxVi+yD7DCNRdD/dRBQ6J5tAACNRdBQi0UIUFDo0GsAAP91EI1F0FBQ6MNrAACNRdBQi0UMUFDotWsAAIvlXcIMAMzMzMzMzMzMzMzMzMzMzFWL7IPsII1F4P91EFDofm0AAI1F4FCLRQhQUOgQbQAA/3UQjUXgUFDoA20AAI1F4FCLRQxQUOj1bAAAi+VdwgwAzMzMzMzMzMzMzMzMzMzMVYvsg+wgU1aLdQgzyVeJTeyBBM4AAAEAiwTOg1TOBACLXM4ED6zYEMH7EIlF6IP5D3UVx0X8AQAAAIvQx0XwAAAAAIld+OsiD1fAZg8TRfSLRfiJRfCLRfRmDxNF4ItV4IlF/ItF5IlF+IP5D415AWoAG8D32A+vxytV/GoljTTGi0X4G0XwUFLoEvz//4tN6APBE9OD6AGD2gABBotF7BFWBIt1CA+kyxDB4RApDMaLz4lN7BlcxgSD+RAPgk////9fXluL5V3CBADMzMzMzFWL7IPsEItVDFZXD7YKD7ZCAcHhCAvID7ZCAsHhCAvID7ZCA8HhCAvID7ZCBYlN8A+2SgTB4QgLyA+2QgbB4QgLyA+2QgfB4QgLyA+2QgmJTfQPtkoIweEIC8gPtkIKweEIC8gPtkILweEIC8gPtkIMiU34D7ZKDcHgCAvID7ZCDsHhCAvID7ZCD8HhCAvIiU38i00IizmNcQSLx8HgBAPwjUXwVlDolfz//4PuEIPH/3QtjUXwUOikPAAAjUXwUOg7PQAAVo1F8FDocfz//41F8FDoSDwAAIPuEIPvAXXTjUXwUOh3PAAAjUXwUOgOPQAAVo1F8FDoRPz/" & _
                                                    "/4t1EItV8IvCi030wegYiAaLwsHoEIhGAYvCwegIiEYCi8HB6BiIVgOIRgSLwcHoEIhGBYvBwegIiEYGiE4Hi034i8HB6BiIRgiLwcHoEIhGCYvBwegIiEYKiE4Li038i8HB6BiIRgyLwcHoEIhGDYvBwegIiEYOX4hOD16L5V3CDADMzFWL7IPsEFNWV4tVDItdCA+2Cg+2QgHB4QiNcwQLyA+2QgLB4QgLyA+2QgPB4QgLyA+2QgWJTfAPtkoEweEIC8gPtkIGweEIC8gPtkIHweEIC8gPtkIJiU30D7ZKCMHhCAvID7ZCCsHhCAvID7ZCC8HhCAvID7ZCDIlN+A+2Sg3B4AgLyA+2Qg7B4QgLyA+2Qg/B4QgLyI1F8FZQiU386B37//+/AQAAAIPGEDk7di6QjUXwUOgHUwAAjUXwUOieUQAAjUXwUOi1PAAAVo1F8FDo6/r//0eDxhA7O3LTjUXwUOjaUgAAjUXwUOhxUQAAVo1F8FDox/r//4t1EItV8IvCi030wegYiAaLwsHoEIhGAYvCwegIiEYCi8HB6BiIVgOIRgSLwcHoEIhGBYvBwegIiEYGiE4Hi034i8HB6BiIRgiLwcHoEIhGCYvBwegIiEYKiE4Li038i8HB6BiIRgyLwcHoEIhGDYvBwegIiEYOX4hOD15bi+VdwgwAzMzMzFWL7FaLdQho9AAAAGoAVuiMOwAAi0UQg8QMg/gQdDWD+Bh0GoP4IHU8UP91DMcGDgAAAFboN/r//15dwgwAahj/dQzHBgwAAABW6CH6//9eXcIMAGoQ/3UMxwYKAAAAVugL+v//Xl3CDADMzMzMzMxVi+yB7AABAABW6PHr//++oFLXAIHuAEDXAAPw6N/r" & _
                                                    "////dSi5IFHXAMdF9BAAAAD/dSSB6QBA1wCJdfgDwYlF/I2FAP///1DoQ/////91CI2FAP///2oQ/3UUagz/dSD/dRz/dRj/dRD/dQxQjUX0UOg6DwAAXovlXcIkAMzMzFWL7IHsAAEAAFbocev//76gUtcAge4AQNcAA/DoX+v///91KLkgUdcAx0X0EAAAAP91JIHpAEDXAIl1+APBiUX8jYUA////UOjD/v//ahD/dQyNhQD/////dQhqDP91IP91HP91GP91FP91EFCNRfRQ6HoQAABei+VdwiQAzMzMVYvsUVOLXRgzwIlF/IXbdHGLVRCLTQxWx0UYAQAAAFeLOYvyK/c73g9C84XAdR0PtkUUVlCLRQgDx1Do8DkAAItNDIPEDItF/ItVEIX/dQk78g9ERRiJRfyNBD47wnUX/3UI/3Ug/1Uci00Mi1UQxwEAAAAA6wIBMYtF/CvedaBfXluL5V3CHADMzMzMzMzMVYvsVot1IIvGg+gAdGCD6AEPhKwAAABTg+gBV41FFHRti30oi10kV1NqAVD/dRD/dQz/dQjotgAAAItNGFdTOE0cdC+NRv6LdRBQUVb/dQz/dQjoGP///1dTagGNRRxQVv91DP91COiEAAAAX1teXcIkAI1G/4t1EFBRVv91DP91COjp/v//X1teXcIkAP91KItdEP91JIt9DIt1CGoBUFNXVuhIAAAA/3UojUUc/3UkagFQU1dW6DQAAABfW15dwiQA/3UoikUc/3UkMEUUjUUUagFQ/3UQ/3UM/3UI6A0AAABeXcIkAMzMzMzMzMzMVYvs/3Ugi0UcUFD/dRj/dRT/dRD/dQz/dQjoEQAAAF3CHADMzMzMzMzMzMzMzMzMVYvs" & _
                                                    "i00Mi0UkU4tdFIsRVot1GFeF0nRZhfZ0VYtFEIv+K8I7xg9C+IvCA0UIV1NQ6Bs4AACLRQwD3yv3g8QMATiLfRA5OItFJHUp/3UIUIX2dQ3/VSCLTQyLRSSJMesU/1Uci00Mi0UkxwEAAAAA6wOLfRA793IZU1A793UF/1Ug6wP/VRyLRSQr9wPfO/dz54X2dC6LRQyLCIvHK8GL/jvGD0L4i0UIVwPBU1DonjcAAItFDAPfg8QMATgr94t9EHXVX15bXcIgAMzMzMzMzFWL7ItNHIPsCFeLfRiFyXR2U4tdDFaDOwB1Ef91CP91JP9VIItFEItNHIkDiwOL8YtVECvQO8GJVRgPQvAzwIl1/IX2dC+LXRQr34ld+GaQi3X8jRQ4igwTi1UYA1UIi134MgwCjRQ4QIgKO8Zy4YtdDItNHCkzK84BdRQD/olNHIXJdZFeW1+L5V3CIADMzFWL7Ojo5///uYBf1wCB6QBA1wADwYtNCFFQ/3UUjUF0/3UQ/3UMakBQjUE0UOg+////XcIQAMzMzMzMzMzMzMxVi+yD7GyLTRRTVlcPtlkDD7ZBAg+2UQfB4gjB4wgL2A+2QQHB4wgL2A+2AcHjCAvYD7ZBBgvQiV3YweIID7ZBBQvQD7ZBBMHiCAvQD7ZBColV9IlV1A+2UQvB4ggL0A+2QQnB4ggL0A+2QQjB4ggL0A+2QQ6JVfCJVdAPtlEPweIIC9APtkENweIIC9APtkEMi00IweIIC9CJVfgPtkECiVXMD7ZRA8HiCAvQD7ZBAcHiCAvQD7YBweIIC9APtkEGiVXsiVXID7ZRB8HiCAvQD7ZBBcHiCAvQD7ZBBMHiCAvQD7ZBColV6IlVxA+2UQvB4ggL0MHi" & _
                                                    "CA+2QQkL0A+2QQjB4ggL0A+2QQ6JVeSJVcAPtlEPweIIC9APtkENweIIC9APtkEMi00MweIIC9CJVeAPtkECiVW8D7ZRA8HiCAvQD7ZBAcHiCAvQD7YBweIIC9APtkEGiVUIiVW4D7ZRB8HiCAvQD7ZBBcHiCAvQD7ZBBMHiCAvQD7ZBColVFIlVtA+2UQvB4ggL0A+2QQnB4ggL0A+2QQjB4ggL0A+2QQ6JVQyJVbAPtlEPweIIC9APtkENweIIC9APtkEMweIIC9CJVfyJVayLVRAPtkoDD7ZCAsHhCAvID7ZCAcHhCAvID7YCweEIC8iJTdyJTagPtnIHD7ZCBg+2egsPtkoOweYIC/DB5wgPtkIFweYIC/DHRZgKAAAAD7ZCBMHmCAvwD7ZCCgv4iXWkD7ZCCcHnCAv4D7ZCCMHnCAv4D7ZCD8HgCAvBiX2gD7ZKDcHgCAvBD7ZKDItV3MHgCAvBi03siUWc6wOLXRAD2YtNCDPTiV0QwcIQA8qJTQgzTezBwQwD2TPTiV0Qi10IwcIIA9qJVdyLVfQDVegz8oldCDPZwcYQi00UA87BwweJTRQzTejBwQwD0TPyiVX0i1UUwcYIA9aJdeyLdfADdeQz/olVFDPRwccQi00MA8/BwgeJTQwzTeTBwQwD8TP+iXXwi3UMwccIA/eJfZSLffgDfeAzx4l1DDPxwcAQi038A8jBxgeJTfwzTeDBwQwD+TPHiX34i338wcAIA/iJffwz+YtNEAPKwccHM8GJTRCLTQzBwBADyIlNDDPKi1UQwcEMA9EzwolVEItVDMHACAPQiVUMM9GLTfQDzsHCB4lN9IlV6ItV3DPRi038wcIQA8qJTfwzzot19MHBDAPxM9aJ" & _
                                                    "dfSLdfzBwggD8ol1/DPxi03wA8/BxgeJTfCJdeSLdewz8YtNCMHGEAPOiU0IM8+LffDBwQwD+TP3iX3wi30IwcYIA/6JfQgz+YtN+APLwccHiX3gi32UM/mJTfiLTRTBxxADz4lNFDPLi134wcEMA9kz+4ld+MHHCAF9FItdFDPZi8uJXezBwQeDbZgBi134iU3sD4VA/v//AUWcAV3Mi03YA00QAVWoi1UYiU3Yi13Yi8OLTdQDTfSIGolN1ItN0ANN8MHoCIhCAYvDiU3Qi03sAU3Ii03EA03owegQiEICwesYiFoDi13Ui8OIWgTB6AiIQgWLw4lNxItNwANN5MHoEIhCBolNwItNvANN4AF1pAF9oMHrGIhaB4td0IvDiFoIiU28i024A00IwegIiEIJi8OJTbiLTbQDTRTB6BCIQgrB6xiIWguLXcyLw4lNtItNsANNDIhaDMHoCIhCDYvDiU2wi02sA038wegQiEIOwesYiFoPi13Ii8OJTayIWhDB6AiIQhGLw8HoEIhCEsHrGIhaE4tdxIvDiFoUwegIiEIVi8PB6BCIQhbB6xiIWheLXcCLw4haGMHoCIhCGYvDwegQiEIawesYiFobi128i8OIWhzB6AiIQh2Lw8HoEIhCHsHrGIhaH4tduIvDiFogwegIiEIhi8PB6BCIQiLB6xiIWiOLXbSLw4haJMHoCIhCJYvDwegQiEImwesYiFoni12wi8OIWijB6AiIQimLw8HoEIhCKsHrGIhaK4vZiFosi8PB6AiIQi2Lw8HoEIhCLsHrGIhaL4tdqIvDiFowwegIiEIxjUo8i8PB6xjB6BCIQjKIWjOLXaSLw4haNMHoCIhCNYvDwegQiEI2wesYiFo3" & _
                                                    "i12gi8OIWjjB6AiIQjmLw8HoEIhCOsHrGIhaO4tVnIvCwegIiBGIQQGLwl/B6BDB6hheiEECiFEDW4vlXcIUAMxVi+xW/3UQi3UI/3UMVuhtPwAAahD/dRSNRiBQ6C8wAACLRRiDxAzHRnQAAAAAiUZ4Xl3CFADMzMzMzMzMzMzMVYvsVot1CFf/dQz/djCNfiBXjUYQUFboRPn//4tWeDPAgAcBdQtAO8J0BoAEOAF09V9eXcIIAMzMzMzMzMzMzFWL7IPsEI1F8GoQ/3UgUOi8LwAAg8QMjUXwUGoA/3Uk/3Uc/3UY/3UU/3UQ/3UM/3UI6Bk6AACL5V3CIADMzMxVi+z/dSRqAf91IP91HP91GP91FP91EP91DP91COjuOQAAXcIgAMzMzMzMzMzMzMxVi+zoWOD//7kwc9cAgekAQNcAA8GLTQhRUP91FIsB/3UQ/3UM/zCNQShQjUEYUOis9///XcIQAMzMzMzMzMzMVYvsi00Ii0UMiUEsi0UQiUEwXcIMAMzMzMzMzMzMzMxVi+xWi3UIajRqAFboHy8AAItNDMdGLAAAAACLAYlGMItFEIlGBI1GCIkOx0YoAAAAAP8x/3UUUOjDLgAAg8QYXl3CEADMzMzMzMzMzMzMzFWL7IHsIAQAAFNWV2pwjYVw/f//x4Vg/f//QdsAAGoAUMeFZP3//wAAAADHhWj9//8BAAAAx4Vs/f//AAAAAOicLgAAi3UMjYVg////ah9WUOhaLgAAikYfg8QYgKVg////+CQ/DECIhX////+NheD7////dRBQ6MRFAAAPV8CNtWD+//9mDxOFYP7//429aP7//7keAAAAZg8TRYDzpbkeAAAAZg8TheD+//+NdYDHhWD+" & _
                                                    "//8BAAAAjX2Ix4Vk/v//AAAAAPOluR4AAADHRYABAAAAjbXg/v//x0WEAAAAAI296P7//7v+AAAA86W5IAAAAI214Pv//4294P3///Oli8MPtsvB+AOD4QcPtrQFYP///42F4P3//9Pug+YBVlCNRYBQ6JY7AABWjYVg/v//UI2F4P7//1DogjsAAI2F4P7//1CNRYBQjYXg/P//UOgr6///jYXg/v//UI1FgFBQ6HpDAACNhWD+//9QjYXg/f//UI2F4P7//1DoAOv//42FYP7//1CNheD9//9QUOhMQwAAjYXg/P//UI2FYP7//1DoGUMAAI1FgFCNhWD8//9Q6AlDAACNRYBQjYXg/v//UI1FgFDoBS8AAI2F4Pz//1CNheD9//9QjYXg/v//UOjrLgAAjYXg/v//UI1FgFCNheD8//9Q6ITq//+NheD+//9QjUWAUFDo00IAAI1FgFCNheD9//9Q6KNCAACNhWD8//9QjYVg/v//UI2F4P7//1DoqUIAAI2FYP3//1CNheD+//9QjUWAUOiCLgAAjYVg/v//UI1FgFBQ6CHq//+NRYBQjYXg/v//UFDoYC4AAI2FYPz//1CNhWD+//9QjUWAUOhJLgAAjYXg+///UI2F4P3//1CNhWD+//9Q6C8uAACNheD8//9QjYXg/f//UOgMQgAAVo2F4P3//1CNRYBQ6Ps5AABWjYVg/v//UI2F4P7//1Do5zkAAIPrAQ+JH/7//42F4P7//1BQ6MEpAACNheD+//9QjUWAUFDo0C0AAI1FgFD/dQjo5DAAAF9eW4vlXcIMAMzMzMzMzMzMzMzMVYvsg+wgjUXgxkXgCVD/dQwPV8DHRfkAAAAA/3UIDxFF4WbHRf0A"
Private Const STR_THUNK2                As String = "AGYP1kXxxkX/AOiq/P//i+VdwggAzMzMzFWL7IHsFAEAAFOLXQiNRfBWV4t9DA9XwFBQi0MEV8ZF8ABmD9ZF8cdF+QAAAABmx0X9AADGRf8A/9CLdSSD/gx1IFb/dSCNRdBQ6AErAACDxAxmx0XdAADGRdwAxkXfAeswjUXwUI2F7P7//1DorigAAFb/dSCNhez+//9Q6M4mAACNRdBQjYXs/v//UOh+JwAAjUXwUI2FPP///1DofigAAP91HI2FPP////91GFDofCYAAI1F0MZF4ABQV1ONRYzHRekAAAAAD1fAZsdF7QAAUGYP1kXhxkXvAOhw+///agRqDI1FjFDoQ/v//2oQjUXgUFCNRYxQ6PP6////dRSNhTz/////dRBQ6EEmAACNRcBQjYU8////UOjxJgAAi3UsjUXgVlCNRcBQUOi/ZAAAMtKNRcC7AQAAAIX2dBqLfSiLyCv5igwHjUABMkj/CtEr83XxhNJ1FP91FI1FjP91MP91EFDohfr//zPbD1fADxFF8IpF8A8RRdCKRdAPEUXgikXgDxFFwIpFwGpQjYU8////agBQ6OQpAACKjTz///+NRYxqNGoAUOjRKQAAik2Mg8QYi8NfXluL5V3CLABVi+yB7BQBAABTi10IjUXwVleLfQwPV8BQUItDBFfGRfAAZg/WRfHHRfkAAAAAZsdF/QAAxkX/AP/Qi3Ukg/4MdSBW/3UgjUXQUOhBKQAAg8QMZsdF3QAAxkXcAMZF3wHrMI1F8FCNhez+//9Q6O4mAABW/3UgjYXs/v//UOgOJQAAjUXQUI2F7P7//1DoviUAAI1F8FCNhTz///9Q6L4mAAD/dRyNhTz/////dRhQ6LwkAACNRdDGReAAUFdTjUWMx0XpAAAAAA9X" & _
                                                    "wGbHRe0AAFBmD9ZF4cZF7wDosPn//2oEagyNRYxQ6IP5//9qEI1F4FBQjUWMUOgz+f//i30UjUWMi3UoV1b/dRBQ6B/5//9XVo2FPP///1DocSQAAI1FwMZFwABQjYU8////x0XJAAAAAA9XwGbHRc0AAFBmD9ZFwcZFzwDoBCUAAP91MI1F4FCNRcBQ/3Us6NFiAAAPV8APEUXwikXwDxFF0IpF0A8RReCKReAPEUXAikXAalCNhTz///9qAFDoMigAAIqFPP///2o0jUWMagBQ6B8oAACKRYyDxBhfXluL5V3CLABVi+yLVQyLTRBWi3UIiwYzAokBi0YEM0IEiUEEi0YIM0IIiUEIi0YMM0IMiUEMXl3CDADMzMzMzMzMzMzMzMzMVYvsUVOLXQxWV4t9CGbHRfwA4YsPi8HR6IPhAYkDi1cEi8LR6IPiAcHhHwvIweIfiUsEi3cIi8bR6IPmAQvQweYfiVMIi08Mi8HR6IPhAQvwX4lzDA+2RA38weAYMQNeW4vlXcIIAMzMzMzMzMzMzFWL7ItVDFaLdQgPtg4PtkYBweEIC8gPtkYCweEIC8gPtkYDweEIC8iJCg+2TgQPtkYFweEIC8gPtkYGweEIC8gPtkYHweEIC8iJSgQPtk4ID7ZGCcHhCAvID7ZGCsHhCAvID7ZGC8HhCAvIiUoID7ZODA+2Rg3B4QgLyA+2Rg7B4QgLyA+2Rg/B4QgLyIlKDF5dwggAzMzMzMzMzMzMzMxVi+yD7CBWV2oQjUXgagBQ6KsmAABqEP91DI1F8FDobSYAAIt9CIPEGA8QTeAz9pCLxrkfAAAAg+AfK8iLxsH4BYsEh9PoqAF0DA8QRfBmD+/IDxFN4I1F8FBQ6JD+" & _
                                                    "//9Ggf6AAAAAfMdqEI1F4FD/dRDoGSYAAIPEDF9ei+VdwgwAzMzMzMzMzMzMzMzMzMxVi+xWi3UMV4t9CIsXi8LB6BiIBovCwegQiEYBi8LB6AiIRgKIVgOLTwSLwcHoGIhGBIvBwegQiEYFi8HB6AiIRgaITgeLTwiLwcHoGIhGCIvBwegQiEYJi8HB6AiIRgqITguLTwyLwcHoGIhGDIvBwegQiEYNi8HB6AiIRg5fiE4PXl3CCADMzMzMzMzMzMxVi+yD7ERWi3UIg76oAAAAAHQGVuiHLQAAM8kPH0QAAA+2hA6IAAAAiUSNvEGD+RBy7lbHRfwAAAAA6GEsAACNRbxQVuj3KwAAi1UMM8lmkIoEjogEEUGD+RBy9GisAAAAagBW6DclAACKBoPEDF6L5V3CCADMzMzMzMzMzMzMzFWL7FaLdQhorAAAAGoAVugMJQAAi00MahD/dRAPtgGJRkQPtkEBiUZID7ZBAolGTA+2QQOD4A+JRlAPtkEEJfwAAACJRlQPtkEFiUZYD7ZBBolGXA+2QQeD4A+JRmAPtkEIJfwAAACJRmQPtkEJiUZoD7ZBColGbA+2QQuD4A+JRnAPtkEMJfwAAACJRnQPtkENiUZ4D7ZBDolGfA+2QQ+D4A/HhoQAAAAAAAAAiYaAAAAAjYaIAAAAUOgxJAAAg8QYXl3CDADMzMzMzMzMzMxVi+zoGNX//7nQmdcAgekAQNcAA8GLTQhRUP91EI2BqAAAAP91DGoQUI2BmAAAAFDoa+v//13CDADMzMzMzMzMVYvsg+wYU1ZX6NLU////dQi+QJ/XALlAAAAAge4AQNcAA/CLRQhWjXhki0Bg9+EDB4vYg9IAg8AIg+A/K8hRagBq" & _
                                                    "AGiAAAAAakBXi30ID6TaA4lV/I1HIMHjA1CJVfjoDOr//4tV/IvLi8KIXe/B6BiIReiLwsHoEIhF6YvCwegIiEXqikX4iEXri8IPrMEYagjB6BiITeyLwovLD6zBEMHoEIvDiE3tD6zQCIhF7o1F6FDB6ghX6GQBAACLF4vCi3UMwegYiAaLwsHoEIhGAYvCwegIiEYCiFYDi08Ei8HB6BiIRgSLwcHoEIhGBYvBwegIiEYGiE4Hi08Ii8HB6BiIRgiLwcHoEIhGCYvBwegIiEYKiE4Li08Mi8HB6BiIRgyLwcHoEIhGDYvBwegIiEYOiE4Pi08Qi8HB6BiIRhCLwcHoEIhGEYvBwegIiEYSiE4Ti08Ui8HB6BiIRhSLwcHoEIhGFYvBwegIiEYWiE4Xi08Yi8HB6BiIRhiLwcHoEIhGGYvBwegIiEYaiE4bi08ci8HB6BiIRhyLwcHoEIhGHYvBamjB6AhqAIhGHleITh/oWSIAAIPEDF9eW4vlXcIIAMzMzMzMzMzMzMzMzMxVi+xWi3UIamhqAFboLyIAAIPEDMcGZ+YJasdGBIWuZ7vHRghy8248x0YMOvVPpcdGEH9SDlHHRhSMaAWbx0YYq9mDH8dGHBnN4FteXcIEAFWL7Oi40v//uUCf1wCB6QBA1wADwYtNCFFQ/3UQjUFk/3UMakBQjUEgUOgR6f//XcIMAMzMzMzMzMzMzMzMzMxVi+yD7ECNRcBQ/3UI6L4AAABqMI1FwFD/dQzoYCEAAIPEDIvlXcIIAMzMzMzMzMxVi+xWi3UIaMgAAABqAFbobCEAAIPEDMcG2J4FwcdGBF2du8vHRggH1Xw2x0YMKimaYsdGEBfdcDDHRhRaAVmRx0YYOVkO" & _
                                                    "98dGHNjsLxXHRiAxC8D/x0YkZyYzZ8dGKBEVWGjHRiyHSrSOx0Ywp4/5ZMdGNA0uDNvHRjikT/q+x0Y8HUi1R15dwgQAzMzMzMzpGwQAAMzMzMzMzMzMzMzMVYvsg+wci0UIU42YxAAAAFaLgMAAAABXv4AAAAD354vwAzOLxoPSAA+kwgPB4AOJVfyJRfiJVfToc9H///91CLkQodcAgekAQNcAA8FQjUYQi3UIg+B/K/hXagBqAGiAAAAAaIAAAABTjUZAUOjO5v//agiNReTHReQAAAAAUFbHRegAAAAA6IQDAACLXfyLw4tV+IvKwegYiEXki8PB6BCIReWLw8HoCIhF5opF9IhF54vDD6zBGGoIwegYiE3oi8OLyohV6w+swRDB6BCLwohN6Q+s2AiIReqNReRQVsHrCOgpAwAAi14Ei8OLDolN/MHoGIt9DIgHi8PB6BCIRwGLw8HoCIhHAovDD6zBGIhfA8HoGIhPBIvDi038D6zBEMHoEIhPBYtN/IvBD6zYCIhHBovGiE8HwesIi1gIi8uLUAyLwsHoGIhHCIvCwegQiEcJi8LB6AiIRwqLwg+swRiIVwvB6BiITwyLwovLD6zBEMHoEIhPDYvDD6zQCIhHDovGiF8PweoIi1gQi8uLUBSLwsHoGIhHEIvCwegQiEcRi8LB6AiIRxKLwg+swRiIVxPB6BiITxSLwovLD6zBEMHoEIvDiE8VD6zQCIhHFovGweoIiF8Xi1gYi8uLUByLwsHoGIhHGIvCwegQiEcZi8LB6AiIRxqLwg+swRiIVxvB6BiITxyLwovLD6zBEMHoEIhPHYvDD6zQCIhHHovGiF8fweoIi1ggi8uLUCSLwsHoGIhHIIvCwegQ" & _
                                                    "iEchi8LB6AiIRyKLwg+swRiIVyPB6BiITySLwovLD6zBEMHoEIhPJYvDD6zQCIhHJovGiF8nweoIi1goi8uLUCyLwsHoGIhHKIvCwegQiEcpi8LB6AiIRyqLwg+swRiIVyvB6BiITyyLwovLD6zBEMHoEIvDiE8tD6zQCMHqCIhHLovGiF8vjXc4aMgAAABqAItYMIvLi1A0i8LB6BiIRzCLwsHoEIhHMYvCwegIiEcyi8IPrMEYiFczwegYiE80i8KLyw+swRDB6BCITzWLww+s0AiIRzaIXzeLfQjB6ghXi1c8i8KLXziLy8HoGIgGi8LB6BCIRgGLwsHoCIhGAovCD6zBGIhWA8HoGIhOBIvCi8sPrMEQwegQi8OITgUPrNAIiEYGweoIiF4H6HUdAACDxAxfXluL5V3CCADMzMzMzMzMzMxVi+xWi3UIaMgAAABqAFboTB0AAIPEDMcGCMm888dGBGfmCWrHRgg7p8qEx0YMha5nu8dGECv4lP7HRhRy8248x0YY8TYdX8dGHDr1T6XHRiDRguatx0Ykf1IOUcdGKB9sPivHRiyMaAWbx0Ywa71B+8dGNKvZgx/HRjh5IX4Tx0Y8Gc3gW15dwgQAzMzMzMxVi+zomM3//7kQodcAgekAQNcAA8GLTQhRUP91EI2BxAAAAP91DGiAAAAAUI1BQFDo6+P//13CDADMzMzMzMzMVYvsVot1CP91DIsOjUYIUP92BItBBP/Qi1Ysi0YwA9ZIXoBEAggBdRMPH4AAAAAAhcB0CEiARAIIAXT0XcIIAFWL7FOLXQxWV4t9CA+2QyiZi8iL8g+kzggPtkMpweEImQvIC/IPpM4ID7ZDKsHhCJkLyAvyD6TOCA+2QyvB" & _
                                                    "4QiZC8gL8g+2QywPpM4ImcHhCAvyC8gPtkMtD6TOCJnB4QgL8gvID7ZDLg+kzgiZweEIC/ILyA+2Qy8PpM4ImcHhCAvyC8iJdwSJDw+2QyCZi8iL8g+2QyEPpM4ImcHhCAvyC8gPtkMiD6TOCJnB4QgL8gvID7ZDIw+kzgiZweEIC/ILyA+2QyQPpM4ImcHhCAvIC/IPpM4ID7ZDJcHhCJkLyAvyD6TOCA+2QybB4QiZC8gL8g+kzggPtkMnweEImQvIC/KJTwiJdwwPtkMYmYvIi/IPpM4ID7ZDGcHhCJkLyAvyD7ZDGg+kzgiZweEIC/ILyA+2QxsPpM4ImcHhCAvyC8gPtkMcD6TOCJnB4QgL8gvID7ZDHQ+kzgiZweEIC/ILyA+2Qx4PpM4ImcHhCAvyC8gPtkMfD6TOCJnB4QgL8gvIiXcUiU8QD7ZDEJmLyIvyD7ZDEQ+kzgiZweEIC/ILyA+2QxIPpM4IweEImQvIC/IPpM4ID7ZDE8HhCJkLyAvyD6TOCA+2QxTB4QiZC8gL8g+kzggPtkMVweEImQvIC/IPpM4ID7ZDFsHhCJkLyAvyD7ZDFw+kzgiZweEIC/ILyIl3HIlPGA+2QwiZi8iL8g+2QwkPpM4ImcHhCAvyC8gPtkMKD6TOCJnB4QgL8gvID7ZDCw+kzgiZweEIC/ILyA+2QwwPpM4ImcHhCAvyC8gPtkMND6TOCJnB4QgL8gvID7ZDDg+kzgiZweEIC/ILyA+2Qw8PpM4ImcHhCAvIC/KJTyCJdyQPtgOZi8iL8g+2QwEPpM4ImcHhCAvyC8gPtkMCD6TOCJnB4QgL8gvID7ZDAw+kzgiZweEIC/ILyA+2QwQPpM4ImcHhCAvyC8gPtkMF" & _
                                                    "D6TOCJnB4QgL8gvID7ZDBg+kzgiZweEIC/ILyA+2QwcPpM4ImcHhCAvIC/KJdyyJTyhfXltdwggAzMzMzMxVi+xTi10MVleLfQgPtkMYmYvIi/IPpM4ID7ZDGcHhCJkLyAvyD6TOCA+2QxrB4QiZC8gL8g+kzggPtkMbweEImQvIC/IPtkMcD6TOCJnB4QgL8gvID7ZDHQ+kzgiZweEIC/ILyA+2Qx4PpM4ImcHhCAvyC8gPtkMfD6TOCJnB4QgL8gvIiXcEiQ8PtkMQmYvIi/IPtkMRD6TOCJnB4QgL8gvID7ZDEg+kzgiZweEIC/ILyA+2QxMPpM4ImcHhCAvyC8gPtkMUD6TOCJnB4QgLyAvyD6TOCA+2QxXB4QiZC8gL8g+kzggPtkMWweEImQvIC/IPpM4ID7ZDF8HhCJkLyAvyiU8IiXcMD7ZDCJmLyIvyD6TOCA+2QwnB4QiZC8gL8g+2QwoPpM4ImcHhCAvyC8gPtkMLD6TOCJnB4QgL8gvID7ZDDA+kzgiZweEIC/ILyA+2Qw0PpM4ImcHhCAvyC8gPtkMOD6TOCJnB4QgL8gvID7ZDDw+kzgiZweEIC/ILyIl3FIlPEA+2A5mLyIvyD7ZDAQ+kzgiZweEIC/ILyA+2QwIPpM4IweEImQvIC/IPtkMDD6TOCJnB4QgL8gvID7ZDBA+kzgiZweEIC/ILyA+2QwUPpM4ImcHhCAvyC8gPtkMGD6TOCJnB4QgL8gvID7ZDBw+kzgiZweEIC8gL8ol3HIlPGF9eW13CCADMzMxVi+yB7JAAAACNRdD/dQxQ6Mv6//+NRdBQ6FI0AACFwHQIM8CL5V3CCACNRdBQ6K3H//8FYAEAAFDoUjMAAIP4AXQV6JjH" & _
                                                    "//8FYAEAAFCNRdBQUOiYTQAAagCNRdBQ6H3H//8FAAEAAFCNhXD///9Q6NvK//+NhXD///9Q6G/K//+FwHWdikWgi00IJAEEAogBjYVw////UI1BAVDorwAAALgBAAAAi+VdwggAzMzMzFWL7IPsYI1F4P91DFDoLv3//41F4FDo1TMAAIXAdAgzwIvlXcIIAI1F4FDoAMf//4PogFDoFzMAAIP4AXQT6O3G//+D6IBQjUXgUFDo/04AAGoAjUXgUOjUxv//g8BAUI1FoFDo58v//41FoFDo/sn//4XAdamKRcCLTQgkAQQCiAGNRaBQjUEBUOiBAgAAuAEAAACL5V3CCADMzMzMzMxVi+xWi3UIsShXi30MD7ZHB4hGKA+2RwaIRimLB4tXBOg70///iEYqsSCLB4tXBOgs0///iEYriw+LRwQPrMEYiE4siw/B6BiLRwQPrMEQiE4tiw/B6BCLRwQPrMEIiE4usSjB6AgPtgeIRi8PtkcPiEYgD7ZHDohGIYtHCItXDOjb0v//iEYisSCLRwiLVwzoy9L//4hGI4tPCItHDA+swRiITiSLTwjB6BiLRwwPrMEQiE4li08IwegQi0cMD6zBCIhOJrEowegID7ZHCIhGJw+2RxeIRhgPtkcWiEYZi0cQi1cU6HbS//+IRhqxIItHEItXFOhm0v//iEYbi08Qi0cUD6zBGIhOHItPEMHoGItHFA+swRCITh2LTxDB6BCLRxQPrMEIiE4esSjB6AgPtkcQiEYfD7ZHH4hGEA+2Rx6IRhGLRxiLVxzoEdL//4hGErEgi0cYi1cc6AHS//+IRhOLTxiLRxwPrMEYiE4Ui08YwegYi0ccD6zBEIhOFYtPGMHoEItHHA+s" & _
                                                    "wQiIThaxKMHoCA+2RxiIRhcPtkcniEYID7ZHJohGCYtHIItXJOis0f//iEYKsSCLRyCLVyTonNH//4hGC4tPIItHJA+swRjB6BiITgyLTyCLRyQPrMEQwegQiE4Ni08gi0ckD6zBCMHoCIhODg+2RyCIRg8PtkcviAYPtkcuiEYBsSiLRyiLVyzoSNH//4hGArEgi0coi1cs6DjR//+IRgOLTyiLRywPrMEYwegYiE4Ei08oi0csD6zBEMHoEIhOBYtPKItHLA+swQjB6AiITgYPtkcoX4hGB15dwggAzMzMzMzMzMxVi+xWi3UIsShXi30MD7ZHB4hGGA+2RwaIRhmLB4tXBOjL0P//iEYasSCLB4tXBOi80P//iEYbiw+LRwQPrMEYiE4ciw/B6BiLRwQPrMEQiE4diw/B6BCLRwQPrMEIiE4esSjB6AgPtgeIRh8PtkcPiEYQD7ZHDohGEYtHCItXDOhr0P//iEYSsSCLRwiLVwzoW9D//4hGE4tPCItHDA+swRiIThSLTwjB6BiLRwwPrMEQiE4Vi08IwegQi0cMD6zBCIhOFrEowegID7ZHCIhGFw+2RxeIRggPtkcWiEYJi0cQi1cU6AbQ//+IRgqxIItHEItXFOj2z///iEYLi08Qi0cUD6zBGIhODItPEMHoGItHFA+swRCITg2LTxDB6BCLRxQPrMEIiE4OsSjB6AgPtkcQiEYPD7ZHH4gGD7ZHHohGAYtHGItXHOiiz///iEYCsSCLRxiLVxzoks///4hGA4tPGItHHA+swRjB6BiITgSLTxiLRxwPrMEQwegQiE4Fi08Yi0ccD6zBCMHoCIhOBg+2RxhfiEYHXl3CCADMzFWL7IPsMFOLXQgPV8BW" & _
                                                    "i3UMx0XQAwAAAMdF1AAAAAAPEUXYjUYBZg/WRfhQUw8RRejoSvX//4A+BHUVjUYxUI1DMFDoOPX//15bi+VdwggAV1ONezBX6BU/AADoIML//wWgAAAAUI1F0FBXV+hfPwAAU1dX6Dc9AADoAsL//wWgAAAAUOj3wf//BdAAAABQV1foejYAAFfo1BEAAIoGM/aLDyQBD7bAg+EBmTvIdQQ78nQSV+jHwf//BaAAAABQV+jLRwAAX15bi+VdwggAzMxVi+yD7CBTi10ID1fAV4t9DMdF4AMAAADHReQAAAAADxFF6I1HAWYP1kX4UFPojvf//4A/BHUVjUchUI1DIFDofPf//19bi+VdwggAVlONcyBW6Hk+AADoVMH//1CNReBQVlboyD4AAFNWVugAPgAA6DvB//9Q6DXB//+DwCBQVlbo+jUAAFboxBEAAIoHM/+LDiQBD7bAg+EBmTvIdQQ7+nQNVugHwf//UFboIEkAAF5fW4vlXcIIAMzMzMzMzMxVi+yB7PAAAACNhRD/////dQhQ6Fj+////dQyNRdBQ6Mzz//9qAI1F0FCNhRD///9QjYVw////UOgjxP//jYVw////UP91EOgU+v//jYVw////UOiow///99gbwECL5V3CDADMzMzMzMzMzMzMzMzMVYvsgeygAAAAjYVg/////3UIUOi4/v///3UMjUXgUOhs9v//agCNReBQjYVg////UI1FoFDoZsX//41FoFD/dRDoGvz//41FoFDoccP///fYG8BAi+VdwgwAzMzMzMzMVYvsg+xgjUWgVv91CFDojf3//4t1DI1FoFCNRgHGBgRQ6Gr5//+NRdBQjUYxUOhd+f//uAEAAABei+VdwggAzFWL" & _
                                                    "7IPsQI1FwFb/dQhQ6B3+//+LdQyNRcBQjUYBxgYEUOia+///jUXgUI1GIVDojfv//7gBAAAAXovlXcIIAMxVi+yB7MAAAABXi30QV+gdLAAAhcB0CTPAX4vlXcIQAFfoer///wVgAQAAUOgfKwAAg/gBdBLoZb///wVgAQAAUFdX6GhFAABqAFfoUL///wUAAQAAUI2FQP///1DorsL//42FQP///1DoMr///wVgAQAAUOjXKgAAg/gBdBjoHb///wVgAQAAUI2FQP///1BQ6BpFAACNhUD///9Q6I4rAACFwA+Fbf///1aLdRSNhUD///9QVuhV+P///3UIjUWgUOjZ8f//6NS+//8FYAEAAFCNRaBQjYVA////UI1F0FDoajgAAP91DI1FoFDorvH//+ipvv//BWABAABQjUXQUI1FoFCNRdBQ6CIzAADojb7//wVgAQAAUFdX6JAzAADoe77//wVgAQAAUFeNRdBQUOgaOAAAjUXQUI1GMFDozff//164AQAAAF+L5V3CEABVi+yB7IAAAABXi30QV+j9KgAAhcB0CTPAX4vlXcIQAFfoKr7//4PogFDoQSoAAIP4AXQQ6Be+//+D6IBQV1foLEYAAGoAV+gEvv//g8BAUI1FgFDoF8P//41FgFDo7r3//4PogFDoBSoAAIP4AXQT6Nu9//+D6IBQjUWAUFDo7UUAAI1FgFDohCoAAIXAdYdWi3UUjUWAUFbokvn///91CI1FwFDotvP//+ihvf//g+iAUI1FwFCNRYBQjUXgUOjsOAAA/3UMjUXAUOiQ8///6Hu9//+D6IBQjUXgUI1FwFCNReBQ6DYyAADoYb3//4PogFBXV+jGNAAA6FG9//+D6IBQV41F" & _
                                                    "4FBQ6KI4AACNReBQjUYgUOgV+f//XrgBAAAAX4vlXcIQAMzMzMzMzMzMVYvsgeyAAgAAjYWA/f//Vv91CFDoh/r//4t1EI2F0P7//1ZQ6Pfv//+NRjBQjYVg////UOjn7///jYXQ/v//UOhrKQAAhcAPhZwDAACNhWD///9Q6FcpAACFwA+FiAMAAI2F0P7//1Dos7z//wVgAQAAUOhYKAAAg/gBD4VoAwAAjYVg////UOiTvP//BWABAABQ6DgoAACD+AEPhUgDAABTV+h4vP//BWABAABQjYVg////UI1FwFDocjEAAP91DI2FAP///1DoU+///+hOvP//BWABAABQjUXAUI2FAP///1BQ6Oc1AADoMrz//wVgAQAAUI1FwFCNhdD+//9QjYWg/v//UOjFNQAAjYWA/f//UI2FEP7//1DoMj4AAI2FsP3//1CNhUD+//9Q6B8+AADo6rv//wUAAQAAUI2FMP///1DoCD4AAOjTu///BTABAABQjYVg////UOjxPQAA6Ly7//8FoAAAAFCNhTD///9QjYUQ/v//UI1FwFDo7zgAAI2FQP7//1CNhRD+//9QjYVg////UI2FMP///1DoPsL//+h5u///BaAAAABQjUXAUFDoeTAAAI1FwFCNhUD+//9QjYUQ/v//UOiSyv//x0XwAAAAAOhGu///BQABAACJRfSNhYD9//+JRfiNhRD+//+JRfyNhaD+//9Q6JA7AACL2I2FAP///1DogjsAADvDD0fYjYUA////jXP/VlDofUQAAAvCdAe/AQAAAOsCM/9WjYWg/v//UOhjRAAAC8J0B74CAAAA6wIz9gv3jUWQi3S18FZQ6PY8AACNRjBQjYVw/v//UOjmPAAA" & _
                                                    "jUXAUOitJQAAjXP+x0XAAQAAAMdFxAAAAACF9g+I6AAAAA8fQACNRcBQjYVw/v//UI1FkFDovLr//1aNhQD///9Q6O9DAAALwnQHvwEAAADrAjP/Vo2FoP7//1Do1UMAAAvCdAe4AgAAAOsCM8ALx4t8hfCF/w+EhQAAAFeNhTD///9Q6F08AACNRzBQjYVg////UOhNPAAAjUXAUI2FYP///1CNhTD///9Q6EbJ///oAbr//wWgAAAAUI2FMP///1CNRZBQjYXg/f//UOg0NwAAjYVw/v//UI1FkFCNhWD///9QjYUw////UOiGwP//jYXg/f//UI1FwFBQ6OU0AACD7gEPiRz////op7n//wWgAAAAUI1FwFBQ6KcuAACNRcBQjYVw/v//UI1FkFDow8j//41FkFDoern//wVgAQAAUOgfJQAAX1uD+AF0Fehjuf//BWABAABQjUWQUFDoYz8AAI2F0P7//1CNRZBQ6PMkAAD32F4bwECL5V3CDAAzwF6L5V3CDADMzMzMzMzMzMzMzMzMzFWL7IHssAEAAI2FUP7//1b/dQhQ6Ff3//+LdRCNhTD///9WUOgH7///jUYgUI1FkFDo+u7//42FMP///1DoniUAAIXAD4VUAwAAjUWQUOiNJQAAhcAPhUMDAACNhTD///9Q6Lm4//+D6IBQ6NAkAACD+AEPhSUDAACNRZBQ6J64//+D6IBQ6LUkAACD+AEPhQoDAABTV+iFuP//g+iAUI1FkFCNReBQ6OQvAAD/dQyNhVD///9Q6HXu///oYLj//4PogFCNReBQjYVQ////UFDoqzMAAOhGuP//g+iAUI1F4FCNhTD///9QjYUQ////UOiLMwAAjYVQ/v//UI2F"
Private Const STR_THUNK3                As String = "sP7//1DoqDoAAI2FcP7//1CNhdD+//9Q6JU6AADoALj//4PAQFCNhXD///9Q6IA6AADo67f//4PAYFCNRZBQ6G46AADo2bf//1CNhXD///9QjYWw/v//UI1F4FDoQTUAAI2F0P7//1CNhbD+//9QjUWQUI2FcP///1DoY7///+iet///UI1F4FBQ6AMvAACNReBQjYXQ/v//UI2FsP7//1DoDMf//8dF0AAAAADocLf//4PAQIlF1I2FUP7//4lF2I2FsP7//4lF3I2FEP///1DoDDgAAIv4jYVQ////UOj+NwAAO8cPR/iNhVD///+NX/9TUOipQAAAC8J0B74BAAAA6wIz9lONhRD///9Q6I9AAAALwnQHuAIAAADrAjPAC/CNRbCLdLXQVlDogjkAAI1GIFCNhfD+//9Q6HI5AACNReBQ6DkiAACNd/7HReABAAAAx0XkAAAAAIX2D4jSAAAAjUXgUI2F8P7//1CNRbBQ6Gy4//9WjYVQ////UOgfQAAAC8J0B78BAAAA6wIz/1aNhRD///9Q6AVAAAALwnQHuAIAAADrAjPAC8eLfIXQhf90d1eNhXD///9Q6PE4AACNRyBQjUWQUOjkOAAAjUXgUI1FkFCNhXD///9Q6NDF///oO7b//1CNhXD///9QjUWwUI2FkP7//1DoozMAAI2F8P7//1CNRbBQjUWQUI2FcP///1DoyL3//42FkP7//1CNReBQUOi3MgAAg+4BD4ku////6Om1//9QjUXgUFDoTi0AAI1F4FCNhfD+//9QjUWwUOhaxf//jUWwUOjBtf//g+iAUOjYIQAAX1uD+AF0E+istf//g+iAUI1FsFBQ6L49AACNhTD///9QjUWwUOiuIQAA99heG8BAi+VdwgwAM8Be" & _
                                                    "i+VdwgwAzMzMzMzMzMzMVYvsi00Ii8HB6AeB4X9/f/8lAQEBAQPJa8AbM8FdwgQAzMzMzMzMzMzMzMzMzMzMVYvs6Fi1//+5EIzXAIHpAEDXAAPBi00IUVD/dRCNQTD/dQxqEFCNQSBQ6LHL//9dwgwAzMzMzMzMzMzMzMzMzFWL7ItNCItFEAFBOINRPACJRRCJTQhd6aT////MzMzMVYvsVot1CIN+SAF1DVboLQAAAMdGSAIAAACLRRABRkBQ/3UMg1ZEAFbocv///15dwgwAzMzMzMzMzMzMzMzMzFWL7FaLdQiLTjCFyXQpuBAAAAArwVCNRiADwWoAUOjNAwAAg8QMjUYgUFboEAAAAMdGMAAAAABeXcIEAMzMzMxVi+yD7BCNRfBWV1D/dQzoPNz//4t1CI1F8I1+EFdXUOh72///V1ZX6MPc//9fXovlXcIIAMzMzMzMzMzMzMzMVYvsg+wUU1aLdQiLRkiD+AF0BYP4AnUNVuhi////x0ZIAAAAAIteOItWPA+k2gNqCIvCweMDwegYi8uIReyLwsHoEIhF7YvCwegIiEXuD7bCiEXvi8IPrMEYiVX8wegYiE3wi8KLy4hd8w+swRDB6BCLw4hN8Q+s0AiIRfKNRexQweoIVuhW/v//i15Ai1ZED6TaA2oIi8LB4wPB6BiLy4hF7IvCwegQiEXti8LB6AiIRe4PtsKIRe+Lwg+swRiJVfzB6BiITfCLwovLiF3zD6zBEMHoEIvDiE3xD6zQCIhF8o1F7FDB6ghW6PH9////dQyNRhBQ6EXc//9eW4vlXcIIAMzMzMzMzMzMzMzMzMxVi+xWi3UIalBqAFboTwIAAIPEDFb/dQzo49r//8dGSAEAAABe" & _
                                                    "XcIIAMzMzMzMzMxVi+yB7IAAAAC5IAAAAFOLXQxWV4vzjX2A86W+/QAAAI1FgFBQ6OYXAACD/gJ0EIP+BHQLU41FgFBQ6OEDAACD7gF53It9CI11gLkgAAAA86VfXluL5V3CCADMzMzMzMxVi+xTVot1CFdW6AH9//+L2FPo+fz//4vQUujx/P//i/gz/ov3i8czw8HPCDPywcAIi87ByRAzwTPHM8ZfM8MzRQheW13CBADMzMzMzMzMzFWL7FaLdQj/Nuii/////3YEiQbomP////92CIlGBOiN/////3YMiUYI6IL///+JRgxeXcIEAMzMzMzMzMzMzMxVi+xTi10IVlcPtnsHD7ZDAg+2cwsPtlMPwecIC/gPtksDD7ZDDcHnCAv4weYID7ZDCMHnCAv4weIID7ZDBgvwweEID7ZDAcHmCAvwD7ZDDMHmCAvwD7ZDCgvQD7ZDBcHiCAvQD7YDweIIC9APtkMOC8iJUwwPtkMJweEIC8iJcwgPtkMEiXsEweEIXwvIXokLW13CBADMzMzMzMzMzMzMVYvsVuhHsf//i3UIBYMGAABQ/zbolxcAAIkG6DCx//8FgwYAAFD/dgToghcAAIlGBOgasf//BYMGAABQ/3YI6GwXAACJRgjoBLH//wWDBgAAUP92DOhWFwAAiUYMXl3CBADMzMzMzMzMzMzMzMzMzFWL7ItFCIvQVot1EIX2dBVXi30MK/iKDBeNUgGISv+D7gF18l9eXcPMzMzMzMzMzFWL7ItNEIXJdB8PtkUMVovxacABAQEBV4t9CMHpAvOri86D4QPzql9ei0UIXcPMzFWL7FaLdQhW6AP7//+L0IvOM9bByRDBwgjBzggz0TPWM8JeXcIEAMzM" & _
                                                    "zMzMzMzMzFWL7FaLdQj/NujC/////3YEiQbouP////92CIlGBOit/////3YMiUYI6KL///+JRgxeXcIEAMzMzMzMzMzMzMxVi+yD7GAPV8DHRaABAAAAVleNRaDHRaQAAAAAUA8RRajHRdABAAAADxFFuMdF1AAAAABmD9ZFyA8RRdgPEUXoZg/WRfjoxq///wWgAAAAUI1FoFDoNxcAAI1FoFDoHjAAAIt9CI1w/4P+AXYsDx8AjUXQUFDohiwAAFaNRaBQ6Aw5AAALwnQLV41F0FBQ6K0qAABOg/4Bd9eNRdBQV+idMQAAX16L5V3CBADMzMzMzFWL7IPsQFYPV8DHRcABAAAAV41FwMdFxAAAAABQDxFFyMdF4AEAAABmD9ZF2MdF5AAAAAAPEUXoZg/WRfjoHq///1CNRcBQ6NQYAACNRcBQ6MsvAACLfQiNcP+D/gF2KY1F4FBQ6BYsAABWjUXAUOhsOAAAC8J0C1eNReBQUOidKwAAToP+AXfXjUXgUFfoXTEAAF9ei+VdwgQAzMzMzMxVi+yB7AABAACLRQwPV8BTVle5PAAAAGYPE4UA////jbUA////x0X8EAAAAI29CP////Oli00QjZ0I////g8EQi9MrwolN+IlFDGYPH0QAAIv5x0UQBAAAAIvzDx9EAAD/dBgE/zQY/3f0/3fw6M66//8BRviLRQwRVvz/dBgE/zQY/3f8/3f46LO6//8BBotFDBFWBP90GAT/NBj/dwT/N+iauv//AUYIi0UMEVYM/3QYBP80GP93DP93COh/uv//AUYQjX8gi0UMEVYUjXYgg20QAXWKi034g8MIg238AQ+Fav///zP2agBqJv909YT/dPWA6Ee6//8BhPUA" & _
                                                    "////agARlPUE////aib/dPWM/3T1iOgouv//AYT1CP///2oAEZT1DP///2om/3T1lP909ZDoCbr//wGE9RD///9qABGU9RT///9qJv909Zz/dPWY6Oq5//8BhPUY////agARlPUc////aib/dPWk/3T1oOjLuf//AYT1IP///xGU9ST///+DxgWD/g8Pgln///+LXQiNtQD///+5IAAAAIv786VT6Pm8//9T6PO8//9fXluL5V3CDADMzMzMzMzMzMzMVYvsg+wQU1aLdQxXi30YagBWagD/dRToZLn//2oAVmoAV4lF8Iva6FS5//9qAP91EIlF9IvyagBX6EK5//9qAP91EIlF/GoA/3UUiVX46C25//+L+ItF9AP7g9IAA/gT1jvWdw5yBDv4cwiDRfwAg1X4AYtFCDPJC03wiQgzyQNV/Il4BBNN+F9eiVAIiUgMW4vlXcIUAMzMzMzMzMzMzFWL7IPsMFNWi3UIV4t9DFdW6HouAABqIFeNRdBQ6C4ZAACL2IlVDI1OCIPGOI1F0FBRUeioEwAAA8OJBhNVDItFCIPAEIlWBFdQUOiQEwAAi30IiUdAjUXQUFdXiVdE6AwyAACLTQwD2BPKi1cwK9OLXzQb2TtfNHItdwU7VzB2JoMG/4sGg1YE/yNGBIP4/3UVg0YI/412CINWBP+LBiNGBIP4/3TriV80iVcwX15bi+VdwggAzMzMzMzMzMzMzFWL7IHsCAEAAI2FeP///1NWV/91DFDopQkAAI2FeP///1DoWbv//42FeP///1DoTbv//42FeP///1DoQbv//429+P7//7sCAAAAZg8fRAAAi414////i4V8////gent/wAAiY34/v//g9gAiYX8/v//" & _
                                                    "uAgAAABmZg8fhAAAAAAAi3QH+ItMB/yLlAV4////iXX4D6zOEIuMBXz///+D5gHHRAf8AAAAACvWg9kAger//wAAiZQF+P7//4PZAImMBfz+//8Pt034iUwH+IPACIP4eHKsi41o////i4Vs////i1XwD6zBEA+3hWj///+D4QGJhWj///8r0ceFbP///wAAAACLTfS4AQAAAIPZAIHq/38AAImVcP///4PZAImNdP///w+syhCD4gHB+RArwlCNhfj+//9QjYV4////UOiNBwAAg+sBD4UE////i3UIM9KKhNV4////i4zVeP///4gEVouE1Xz///8PrMEIiExWAULB+AiD+hBy119eW4vlXcIIAMzMzMzMzMzMzMzMzMxVi+yLRQgz0lZXi30MK/iNchGLDAeNQAQDSPwD0Q+2ysHqCIlI/IPuAXXnX15dwggAzMzMzMzMzMzMzMzMzMzMVYvsVv91DIt1CFbosP///41GRFBW6LYBAABeXcIIAMxVi+yD7ERTVot1CFcPEAaLRkCJRfwPEUW8DxBGEA8RRcwPEEYgDxFF3A8QRjAPEUXs6Hqp//8FNAUAAFCNRbxQ6Fv///+LRfyNfbz30I1VzCWAAAAAK/65AgAAAI1Y//fQwegfwesfI9j32yvWi8P30IlFCGYPbsNmD3DQAGYPbsCLxmYPcNgADx+EAAAAAACNQCAPEEDgDxBMB+BmD9vCZg/by2YP68gPEUjgDxBA8A8QTALgZg/bwmYP28tmD+vIDxFI8IPpAXXGjVZAjXEBiww6jVIEI00Ii8MjQvwLyIlK/IPuAXXoX15bi+VdwgQAzMzMzMzMzMzMzMzMzMzMVYvsg+xEjUW8VmpEagBQ6Oz3//+L" & _
                                                    "dQiDxAwzwIuWqAAAAIXSdBtmZg8fhAAAAAAAD7aMBpgAAACJTIW8QDvCcu+NRbzHRJW8AQAAAFBW6I3+//9ei+VdwgQAzMzMzMzMVYvsVot1CDPAM9IPH0QAAAMElg+2yIkMlkLB6AiD+hB87gNGQIvIwegCg+EDM9KJTkCNDIADDJYPtsGJBJZCwekIg/oQfO4BTkBeXcIEAMxVi+yD7FSLRQyNTaxTVot1CDPbK8HHRfgQAAAAV4lF8DPSM/8zwIlVCIlV/IXbeFGNSwGD+QJ8MItN8I1VrI0MmQPRiwyGjVL4D69KCAFNCItMhgSDwAIPr0oEAU38jUv/O8F+3otVCDvDfw6LfQyLyyvIizyPD688hotF/APCA/iNQwEz0olVCIvIiVX8iUX0g/gRfXKDffgCfEOLVQyLwyvBjRSCg8JADx+AAAAAAIsEjo1S+A+vQgyNBIDB4AYBRQiLRI4Eg8ECD69CCI0EgMHgBgFF/IP5EHzUi1UIg/kRfRqLVQyLwyvBi0SCRA+vBI6LVQiNBIDB4AYD+ItF/APCA/iLRfSLTfhJiXydrIlN+IvYg/n/D48C////jUWsUOiJ/v//DxBFrItF7F8PEQYPEEW8DxFGEA8QRcwPEUYgDxBF3A8RRjCJRkBeW4vlXcIIAMzMzMzMzMzMzMzMVYvsi1UMg+xEM8APH0QAAA+2DBCJTIW8QIP4EHzyjUW8x0X8AQAAAFD/dQjon/z//4vlXcIIAMzMzMzMzMzMzFWL7IHsfAEAAFNWV2oM/3UMjUXgxkXcAA9XwMdF5QAAAABQZg/WRd1mx0XpAADGResA6En1//+DxAzGRbwAjUXcx0XVAAAAAA9XwGbHRdkAAA8RRb1qBFBq" & _
                                                    "IP91CI2FMP///2YP1kXNUMZF2wDovsT//2ogjUW8UFCNhTD///9Q6Au+//+NRcxQjUW8UI2FhP7//1Do98///w9XwA8RRbyKRbxqII1FvFBQjYUw////UA8RRczo1r3//4t1FA9XwFb/dRAPEUW8ikW8jYWE/v//xkXsAFAPEUXMx0X1AAAAAGYP1kXtZsdF+QAAxkX7AOhr0P//i8b32IPgD1CNRexQjYWE/v//UOhT0P//g30kAYt9IItdHFN1FFf/dRiNhTD///9Q6Ga9//9TV+sD/3UYjYWE/v//UOgj0P//i8P32IPgD1CNRexQjYWE/v//UOgL0P//M9KIXfSLxolV6IhF7IvIi8IPrMEIahDB6AiITe2LwovOD6zBEMHoEIhN7ovCi84PrMEYwegYD7bCiEXwi8LB6AiIRfGLwsHoEIhF8sHqGIhN74vLiFXzM9KLwolV6A+swQjB6AiITfWLwovLD6zBEMHoEIhN9ovCi8sPrMEYwegYD7bCiEX4i8LB6AiIRfmLwsHoEIhF+o1F7FCNhYT+///B6hhQiE33iFX76FvP//+DfSQBdTP/dSiNhYT+//9Q6PbN//9qfI2FMP///2oAUOiG8///ioUw////g8QMM8BfXluL5V3CJACNRaxQjYWE/v//UOjCzf//i3UojU2si8Ey27oQAAAAK/CQigQOjUkBMkH/CtiD6gF18ItFHITbdT9QV/91GI2FMP///1DoCLz//2p8jYUw////agBQ6Bjz//+KhTD///+DxAwPV8APEUWsikWsX14zwFuL5V3CJACFwHQOUGoAV+jt8v//igeDxAxqfI2FMP///2oAUOjY8v//ioUw////g8QMD1fADxFFrIpFrF9e" & _
                                                    "uAEAAABbi+VdwiQAzMzMzMzMzFWL7FZXi30ID7YHmYvIi/IPtkcBD6TOCJnB4QgL8gvID7ZHAg+kzgiZweEIC/ILyA+2RwMPpM4ImcHhCAvyC8gPtkcED6TOCJnB4QgL8gvID7ZHBQ+kzgiZweEIC/ILyA+2RwYPpM4ImcHhCAvyC8gPtkcHD6TOCJnB4QgLwQvWX15dwgQAzMzMzMzMzMzMzFWL7IPsCItFEEj30JlTi10IiUX4i0UMiVX88w9+XfiNS3hWM/ZmD2zbjVB4O8F3SzvTckcr2MdFEBAAAABXZpCLPBiNQAiLdBj8i0j4i1D8M88jTfgz1iNV/DP5M/KJfBj4iXQY/DFI+DFQ/INtEAF1zl9eW4vlXcIMAIvTjUgQK9APEAzzjUkgDxBR0GYP79FmD9vTDyjCZg/vwQ8RBPODxgQPEEHQZg/v0A8RUdAPEEwK4A8QUeBmD+/RZg/b0w8owmYP78EPEUQK4A8QQeBmD+/CDxFB4IP+EHKlXluL5V3CDADMzMzMzMzMzMzMzFWL7ItVDItFCCvQVr4QAAAAiwwCjUAIiUj4i0wC/IlI/IPuAXXrXl3CCADMzMzMzFWL7ItFEFZXg/gQdDmD+CB1X4t1DIt9CGoQVlfor/D//2oQjUYQUI1HEFDooPD//4PEGOh4of//BSEFAACJRzBfXl3CDACLdQyLfQhqEFZX6Hvw//9qEI1HEFZQ6G/w//+DxBjoR6H//wUQBQAAiUcwX15dwgwAzMzMzMzMzMzMVYvsg+xsi0UIjVWUU1a7kAEAADP2i0gEiU34i0gIiU30i0gMiU3oi0gQiU38i0gUiU3wi0gYiU3si00Mg8ECiXXcV4s4K9OLQByJfeCJReSJ" & _
                                                    "TdiJXQyJVdQPH4AAAAAAg/4QcykPtnH+D7ZB/8HmCAvwD7YBweYIC/APtkEBweYIC/CDwQSJNBqJTdjrVI1eAYPmD41D/YPgD419lI08t4tMhZSLw4PgD4vxwcYPi1SFlIvBwcANM/DB6Qoz8YvCi8rByAfBwQ4zyMHqA41D+DPKi10Mg+APA/EDdIWUAzeJN+hJoP//i338i9fByguLz8HBBzPRi8/ByQb31yN97DPRiwwYg8MEi0XwA8ojRfwDzot14DP4i9aJXQzByg2LxsHACgP5A33kM9CLxsHIAjPQi0X4i8gjxjPOI030M8iLReyJReQD0YtF8ItN+IlF7ItF/IlF8ItF6APHiXX4i3XcA/qLVdRGiUX8i0X0iU30i03YiUXoiX3giXXcgfuQAgAAD4LX/v//i0UIi034i1X8AUgEi030AUgIAVAQATiLTeiLVfABSAwBUBSLVeyLTeQBUBgBSBz/QGBfXluL5V3CCADMzMzMzMzMzMzMzMxVi+yB7OAAAABTVot1CLuQAgAAV4lduIsGiUXsi0YEiUXwi0YMi34IiUXgi0YQiUXUi0YUiUXQi0YYiUW0i0YciUWwi0YgiUXoi0YkiUX0i0YoiUXMi0YsiUXIi0YwiUXEi0Y0iUXAi0Y4iUWsi0Y8i3UMiX3Yjb0g////iUWoM8Ar+4lF3Il9oA8fgAAAAACD+BBzH1boZfv//4vIg8YIi8KJTQyJReSJDB+JRB8E6RMBAACNUAHHRQwAAAAAjUL9g+APi4zFIP///4uExST///+JRfiLwoPgD4lN/I2NIP///4uUxSD///+L+oucxST///+LRdyD4A+JVbzB5xiNBMGLy4lFpIvCD6zICAlFDItFvMHp" & _
                                                    "CAv5i8sPrMgBiX3ki/rR6TPSC9DB5x8xVQwL+YtFvItN5A+s2AczzzFFDItF/MHrBzPLM9uJTeSLTfiL0Q+kwQPB6h3B4AML2YtN+AvQi0X8i/gPrMgTiVW8M9IL0MHpE4tFvDPCwecNi1X8C/mLTfgz3w+sygYzwsHpBotVDDPZi03kA9CLRdwTy4PA+YPgDwOUxSD///8TjMUk////i0WkAxCJVQwTSASJEIlN5IlIBOiUnf//i1X0M/+LTeiL2g+kyhfB6wkL+sHhF4tV9AvZi03oiV38i9kPrNESiX34M/8L+cHqEjF9/DP/i03oweMOC9qLVfQxXfiL2Q+s0Q7B4xIL+cHqDjF9/Avai034i1W4M8uLXfyLfej31wMcEBNMEAQjfcSLVfSLRcj30iNF9CNVwDPQiU34i03MI03oi0X4M/mLTfAD3xPCA10ME0XkA12siV38E0WoM9uJRfiLReyL0A+syBzB4gTB6RwL2ItF7AvRi03wi/kPpMEeiVUMM9LB7wIL0cHgHgv4M98xVQwz0otN8Iv5i0XsD6TBGcHvBwvRweAZMVUMC/iLTdgz34tV4Iv5M33sI33UI03sM1XwM/kjVdCLReAjRfCLTcQz0ItFDAPfi334E8KJTayLTcCLVfwDVbSJTagTfbCLTcwDXfyJTcSLTciJTcCLTeiJTcyLTfSJffSLfdSJfbSLfdCJfbCLfdiJfdSLfeCJfdCLfeyJTciLyBNN+ItF3Ild7ECLXbiJfdiDwwiLffCJfeCLfaCJVeiJTfCJRdyJXbiB+xAFAAAPghv9//+LdQiLReyLfdgBBotF4BFOBIvKAX4Ii320EUYMi0XUAUYQi0XQEUYUAX4Yi0WwEUYcAU4g" & _
                                                    "i0X0EUYki0XMAUYoi0XIEUYsi0XEAUYwi0XAEUY0i02sAU44i02oEU48/4bAAAAAX15bi+VdwggAzMzMzMzMzMzMzMzMzMxVi+xTi10IVlcPtnsHD7ZDCg+2cwsPtlMPwecIC/gPtksDD7ZDDcHnCAv4weYID7YDwecIC/jB4ggPtkMOC/DB4QgPtkMBweYIC/APtkMEweYIC/APtkMCC9APtkMFweIIC9APtkMIweIIC9APtkMGC8iJewQPtkMJweEIC8iJcwgPtkMMweEIXwvIiVMMXokLW13CBADMzMzMzMzMzMzMVYvsi0UMUFD/dQjoAOz//13CCADMzMzMzMzMzMzMzMxVi+yLRRBTVot1CI1IeFeLfQyNVng78XcEO9BzC41PeDvxdzA713IsK/i7EAAAACvwixQ4KxCLTDgEG0gEjUAIiVQw+IlMMPyD6wF15F9eW13CDACL141IEIveK9Ar2Cv+uAQAAACNdiCNSSAPEEHQDxBMN+BmD/vIDxFO4A8QTArgDxBB4GYP+8gPEUwL4IPoAXXSX15bXcIMAMzMzMzMVYvsVugHmv//i3UIBXgFAABQ/zboVwAAAIkG6PCZ//8FeAUAAFD/dgToQgAAAIlGBOjamf//BXgFAABQ/3YI6CwAAACJRgjoxJn//wV4BQAAUP92DOgWAAAAiUYMXl3CBADMzMzMzMzMzMzMzMzMzFWL7ItVDFOLXQiLw8HoGIvLVsHpCA+2yQ+2NBCLw8HoEA+2wA+2DBHB5ggPtgQQC8bB4AgLwQ+2y8HgCF5bD7YMEQvBXcIIAMzMzMzMzMzMVYvsi00MU4tdCFaDwxDHRQwEAAAAV4PBAw8fgAAAAAAPtkH+jVsgmY1JCIvw" & _
                                                    "i/oPtkH1D6T3CJnB5ggD8Ilz0BP6iXvUD7ZB95mL8Iv6D7ZB+JkPpMIIweAIA/CJc9gT+ol73A+2QfqZi/CL+g+2QfkPpPcImcHmCAPwiXPgE/qJe+QPtkH8mYvwi/oPtkH7D6T3CJnB5ggD8Ilz6BP6g20MAYl77A+FdP///4tNCF9eW4FheP9/AADHQXwAAAAAXcIIAMzMzMzMzMzMzMzMzFWL7IPsDFOLXQwPV8BWV4t9EIsTi/KLQwSLyGYPE0X0AzcTTwQ78nUGO8h1BOsYO8h3D3IEO/JzCbgBAAAAM9LrC2YPE0X0i0X0i1X4i30IiU8EiTeLSwiLdRCJTfyLSwyJTfiLTggDTfyJTQiLTgwTTfiLXQgD2IldCBPKO138i10MdQU7Swx0IztLDHcTcgiLQwg5RQhzCbgBAAAAM9LrC2YPE0X0i1X4i0X0i3UIiU8MiXcIi0sQi3UQiU38i0sUiU34i04QA038iU0Ii04UE034i10IA9iJXQgTyjtd/ItdDHUFO0sUdCM7SxR3E3IIi0MQOUUIcwm4AQAAADPS6wtmDxNF9ItV+ItF9It1CIlPFIl3EItLGIt1EIlN/ItLHIlN+ItOGANN/IlNCItOHBNN+ItdCAPYiV0IE8o7XfyLXQx1BTtLHHQjO0scdxNyCItDGDlFCHMJuAEAAAAz0usLZg8TRfSLVfiLRfSLdQiJTxyJdxiLSyCLdRCJTfyLSySJTfiLTiADTfyJTQiLTiQTTfiLXQgD2IldCBPKO138i10MdQU7SyR0IztLJHcTcgiLQyA5RQhzCbgBAAAAM9LrC2YPE0X0i1X4i0X0i3UIiXcgi3UQiU8ki0soi1ssi3YoA/GJTQyLTRCLSSwT" & _
                                                    "ywPwE8o7dQx1BDvLdCw7y3cdcgU7dQxzFol3KLgBAAAAiU8sM9JfXluL5V3CDABmDxNF9ItV+ItF9Il3KIlPLF9eW4vlXcIMAMzMzMzMzFWL7IPsCFOLXQwPV8BWV4t9EIsTi/KLQwSLyGYPE0X4AzcTTwQ78nUGO8h1BOsYO8h3D3IEO/JzCbgBAAAAM9LrC2YPE0X4i0X4i1X8i30IiU8Ei00QiTeLcQgDcwiLSQwTSwwD8BPKO3MIdQU7Swx0IDtLDHcQcgU7cwhzCbgBAAAAM9LrC2YPE0X4i1X8i0X4iU8Mi00QiXcIi3EQA3MQi0kUE0sUA/ATyjtzEHUFO0sUdCA7SxR3EHIFO3MQcwm4AQAAADPS6wtmDxNF+ItV/ItF+IlPFIl3EItLGItbHIlNDItNEItxGAN1DItJHBPLA/ATyjt1DHUEO8t0LDvLdx1yBTt1DHMWiXcYuAEAAACJTxwz0l9eW4vlXcIMAGYPE0X4i1X8i0X4iXcYiU8cX15bi+VdwgwAzMzMzMzMVYvsi00IxwEAAAAAx0EEAAAAAIsBiUEIi0EEiUEMi0EIiUEQi0EMiUEUi0EQiUEYi0EUiUEci0EYiUEgi0EciUEki0EgiUEoi0EkiUEsXcIEAMzMzMzMzMzMzMzMzMzMVYvsi0UIxwAAAAAAx0AEAAAAAMdACAAAAADHQAwAAAAAx0AQAAAAAMdAFAAAAADHQBgAAAAAx0AcAAAAAF3CBADMzMzMzMzMzMzMzMzMzMxVi+yLTQy6BQAAAFOLXQhWK9mNQShXiV0IDx+AAAAAAIs0A4tcAwSLeASLCDvfdy5yIjvxdyg733IadwQ78XIUi10Ig+gIg+oBedVfXjPAW13CCABf"
Private Const STR_THUNK4                As String = "XoPI/1tdwggAX164AQAAAFtdwggAzMzMzMzMVYvsi00MugMAAABTi10IVivZjUEYV4ldCA8fgAAAAACLNAOLXAMEi3gEiwg733cuciI78XcoO99yGncEO/FyFItdCIPoCIPqAXnVX14zwFtdwggAX16DyP9bXcIIAF9euAEAAABbXcIIAMzMzMzMzFWL7ItVCDPADx+EAAAAAACLDMILTMIEdQ9Ag/gGcvG4AQAAAF3CBAAzwF3CBADMzFWL7ItVCDPADx+EAAAAAACLDMILTMIEdQ9Ag/gEcvG4AQAAAF3CBAAzwF3CBADMzFWL7IPsEFOLXRC5QAAAAFaLdQgry1eLfQxmD27DiU0QiweLVwSJRfiJVfzzD35N+GYP88hmD9YO6KOf//+LTRCJRfCLRwiJVfSLVwyJRfiJVfzzD35N+GYPbsNmD/PI8w9+RfBmD+vIZg/WTgjobp///4tNEIlF8ItHEIlV9ItXFIlF+IlV/PMPfk34Zg9uw2YP88jzD35F8GYP68hmD9ZOEOg5n///i00QiUXwi0cYiVX0i1cciUX4iVX88w9+TfhmD27DZg/zyPMPfkXwZg/ryGYP1k4Y6ASf//+LTRCJRfCLRyCJVfSLVySJRfiJVfzzD35N+GYPbsNmD/PI8w9+RfBmD+vIZg/WTiDoz57//4lF8ItHKIlV9ItXLIlF+IlV/PMPfk34i00QZg9uw2YP88jzD35F8GYP68hmD9ZOKOianv//X15bi+VdwgwAzFWL7IPsEFOLXRC5QAAAAFaLdQgry1eLfQxmD27DiU0QiweLVwSJRfiJVfzzD35N+GYP88hmD9YO6FOe//+LTRCJRfCLRwiJVfSLVwyJRfiJVfzzD35N+GYPbsNmD/PI8w9+RfBmD+vI" & _
                                                    "Zg/WTgjoHp7//4tNEIlF8ItHEIlV9ItXFIlF+IlV/PMPfk34Zg9uw2YP88jzD35F8GYP68hmD9ZOEOjpnf//i00QiUXwi0cYiVX0i1cciUX4iVX88w9+TfhmD27DZg/zyPMPfkXwZg/ryGYP1k4Y6LSd//9fXluL5V3CDADMzMzMzMzMzMzMzFWL7IPsaFNWi3UMjV4wU+hM/f//hcAPhWQCAABXDx8AjUWYD1fAUGYPE0X46J/7//+NRchQ6Jb7//9TjUWYUOgs5P//U+iG+///ixaL+gN9mItGBIvIE02cO/p1BjvIdQTrGzvIdw9yBDv6cwm4AQAAADPS6w4PV8BmDxNF+ItF+ItV/IteDIk+i34IA32giU4Ei8sTTaQD+BPKO34IdQQ7y3QiO8t3EHIFO34Icwm4AQAAADPS6w4PV8BmDxNF+ItV/ItF+IteFIl+CIt+EAN9qIlODIvLE02sA/gTyjt+EHUEO8t0IjvLdxByBTt+EHMJuAEAAAAz0usOD1fAZg8TRfiLVfyLRfiLXhyJfhCLfhgDfbCJThSLyxNNtAP4E8o7fhh1BDvLdCI7y3cQcgU7fhhzCbgBAAAAM9LrDg9XwGYPE0X4i1X8i0X4i14kiX4Yi34gA324iU4ci8sTTbwD+BPKO34gdQQ7y3QiO8t3EHIFO34gcwm4AQAAADPS6w4PV8BmDxNF+ItV/ItF+IteLIl+IIt+KAN9wIlOJIvLE03EA/gTyjt+KHUEO8t0IjvLdxByBTt+KHMJuAEAAAAz0usOD1fAZg8TRfiLVfyLRfiLXjSJfiiLfjADfciJTiyLyxNNzAP4E8o7fjB1BDvLdCI7y3cQcgU7fjBzCbgBAAAAM9LrDg9XwGYP" & _
                                                    "E0X4i1X8i0X4i148iX4wi344A33QiU40i8sTTdQD+BPKO344dQQ7y3QiO8t3EHIFO344cwm4AQAAADPS6w4PV8BmDxNF+ItV/ItF+IlOPI1eMItN2APIiX44i0XcE8IBTkBTEUZE6On6//+FwA+Eof3//1/oS47//wWgAAAAUFbo7/n//4XAfifoNo7//wWgAAAAUFZW6DkUAADoJI7//wWgAAAAUFboyPn//4XAf9lW/3UI6DsQAABeW4vlXcIIAMzMzFWL7IPsKFOLXQhWV4t9DFdT6HoQAACLRywPV8CJReSLRzCJReiLRzSJReyLRziJRfCLRzyJRfSNRdhqAVBQZg8TRdjHReAAAAAA6PH7//+L8I1F2FBTU+hk9///i084A/CLRzCLVzyJReQzwAtHNIlF6I1F2GoBUFDHReAAAAAAiU3siVXwx0X0AAAAAOiu+///A/CNRdhQU1PoIff//wPwx0XkAAAAAItHIA9XwIlF2ItHJIlF3ItHKIlF4ItHOIlF8ItHPIlF9I1F2FBTU2YPE0Xo6Of2//+LTyQD8DPAiU3YC0coiUXci0cwi1c0i8qJRfgzwAtHLIlF4ItHOIlF6ItHPIlF7DPAC0cgiUX0jUXYUFNTiU3kiVXw6J/2//+LTywD8ItXNDPAC0cwD1fAiUXci0cgiUXwjUXYiU3YM8kLTyhQU1OJVeDHReQAAAAAZg8TReiJTfTowRQAAItXJCvwi0cwD1fAiUXYsSCLRzSJRdyLRziJReCLRzyJReSLRyBmDxNF6OhCmf//C1csiUXwjUXYUFNTiVX06H4UAACLVQwr8ItPNDPAC0c4i18kiUXcM8ALRzyJTdiLTyAz/4lF4ItCKItSLIlN5LEg" & _
                                                    "6NuY//8L2MdF8AAAAACJXegL+otdDIl97It9CItDMIlF9I1F2FBXV+gjFAAAK/DHReAAAAAAi0M4iUXYi0M8iUXci0MkiUXki0MoiUXoi0MsiUXsi0M0iUX0jUXYUFdXx0XwAAAAAOjkEwAAK/B5IOi7i///UFdX6HP1//8D8HjvX15bi+VdwggAZg8fRAAAhfZ1EVfolov//1DosPf//4P4AXTc6IaL//9QV1fonhMAACvw69rMzMzMzMzMzMzMVYvsVv91EIt1CP91DFbo3fL//wvCdQ3/dRRW6AD3//+FwHgK/3UUVlboUhEAAF5dwhAAzMzMzMzMzMzMzMzMzFWL7Fb/dRCLdQj/dQxW6N30//8LwnUN/3UUVugw9///hcB4Cv91FFZW6CITAABeXcIQAMzMzMzMzMzMzMzMzMxVi+yB7MgAAABWi3UMVuht9///hcB0D/91COjR9f//XovlXcIMAFdWjYU4////UOjsDAAAi30QjYVo////V1Do3AwAAI1FyFDoo/X//41FmMdFyAEAAABQx0XMAAAAAOiM9f//jYVo////UI2FOP///1DoKfb//4vQhdIPhL4BAABTi404////D1fAg+EBZg8TRfiDyQB1L42FOP///1DovAsAAItFyIPgAYPIAA+EvwAAAFeNRchQUOiy8f//i/CL2umxAAAAi4Vo////g+ABg8gAdS+NhWj///9Q6H8LAACLRZiD4AGDyAAPhBEBAABXjUWYUFDodfH//4vwi9rpAwEAAIXSD46PAAAAjYVo////UI2FOP///1BQ6OAPAACNhTj///9Q6DQLAACNRZhQjUXIUOhn9f//hcB5C1eNRchQUOgo8f//jUWYUI1FyFBQ6KoP" & _
                                                    "AACLRciD4AGDyAB0EVeNRchQUOgE8f//i/CL2usGi138i3X4jUXIUOjfCgAAC/MPhJgAAACLRfCBTfQAAACAiUXw6YYAAACNhTj///9QjYVo////UFDoUQ8AAI2FaP///1DopQoAAI1FyFCNRZhQ6Nj0//+FwHkLV41FmFBQ6Jnw//+NRchQjUWYUFDoGw8AAItFmIPgAYPIAHQRV41FmFBQ6HXw//+L8Iva6waLXfyLdfiNRZhQ6FAKAAAL83QNi0XAgU3EAAAAgIlFwI2FaP///1CNhTj///9Q6Gz0//+L0IXSD4VE/v//W41FyFD/dQjo1QoAAF9ei+VdwgwAzMzMzMzMzMzMzMzMzFWL7IHsiAAAAFaLdQxW6D31//+FwHQP/3UI6NHz//9ei+VdwgwAV1aNhXj///9Q6OwKAACLfRCNRZhXUOjfCgAAjUXYUOim8///jUW4x0XYAQAAAFDHRdwAAAAA6I/z//+NRZhQjYV4////UOg/9P//i9CF0g+EsAEAAFMPH0AAi414////D1fAg+EBZg8TRfiDyQB1L42FeP///1DovgkAAItF2IPgAYPIAA+EtgAAAFeNRdhQUOiU8f//i/CL2umoAAAAi0WYg+ABg8gAdSyNRZhQ6IcJAACLRbiD4AGDyAAPhAgBAABXjUW4UFDoXfH//4vwi9rp+gAAAIXSD46MAAAAjUWYUI2FeP///1BQ6JsPAACNhXj///9Q6D8JAACNRbhQjUXYUOiC8///hcB5C1eNRdhQUOgT8f//jUW4UI1F2FBQ6GUPAACLRdiD4AGDyAB0EVeNRdhQUOjv8P//i/CL2usGi138i3X4jUXYUOjqCAAAC/MPhJIAAACLRfCBTfQAAACA" & _
                                                    "iUXw6YAAAACNhXj///9QjUWYUFDoDw8AAI1FmFDotggAAI1F2FCNRbhQ6Pny//+FwHkLV41FuFBQ6Irw//+NRdhQjUW4UFDo3A4AAItFuIPgAYPIAHQRV41FuFBQ6Gbw//+L8Iva6waLXfyLdfiNRbhQ6GEIAAAL83QNi0XQgU3UAAAAgIlF0I1FmFCNhXj///9Q6JDy//+L0IXSD4VW/v//W41F2FD/dQjo6QgAAF9ei+VdwgwAzFWL7IHswAAAAFNWi3UUV1boqwYAAP91EIvYjYVA/////3UMUOjXAwAAjYVw////UOiLBgAAi/iF/3QIgceAAQAA6w6NhUD///9Q6HEGAACL+Dv7cxiNhUD///9Q/3UI6BwIAABfXluL5V3CEACNRaBQ6Nrw//+NRdBQ6NHw//+LxyvDi9jB6waD4D90GFCNRaBWjQTYUOil8v//iUTd0IlU3dTrDY1FoFaNBNhQ6M4HAACLXQhT6JXw///HAwEAAADHQwQAAAAAgf+AAQAAdxJWjUWgUOgm8f//hcAPiIIAAACNhXD///9QjUXQUOgO8f//hcB4FnVIjYVA////UI1FoFDo+PD//4XAfzSNRaBQjYVA////UFDoQwsAAAvCdA5TjYVw////UFDoMQsAAI1F0FCNhXD///9QUOggCwAAi3XQjUXQUMHmH+hxBgAAjUWgUOhoBgAACXXMT4t1FOlk////jYVA////UFPoDwcAAF9eW4vlXcIQAMzMzMzMzFWL7IPsYI1FoP91EP91DFDoawIAAI1FoFD/dQjo3/P//4vlXcIMAMzMzMzMzMzMzFWL7IHsgAAAAFNWi3UUV1boSwUAAP91EIvYjUWA/3UMUOiKAwAAjUWgUOgx" & _
                                                    "BQAAi/iF/3QIgccAAQAA6wuNRYBQ6BoFAACL+Dv7cxWNRYBQ/3UI6NgGAABfXluL5V3CEACNRcBQ6Jbv//+NReBQ6I3v//+LxyvDi9jB6waD4D90GFCNRcBWjQTYUOhR8v//iUTd4IlU3eTrDY1FwFaNBNhQ6IoGAACLXQhT6FHv///HAwEAAADHQwQAAAAADx9AAIH/AAEAAHcOVo1FwFDo7u///4XAeHONRaBQjUXgUOjd7///hcB4E3U8jUWAUI1FwFDoyu///4XAfyuNRcBQjUWAUFDouAsAAAvCdAtTjUWgUFDoqQsAAI1F4FCNRaBQUOibCwAAi3XgjUXgUMHmH+g8BQAAjUXAUOgzBQAACXXcT4t1FOl3////jUWAUFPo3QUAAF9eW4vlXcIQAMzMzMxVi+yD7ECNRcD/dRD/dQxQ6DsCAACNRcBQ/3UI6B/1//+L5V3CDADMzMzMzMzMzMxVi+yD7GCNRaD/dQxQ6M4FAACNRaBQ/3UI6CLy//+L5V3CCADMzMzMzMzMzMzMzMxVi+yD7ECNRcD/dQxQ6D4HAACNRcBQ/3UI6ML0//+L5V3CCADMzMzMzMzMzMzMzMxVi+xW/3UQi3UI/3UMVuitCAAAC8J0Cv91FFZW6A/q//9eXcIQAMzMzMzMzMzMzMxVi+xW/3UQi3UI/3UMVuiNCgAAC8J0Cv91FFZW6B/s//9eXcIQAMzMzMzMzMzMzMxVi+yD7GBTD1fAVmYPE0XYi0XcV2YPE0XQM/+LXdSJRfwz9o1H+4P/Bg9XwGYPE0X0i1X0D0PwO/cPh9IAAACLTRCLxw8QRdArxg8RRcCNHMGLRfiJRfCJVfhmDx9EAACD/gYPg6MAAAD/cwSLRQz/" & _
                                                    "M/908AT/NPCNRbBQ6M/U//+D7BCLzIPsEA8QAA8QCIvEDxEBDxBFwA8RTeAPEQCNRaBQ6DiP//9mD3PZDA8QEGYPfsgPKMJmD3PYDGYPfsEPEVXAiU38DxFV0DvIdxNyCItF2DtF6HMJuAEAAAAzyesOD1fAZg8TReiLTeyLReiLVfgD0ItF8IlV+BPBRoPrCIlF8Dv3D4ZU////i13U6wOLRfiLTQiLddCJNPmL8YvKi9CJVdyJXP4ER4t12Itd/Il10Ild1IlN2IlV/IP/Cw+C2/7//4tFCF+JcFheiVhcW4vlXcIMAMzMzMzMzMzMVYvsg+xgUw9XwFZmDxNF2ItF3FdmDxNF0DP/i13UiUX8M/aNR/2D/wQPV8BmDxNF9ItV9A9D8Dv3D4fSAAAAi00Qi8cPEEXQK8YPEUXAjRzBi0X4iUXwiVX4Zg8fRAAAg/4ED4OjAAAA/3MEi0UM/zP/dPAE/zTwjUWwUOhv0///g+wQi8yD7BAPEAAPEAiLxA8RAQ8QRcAPEU3gDxEAjUWgUOjYjf//Zg9z2QwPEBBmD37IDyjCZg9z2AxmD37BDxFVwIlN/A8RVdA7yHcTcgiLRdg7RehzCbgBAAAAM8nrDg9XwGYPE0Xoi03si0Xoi1X4A9CLRfCJVfgTwUaD6wiJRfA79w+GVP///4td1OsDi0X4i00Ii3XQiTT5i/GLyovQiVXciVz+BEeLddiLXfyJddCJXdSJTdiJVfyD/wcPgtv+//+LRQhfiXA4XolYPFuL5V3CDADMzMzMzMzMzFWL7FZXi30IV+iSAAAAi/CF9nUGX15dwgQAi1T3+IvKi0T3/DP/C8h0E2YPH0QAAA+swgFH0eiLygvIdfPB5gaNRsAD" & _
                                                    "x19eXcIEAMzMzMzMVYvsVleLfQhX6HIAAACL8IX2dQZfXl3CBACLVPf4i8qLRPf8M/8LyHQTZg8fRAAAD6zCAUfR6IvKC8h188HmBo1GwAPHX15dwgQAzMzMzMxVi+yLVQi4BQAAAA8fRAAAiwzCC0zCBHUFg+gBefJAXcIEAMzMzMzMzMzMzMzMzMxVi+yLVQi4AwAAAA8fRAAAiwzCC0zCBHUFg+gBefJAXcIEAMzMzMzMzMzMzMzMzMxVi+yD7AiLRQgPV8BTi9hmDxNF+IPAMDvDdjiLTfhWV4t9/IlNCItw+IPoCIvOi1AED6zRAQtNCNHqC9eJCIv+iVAEwecfx0UIAAAAADvDd9VfXluL5V3CBADMzMzMzMxVi+yD7AiLRQgPV8BTi9hmDxNF+IPAIDvDdjiLTfhWV4t9/IlNCItw+IPoCIvOi1AED6zRAQtNCNHqC9eJCIv+iVAEwecfx0UIAAAAADvDd9VfXluL5V3CBADMzMzMzMxVi+yLVQyLTQiLAokBi0IEiUEEi0IIiUEIi0IMiUEMi0IQiUEQi0IUiUEUi0IYiUEYi0IciUEci0IgiUEgi0IkiUEki0IoiUEoi0IsiUEsXcIIAMzMzMzMzMzMzMzMzMxVi+yLVQyLTQiLAokBi0IEiUEEi0IIiUEIi0IMiUEMi0IQiUEQi0IUiUEUi0IYiUEYi0IciUEcXcIIAMzMzMzMVYvsg+xgUw9XwDPJVmYPE0XYi0XcV2YPE0XQi33UiU3oiUXwM/aNQfuD+QYPV8BmDxNF+Itd/A9D8DvxD4cZAQAAi1UMi8EPEEXQK8aJXfQPEUXAjQTCi1X4iUXsiVX8i/kr/jv3D4fqAAAA/3AE/zCLRQz/dPAE" & _
                                                    "/zTwjUWwUOisz///DxAADxFF0Dv3c0OLTdyLwYtV1Iv6wegfAUX8i0XYg9MAwe8fD6TBAYld9DPbA8AL2Qv4iV3ci0XQD6TCAYl92APAiVXUiUXQDxBF0OsGi13ci33Yg+wQi8SD7BAPEQCLxA8QRcAPEQCNRaBQ6MuJ//8PEAgPKMFmD3PYDGYPfsAPEU3AiUXwDxFN0DvDdxByBTl92HMJuAEAAAAzyesOD1fAZg8TReCLTeSLReCLVfyLXfQD0ItF7BPZiVX8i03oRoPoCIld9IlF7DvxD4YK////i33U6wOLVfiLdQiLRdCJBM6LRdiJfM4EQYt98IlV2IvTiUXQiX3UiVXwiVXciU3og/kLD4KV/v//iX5cX4lGWF5bi+VdwggAzMxVi+yD7GBTD1fAM8lWZg8TRdiLRdxXZg8TRdCLfdSJTeiJRfAz9o1B/YP5BA9XwGYPE0X4i138D0PwO/EPhxkBAACLVQyLwQ8QRdArxold9A8RRcCNBMKLVfiJReyJVfyL+Sv+O/cPh+oAAAD/cAT/MItFDP908AT/NPCNRbBQ6AzO//8PEAAPEUXQO/dzQ4tN3IvBi1XUi/rB6B8BRfyLRdiD0wDB7x8PpMEBiV30M9sDwAvZC/iJXdyLRdAPpMIBiX3YA8CJVdSJRdAPEEXQ6waLXdyLfdiD7BCLxIPsEA8RAIvEDxBFwA8RAI1FoFDoK4j//w8QCA8owWYPc9gMZg9+wA8RTcCJRfAPEU3QO8N3EHIFOX3Ycwm4AQAAADPJ6w4PV8BmDxNF4ItN5ItF4ItV/Itd9APQi0XsE9mJVfyLTehGg+gIiV30iUXsO/EPhgr///+LfdTrA4tV+It1CItF0IkEzotF2Il8" & _
                                                    "zgRBi33wiVXYi9OJRdCJfdSJVfCJVdyJTeiD+QcPgpX+//+JfjxfiUY4XluL5V3CCADMzFWL7IPsDFOLXQwPV8BWV4t9EIsTi/KLQwSLyGYPE0X0KzcbTwQ78nUGO8h1BOsYO8hyD3cEO/J2CbgBAAAAM9LrC2YPE0X0i0X0i1X4i30IiU8EiTeLcwiLzol1+It1ECtOCIlNCItLDBtODItdCCvYiV0IG8o7XfiLXQx1BTtLDHQjO0sMchN3CItDCDlFCHYJuAEAAAAz0usLZg8TRfSLVfiLRfSLdQiJTwyJdwiLcxCLzol1/It1ECtOEIlNCItLFBtOFItdCCvYiV0IG8o7XfyLXQx1BTtLFHQjO0sUchN3CItDEDlFCHYJuAEAAAAz0usLZg8TRfSLVfiLRfSLdQiJTxSJdxCLcxiLzol1/It1ECtOGIlNCItLHBtOHItdCCvYiV0IG8o7XfyLXQx1BTtLHHQjO0scchN3CItDGDlFCHYJuAEAAAAz0usLZg8TRfSLVfiLRfSLdQiJdxiLdRCJTxyLSyArTiCJTQyLSyQbTiSLdQwr8BvKO3MgdQU7SyR0IDtLJHIQdwU7cyB2CbgBAAAAM9LrC2YPE0X0i1X4i0X0iXcgiU8ki3Moi0ssi10QiXUIiU0MK3MoG0ssK/CLXQwbyjt1CHUEO8t0LDvLch13BTt1CHYWiXcouAEAAACJTywz0l9eW4vlXcIMAGYPE0X0i1X4i0X0iXcoiU8sX15bi+VdwgwAzMzMzFWL7IPsDFOLXQwPV8BWV4t9EIsTi/KLQwSLyGYPE0X0KzcbTwQ78nUGO8h1BOsYO8hyD3cEO/J2CbgBAAAAM9LrC2YPE0X0i0X0i1X4i30I" & _
                                                    "iU8Ei00QiTeLcwiJdfgrcQiLSwyLXRAbSwwr8ItdDBvKO3X4dQU7Swx0IDtLDHIQdwU7cwh2CbgBAAAAM9LrC2YPE0X0i1X4i0X0iU8Mi00QiXcIi3MQiXX8K3EQi0sUi10QG0sUK/CLXQwbyjt1/HUFO0sUdCA7SxRyEHcFO3MQdgm4AQAAADPS6wtmDxNF9ItV+ItF9IlPFIl3EItLGIvxi30Qi1sciU0Mi00QK3EYi8sbTxwr8It9CBvKO3UMdQQ7y3QsO8tyHXcFO3UMdhaJdxi4AQAAAIlPHDPSX15bi+VdwgwAZg8TRfSLVfiLRfSJdxiJTxxfXluL5V3CDADMzMzMzMzMzMzMzMzMzMxVi+yLTQgz0lZXi30MM/aLx4PgPw+rxoP4IA9D1jPyg/hAD0PWwe8GIzT5I1T5BIvGX15dwggAzMzMzMzMzMzMVYvsi1UUg+wQM8mF0g+EwgAAAFOLXRBWi3UIV4t9DIP6IA+CiwAAAI1D/wPCO/B3CY1G/wPCO8NzeY1H/wPCO/B3CY1G/wPCO8dzZ4vCi9cr04Pg4IlV/IvWK9OJRfCJVfiLw4td+IvXi338K9aJVfSNVhAPEACLdfSDwSCNQCCNUiAPEEwH4GYP78gPEUwD4A8QTBbgi3UIDxBA8GYP78gPEUrgO03wcsqLVRSLfQyLXRA7ynMbK/uNBBkr8yvRigw4jUABMkj/iEww/4PqAXXuX15bi+VdwhAAAAA=" ' 35597, 24.4.2020 15:03:56
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
            Call pvPatchTrampoline(AddressOf pvCryptoCallSecpMakeKey)
            Call pvPatchTrampoline(AddressOf pvCryptoCallSecpSharedSecret)
            Call pvPatchTrampoline(AddressOf pvCryptoCallSecpUncompressKey)
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
    Case ucsTlsAlgoSignaturePss
        CryptoIsSupported = (OsVersion >= ucsOsvVista)  '--- need BCrypt for PSS padding on signatures
    Case ucsTlsAlgoSignaturePkcsSha2
        CryptoIsSupported = (OsVersion >= ucsOsvXp)     '--- need PROV_RSA_AES for SHA-2
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
        If pvCryptoCallSecpMakeKey(m_uData.Pfn(ucsPfnSecp256r1MakeKey), baPublic(0), baPrivate(0)) = 1 Then
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
    If pvCryptoCallSecpSharedSecret(m_uData.Pfn(ucsPfnSecp256r1SharedSecret), baPublic(0), baPrivate(0), baRetVal(0)) = 0 Then
        GoTo QH
    End If
    CryptoEccSecp256r1SharedSecret = baRetVal
QH:
End Function

Public Function CryptoEccSecp256r1UncompressKey(baPublic() As Byte) As Byte()
    Dim baRetVal()      As Byte
    
    ReDim baRetVal(0 To 2 * m_uData.Ecc256KeySize) As Byte
    If pvCryptoCallSecpUncompressKey(m_uData.Pfn(ucsPfnSecp256r1UncompressKey), baPublic(0), baRetVal(0)) = 0 Then
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
    CryptoEccSecp256r1Verify = (pvCryptoCallSecpVerify(m_uData.Pfn(ucsPfnSecp256r1Verify), baPublic(0), baHash(0), baSignature(0)) <> 0)
End Function

Public Function CryptoEccSecp384r1MakeKey(baPrivate() As Byte, baPublic() As Byte) As Boolean
    Const MAX_RETRIES   As Long = 16
    Dim lIdx            As Long
    
    ReDim baPrivate(0 To m_uData.Ecc384KeySize - 1) As Byte
    ReDim baPublic(0 To m_uData.Ecc384KeySize) As Byte
    For lIdx = 1 To MAX_RETRIES
        CryptoRandomBytes VarPtr(baPrivate(0)), m_uData.Ecc384KeySize
        If pvCryptoCallSecpMakeKey(m_uData.Pfn(ucsPfnSecp384r1MakeKey), baPublic(0), baPrivate(0)) = 1 Then
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
    If pvCryptoCallSecpSharedSecret(m_uData.Pfn(ucsPfnSecp384r1SharedSecret), baPublic(0), baPrivate(0), baRetVal(0)) = 0 Then
        GoTo QH
    End If
    CryptoEccSecp384r1SharedSecret = baRetVal
QH:
End Function

Public Function CryptoEccSecp384r1UncompressKey(baPublic() As Byte) As Byte()
    Dim baRetVal()      As Byte
    
    ReDim baRetVal(0 To 2 * m_uData.Ecc384KeySize) As Byte
    If pvCryptoCallSecpUncompressKey(m_uData.Pfn(ucsPfnSecp384r1UncompressKey), baPublic(0), baRetVal(0)) = 0 Then
        GoTo QH
    End If
    CryptoEccSecp384r1UncompressKey = baRetVal
QH:
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

Private Function pvCryptoCallSecpMakeKey(ByVal Pfn As Long, pPubKeyPtr As Byte, pPrivKeyPtr As Byte) As Long
    ' int ecc_make_key(uint8_t p_publicKey[ECC_BYTES+1], uint8_t p_privateKey[ECC_BYTES]);
    ' int ecc_make_key384(uint8_t p_publicKey[ECC_BYTES_384+1], const uint8_t p_privateKey[ECC_BYTES_384])
End Function

Private Function pvCryptoCallSecpSharedSecret(ByVal Pfn As Long, pPubKeyPtr As Byte, pPrivKeyPtr As Byte, pSecretPtr As Byte) As Long
    ' int ecdh_shared_secret(const uint8_t p_publicKey[ECC_BYTES+1], const uint8_t p_privateKey[ECC_BYTES], uint8_t p_secret[ECC_BYTES]);
    ' int ecdh_shared_secret384(const uint8_t p_publicKey[ECC_BYTES_384+1], const uint8_t p_privateKey[ECC_BYTES_384], uint8_t p_secret[ECC_BYTES_384])
End Function

Private Function pvCryptoCallSecpUncompressKey(ByVal Pfn As Long, pPubKeyPtr As Byte, pUncompressedKeyPtr As Byte) As Long
    ' int ecdh_uncompress_key(const uint8_t p_publicKey[ECC_BYTES + 1], uint8_t p_uncompressedKey[2 * ECC_BYTES + 1])
    ' int ecdh_uncompress_key384(const uint8_t p_publicKey[ECC_BYTES_384 + 1], uint8_t p_uncompressedKey[2 * ECC_BYTES_384 + 1])
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
