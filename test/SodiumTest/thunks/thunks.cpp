//#pragma nodefaultlib
//#pragma comment(linker, "/entry:main")
//#pragma comment(linker, "/INCLUDE:_mainCRTStartup")

#pragma intrinsic(memset, memcpy)

#define IMPL_ECC_THUNK

#include <stdio.h>
#include <windows.h>
#include <commctrl.h>

#pragma comment(lib, "crypt32")
#pragma comment(lib, "comctl32")

#pragma code_seg(".supcode")

LPWSTR __stdcall GetCurrentDateTime()
{
    static WCHAR szResult[50];
    SYSTEMTIME  st;
    DATE        dt;
    VARIANT     vdt = { VT_DATE, };
    VARIANT     vstr = { VT_EMPTY };

    GetLocalTime(&st);
    SystemTimeToVariantTime(&st, &dt);
    vdt.date = dt;
    VariantChangeType(&vstr, &vdt, 0, VT_BSTR);
    memcpy(szResult, vstr.bstrVal, sizeof szResult);
    VariantClear(&vstr);
    return szResult;
}


#ifdef IMPL_ECC_THUNK

#define ECC_NO_SIGN
#define assert(expr)
#define MIN(x, y) ((x) < (y) ? (x) : (y))

#include "ecc.h"

typedef struct {
    uint64_t m_curve_p[NUM_ECC_DIGITS];
    uint64_t m_curve_b[NUM_ECC_DIGITS];
    EccPoint m_curve_G;
    uint64_t m_curve_n[NUM_ECC_DIGITS];
    uint32_t m_K256[64];
    uint64_t m_K512[80];
    uint8_t m_chacha20_tau[17];  // "expand 16-byte k";
    uint8_t m_chacha20_sigma[17]; // "expand 32-byte k";
    uint32_t m_negative_1305[17];
} thunk_context_t;

#define getRandomNumber (getContext()->m_getRandomNumber)
#define curve_p (getContext()->m_curve_p)
#define curve_b (getContext()->m_curve_b)
#define curve_G (getContext()->m_curve_G)
#define curve_n (getContext()->m_curve_n)
#define K256 (getContext()->m_K256)
#define K512 (getContext()->m_K512)
#define chacha20_tau (getContext()->m_chacha20_tau)
#define chacha20_sigma (getContext()->m_chacha20_sigma)
#define negative_1305 (getContext()->m_negative_1305)

#pragma code_seg(push, r1, ".mythunk")

int beginOfThunk(int i) { 
    int a[] = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 }; return a[i]; 
}

__declspec(naked) thunk_context_t *getContext() {
    __asm {
        call    _next
_next:
        pop     eax
        sub     eax, 5 + 16*4
        mov     eax, [eax]
        ret
    }
}

__declspec(naked) char *getThunk() {
    __asm {
        call    _next
_next:
        pop     eax
        sub     eax, 5 + 16 + 16*4
        ret
    }
}


#ifdef __cplusplus
extern "C" {
#endif

#include "win32_crt.cpp"
#include "ecc.c"
#include "curve25519.c"
#include "blockwise.c"
#include "sha256.c"
#include "sha512.c"
#include "chacha20.c"
#include "poly1305.c"
#include "chacha20poly1305.c"

#ifdef __cplusplus
}
#endif

#pragma code_seg(pop, r1)

#pragma code_seg(push, r1, ".endthunk")
static int endOfThunk() { return 0; }
#pragma code_seg(pop, r1)


#define THUNK_SIZE (((char *)endOfThunk - (char *)beginOfThunk))
//(((char *)vli_sub - (char *)beginOfThunk) + 352)
// 0x3000

static int _getRandomNumber(uint64_t *p_vli);

typedef int (*ecc_make_key_t)(uint8_t p_publicKey[ECC_BYTES+1], const uint8_t p_privateKey[ECC_BYTES]);
typedef int (*ecdh_shared_secret_t)(const uint8_t p_publicKey[ECC_BYTES+1], const uint8_t p_privateKey[ECC_BYTES], uint8_t p_secret[ECC_BYTES]);
typedef void (*cf_curve25519_mul_t)(uint8_t *q, const uint8_t *n, const uint8_t *p);
typedef void (*cf_curve25519_mul_base_t)(uint8_t *q, const uint8_t *n);
typedef void (*cf_sha256_init_t)(cf_sha256_context *ctx);
typedef void (*cf_sha256_update_t)(cf_sha256_context *ctx, const void *data, size_t nbytes);
typedef void (*cf_sha256_digest_final_t)(cf_sha256_context *ctx, uint8_t hash[CF_SHA256_HASHSZ]);
typedef void (*cf_sha384_init_t)(cf_sha512_context *ctx);
typedef void (*cf_sha384_update_t)(cf_sha512_context *ctx, const void *data, size_t nbytes);
typedef void (*cf_sha384_digest_final_t)(cf_sha512_context *ctx, uint8_t hash[CF_SHA384_HASHSZ]);
typedef void (*cf_chacha20poly1305_encrypt_t)(const uint8_t key[32],
                                              const uint8_t nonce[12],
                                              const uint8_t *header, size_t nheader,
                                              const uint8_t *plaintext, size_t nbytes,
                                              uint8_t *ciphertext,
                                              uint8_t tag[16]);
typedef int (*cf_chacha20poly1305_decrypt_t)(const uint8_t key[32],
                                             const uint8_t nonce[12],
                                             const uint8_t *header, size_t nheader,
                                             const uint8_t *ciphertext, size_t nbytes,
                                             const uint8_t tag[16],
                                             uint8_t *plaintext);

void __cdecl main()
{
    printf("sizeof(cf_sha256_context)=%d\n", sizeof cf_sha256_context);
    printf("sizeof(cf_sha512_context)=%d\n", sizeof cf_sha512_context);
    printf("sizeof(_chacha20_tau)=%d\n", sizeof _chacha20_tau);
    static thunk_context_t ctx;
    memcpy(&ctx.m_curve_p, &_curve_p, sizeof _curve_p);
    memcpy(&ctx.m_curve_b, &_curve_b, sizeof _curve_b);
    memcpy(&ctx.m_curve_G, &_curve_G, sizeof _curve_G);
    memcpy(&ctx.m_curve_n, &_curve_n, sizeof _curve_p);
    memcpy(&ctx.m_K256, &_K256, sizeof _K256);
    memcpy(&ctx.m_K512, &_K512, sizeof _K512);
    memcpy(&ctx.m_chacha20_tau, &_chacha20_tau, sizeof _chacha20_tau);
    memcpy(&ctx.m_chacha20_sigma, &_chacha20_sigma, sizeof _chacha20_sigma);
    memcpy(&ctx.m_negative_1305, &_negative_1305, sizeof _negative_1305);

    CoInitialize(0);
    DWORD dwDummy;
    VirtualProtect(beginOfThunk, 1024, PAGE_EXECUTE_READWRITE, &dwDummy);
    ((void **)beginOfThunk)[0] = &ctx;

    size_t thunkSize = THUNK_SIZE;
    while(thunkSize > 4 && ((char *)beginOfThunk)[thunkSize - 4] == 0)
        thunkSize--;
    void *hThunk = VirtualAlloc(0, 2*THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE);
    printf("hThunk=%p\nTHUNK_SIZE=%d -> %d\n", hThunk, THUNK_SIZE, thunkSize);
    memcpy(hThunk, beginOfThunk, THUNK_SIZE);
    memset(((char *)hThunk) + thunkSize, 0xCC, 2*THUNK_SIZE - thunkSize);

    // tests here
    ecc_make_key_t pfn_ecc_make_key = (ecc_make_key_t)(((char *)hThunk) + ((char *)ecc_make_key - (char *)beginOfThunk));
    ecdh_shared_secret_t pfn_ecdh_shared_secret = (ecdh_shared_secret_t)(((char *)hThunk) + ((char *)ecdh_shared_secret - (char *)beginOfThunk));
    cf_curve25519_mul_t pfn_cf_curve25519_mul = (cf_curve25519_mul_t)(((char *)hThunk) + ((char *)cf_curve25519_mul - (char *)beginOfThunk));
    cf_sha256_init_t pfn_cf_sha256_init = (cf_sha256_init_t)(((char *)hThunk) + ((char *)cf_sha256_init - (char *)beginOfThunk));
    cf_sha256_update_t pfn_cf_sha256_update = (cf_sha256_update_t)(((char *)hThunk) + ((char *)cf_sha256_update - (char *)beginOfThunk));
    cf_sha256_digest_final_t pfn_cf_sha256_digest_final = (cf_sha256_digest_final_t)(((char *)hThunk) + ((char *)cf_sha256_digest_final - (char *)beginOfThunk));

    cf_sha384_init_t pfn_cf_sha384_init = (cf_sha384_init_t)(((char *)hThunk) + ((char *)cf_sha384_init - (char *)beginOfThunk));
    cf_sha384_update_t pfn_cf_sha384_update = (cf_sha384_update_t)(((char *)hThunk) + ((char *)cf_sha384_update - (char *)beginOfThunk));
    cf_sha384_digest_final_t pfn_cf_sha384_digest_final = (cf_sha384_digest_final_t)(((char *)hThunk) + ((char *)cf_sha384_digest_final - (char *)beginOfThunk));

    cf_chacha20poly1305_encrypt_t pfn_cf_chacha20poly1305_encrypt = (cf_chacha20poly1305_encrypt_t)(((char *)hThunk) + ((char *)cf_chacha20poly1305_encrypt - (char *)beginOfThunk));
    cf_chacha20poly1305_decrypt_t pfn_cf_chacha20poly1305_decrypt = (cf_chacha20poly1305_decrypt_t)(((char *)hThunk) + ((char *)cf_chacha20poly1305_decrypt - (char *)beginOfThunk));

    uint8_t pubkey[ECC_BYTES+1] = { 0 };
    uint8_t privkey[ECC_BYTES] = { 0 };
    uint8_t secret[ECC_BYTES] = { 0 };
    do {
        _getRandomNumber((uint64_t *)privkey);
    } while (!ecc_make_key(pubkey, privkey));
    pfn_ecc_make_key(pubkey, privkey);
    ecdh_shared_secret(pubkey, privkey, secret);
    pfn_ecdh_shared_secret(pubkey, privkey, secret);

    cf_curve25519_mul(secret, privkey, pubkey);
    pfn_cf_curve25519_mul(secret, privkey, pubkey);

    cf_sha256_context sha256_ctx = { 0 };
    pfn_cf_sha256_init(&sha256_ctx);
    pfn_cf_sha256_update(&sha256_ctx, "123456", 6);
    uint8_t hash256[32] = { 0 };
    pfn_cf_sha256_digest_final(&sha256_ctx, hash256);

    cf_sha512_context sha512_ctx = { 0 };
    pfn_cf_sha384_init(&sha512_ctx);
    pfn_cf_sha384_update(&sha512_ctx, "123456", 6);
    uint8_t hash384[32] = { 0 };
    pfn_cf_sha384_digest_final(&sha512_ctx, hash384);

    uint8_t key[32] = { 1, 2, 3, 4 };
    uint8_t nonce[12] = { 1, 2, 3, 4 };
    uint8_t tag[16];
    uint8_t aad[] = "header text";
    uint8_t plaintext[] = "this is a test 1234567890";
    uint8_t cyphertext[100] = { 0 };

    pfn_cf_chacha20poly1305_encrypt(key, nonce, aad, sizeof aad, plaintext, sizeof plaintext, cyphertext, tag);
    pfn_cf_chacha20poly1305_decrypt(key, nonce, aad, sizeof aad, cyphertext, sizeof plaintext, tag, cyphertext);

    // init offsets
    int i = 0;
    ((int *)hThunk)[i++] = ((char *)ecc_make_key - (char *)beginOfThunk);
    ((int *)hThunk)[i++] = ((char *)ecdh_shared_secret - (char *)beginOfThunk);
#ifndef ECC_NO_SIGN
    ((int *)hThunk)[i++] = ((char *)ecdsa_sign - (char *)beginOfThunk);
    ((int *)hThunk)[i++] = ((char *)ecdsa_verify - (char *)beginOfThunk);
#endif
    ((int *)hThunk)[i++] = ((char *)cf_curve25519_mul - (char *)beginOfThunk);
    ((int *)hThunk)[i++] = ((char *)cf_curve25519_mul_base - (char *)beginOfThunk);

    ((int *)hThunk)[i++] = ((char *)cf_sha256_init - (char *)beginOfThunk);
    ((int *)hThunk)[i++] = ((char *)cf_sha256_update - (char *)beginOfThunk);
    ((int *)hThunk)[i++] = ((char *)cf_sha256_digest_final - (char *)beginOfThunk);
    ((int *)hThunk)[i++] = ((char *)cf_sha384_init - (char *)beginOfThunk);
    ((int *)hThunk)[i++] = ((char *)cf_sha384_update - (char *)beginOfThunk);
    ((int *)hThunk)[i++] = ((char *)cf_sha384_digest_final - (char *)beginOfThunk);

    ((int *)hThunk)[i++] = ((char *)cf_chacha20poly1305_encrypt - (char *)beginOfThunk);
    ((int *)hThunk)[i++] = ((char *)cf_chacha20poly1305_decrypt - (char *)beginOfThunk);
    

    WCHAR szBuffer[50000] = { 0 };
    DWORD dwBufSize = _countof(szBuffer);
    CryptBinaryToString((BYTE *)&ctx, sizeof ctx, CRYPT_STRING_BASE64, szBuffer, &dwBufSize);
    for(int i = 0, j = 0; (szBuffer[j] = szBuffer[i]) != 0; ) {
        ++i, j += (szBuffer[j] != '\r' && szBuffer[j] != '\n');
        if (j % 900 == 0) {
            memcpy(szBuffer + j, L"\" & _\n\"", 14);
            j += 7;
        }
    }
    printf("Private Const STR_GLOB                  As String = \"%S\" ' %d, %S\n", szBuffer, sizeof ctx, GetCurrentDateTime());
    dwBufSize = _countof(szBuffer);
    CryptBinaryToString((BYTE *)hThunk, thunkSize, CRYPT_STRING_BASE64, szBuffer, &dwBufSize);
    for(int i = 0, j = 0; (szBuffer[j] = szBuffer[i]) != 0; ) {
        ++i, j += (szBuffer[j] != '\r' && szBuffer[j] != '\n');
        if (j % 900 == 0) {
            memcpy(szBuffer + j, L"\" & _\n\"", 14);
            j += 7;
        }
    }
    printf("Private Const STR_THUNK1                As String = \"%S\" ' %d, %S\n", szBuffer, thunkSize, GetCurrentDateTime());
}

//#define WIN32_LEAN_AND_MEAN
//#include <windows.h>
//#include <wincrypt.h>


static int _getRandomNumber(uint64_t *p_vli)
{
    HCRYPTPROV l_prov;

    if(!CryptAcquireContext(&l_prov, NULL, NULL, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT))
    {
        return 0;
    }

    CryptGenRandom(l_prov, ECC_BYTES, (BYTE *)p_vli);
    CryptReleaseContext(l_prov, 0);
    
    return 1;
}

#endif
