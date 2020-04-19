#define IMPL_ECC_THUNK
#define IMPL_SHA_THUNK
#define IMPL_CHACHA20_THUNK
#define IMPL_AESGCM_THUNK

#include <stdio.h>
#include <string.h>
#include <stdint.h>
#include <windows.h>

#pragma intrinsic(memset, memcpy)
#pragma comment(lib, "crypt32")
#pragma comment(lib, "comctl32")
//#pragma nodefaultlib
//#pragma comment(linker, "/entry:main")
//#pragma comment(linker, "/INCLUDE:_mainCRTStartup")

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


#define assert(expr)
#define MIN(x, y) ((x) < (y) ? (x) : (y))

#ifdef IMPL_ECC_THUNK
    #define ECC_NO_SIGN
    #include "ecc.h"
#endif

typedef struct {
#ifdef IMPL_ECC_THUNK
    uint64_t m_curve_p[NUM_ECC_DIGITS];
    uint64_t m_curve_b[NUM_ECC_DIGITS];
    EccPoint m_curve_G;
    uint64_t m_curve_n[NUM_ECC_DIGITS];
#endif
#ifdef IMPL_SHA_THUNK
    uint32_t m_K256[64];
    uint64_t m_K512[80];
#endif
#ifdef IMPL_CHACHA20_THUNK
    uint8_t m_chacha20_tau[17];  // "expand 16-byte k";
    uint8_t m_chacha20_sigma[17]; // "expand 32-byte k";
    uint32_t m_negative_1305[17];
#endif
#ifdef IMPL_AESGCM_THUNK
    uint8_t m_S[256];
    uint8_t m_Rcon[11];
    uint8_t m_S_inv[256];
#endif
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
#define S (getContext()->m_S)
#define Rcon (getContext()->m_Rcon)
#define S_inv (getContext()->m_S_inv)

#pragma code_seg(push, r1, ".mythunk")

int beginOfThunk(int i) { 
    int a[] = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 }; return a[i]; 
}

__declspec(naked) thunk_context_t *getContext() {
    __asm {
        call    _next
_next:
        pop     eax
        sub     eax, 5 + getContext
        add     eax, beginOfThunk
        mov     eax, [eax]
        ret
    }
}

__declspec(naked) uint8_t *getThunk() {
    __asm {
        call    _next
_next:
        pop     eax
        sub     eax, 5 + getThunk
        add     eax, beginOfThunk
        ret
    }
}

#define DECLARE_PFN(t, f) const t pfn_##f = (t)(getThunk() + (((uint8_t *)f) - ((uint8_t *)beginOfThunk)));

#ifdef __cplusplus
extern "C" {
#endif

#include "cf_inlines.h"
#include "win32_crt.cpp"
#ifdef IMPL_ECC_THUNK
    #include "ecc.c"
    #include "curve25519.c"
#endif
#include "blockwise.c"
#ifdef IMPL_SHA_THUNK
    #include "sha256.c"
    #include "sha512.c"
#endif
#ifdef IMPL_CHACHA20_THUNK
    #include "chacha20.c"
    #include "poly1305.c"
    #include "chacha20poly1305.c"
#endif
#ifdef IMPL_AESGCM_THUNK
    #include "aes.c"
    #include "gf128.c"
    #include "modes.c"
    #include "gcm.c"
#endif

#ifdef __cplusplus
}
#endif

#pragma code_seg(pop, r1)

#pragma code_seg(push, r1, ".endthunk")
static int endOfThunk() { return 0; }
#pragma code_seg(pop, r1)


#define THUNK_SIZE (((uint8_t *)endOfThunk - (uint8_t *)beginOfThunk))

#ifdef IMPL_ECC_THUNK
    static int _getRandomNumber(uint64_t *p_vli);
    typedef int (*ecc_make_key_t)(uint8_t p_publicKey[ECC_BYTES+1], const uint8_t p_privateKey[ECC_BYTES]);
    typedef int (*ecdh_shared_secret_t)(const uint8_t p_publicKey[ECC_BYTES+1], const uint8_t p_privateKey[ECC_BYTES], uint8_t p_secret[ECC_BYTES]);
    typedef void (*cf_curve25519_mul_t)(uint8_t *q, const uint8_t *n, const uint8_t *p);
    typedef void (*cf_curve25519_mul_base_t)(uint8_t *q, const uint8_t *n);
#endif
#ifdef IMPL_SHA_THUNK
    typedef void (*cf_sha256_init_t)(cf_sha256_context *ctx);
    typedef void (*cf_sha256_update_t)(cf_sha256_context *ctx, const void *data, size_t nbytes);
    typedef void (*cf_sha256_digest_final_t)(cf_sha256_context *ctx, uint8_t hash[CF_SHA256_HASHSZ]);
    typedef void (*cf_sha384_init_t)(cf_sha512_context *ctx);
    typedef void (*cf_sha384_update_t)(cf_sha512_context *ctx, const void *data, size_t nbytes);
    typedef void (*cf_sha384_digest_final_t)(cf_sha512_context *ctx, uint8_t hash[CF_SHA384_HASHSZ]);
#endif
#ifdef IMPL_CHACHA20_THUNK
    typedef void (*cf_chacha20poly1305_encrypt_t)(const uint8_t key[32], const uint8_t nonce[12], const uint8_t *header, size_t nheader,
                                                  const uint8_t *plaintext, size_t nbytes, uint8_t *ciphertext, uint8_t tag[16]);
    typedef int (*cf_chacha20poly1305_decrypt_t)(const uint8_t key[32], const uint8_t nonce[12], const uint8_t *header, size_t nheader,
                                                 const uint8_t *ciphertext, size_t nbytes, const uint8_t tag[16], uint8_t *plaintext);
#endif
#ifdef IMPL_AESGCM_THUNK
    typedef void (*cf_aesgcm_encrypt_t)(uint8_t *c, uint8_t *mac, const uint8_t *m, const size_t mlen, const uint8_t *ad, const size_t adlen,
                                        const uint8_t *npub, const uint8_t *k, size_t klen);
    typedef int (*cf_aesgcm_decrypt_t)(uint8_t *m, const uint8_t *c, const size_t clen, const uint8_t *mac, const uint8_t *ad, const size_t adlen,
                                       const uint8_t *npub, const uint8_t *k, const size_t klen);
#endif

void __cdecl main()
{
#ifdef IMPL_SHA_THUNK
    printf("sizeof(cf_sha256_context)=%d\n", sizeof cf_sha256_context);
    printf("sizeof(cf_sha512_context)=%d\n", sizeof cf_sha512_context);
#endif
#ifdef IMPL_CHACHA20_THUNK
    printf("sizeof(_chacha20_tau)=%d\n", sizeof _chacha20_tau);
#endif
    static thunk_context_t ctx;
#ifdef IMPL_ECC_THUNK
    memcpy(&ctx.m_curve_p, &_curve_p, sizeof _curve_p);
    memcpy(&ctx.m_curve_b, &_curve_b, sizeof _curve_b);
    memcpy(&ctx.m_curve_G, &_curve_G, sizeof _curve_G);
    memcpy(&ctx.m_curve_n, &_curve_n, sizeof _curve_p);
#endif
#ifdef IMPL_SHA_THUNK
    memcpy(&ctx.m_K256, &_K256, sizeof _K256);
    memcpy(&ctx.m_K512, &_K512, sizeof _K512);
#endif
#ifdef IMPL_CHACHA20_THUNK
    memcpy(&ctx.m_chacha20_tau, &_chacha20_tau, sizeof _chacha20_tau);
    memcpy(&ctx.m_chacha20_sigma, &_chacha20_sigma, sizeof _chacha20_sigma);
    memcpy(&ctx.m_negative_1305, &_negative_1305, sizeof _negative_1305);
#endif
#ifdef IMPL_AESGCM_THUNK
    memcpy(&ctx.m_S, &_S, sizeof _S);
    memcpy(&ctx.m_Rcon, &_Rcon, sizeof _Rcon);
    memcpy(&ctx.m_S_inv, &_S_inv, sizeof _S_inv);
#endif

    CoInitialize(0);
    DWORD dwDummy;
    VirtualProtect(beginOfThunk, 1024, PAGE_EXECUTE_READWRITE, &dwDummy);
    ((void **)beginOfThunk)[0] = &ctx;

    size_t thunkSize = THUNK_SIZE;
    while(thunkSize > 4 && ((uint8_t *)beginOfThunk)[thunkSize - 4] == 0)
        thunkSize--;
    void *hThunk = VirtualAlloc(0, 2*THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE);
    printf("hThunk=%p\nTHUNK_SIZE=%d -> %d\n", hThunk, THUNK_SIZE, thunkSize);
    memcpy(hThunk, beginOfThunk, THUNK_SIZE);
    memset(((uint8_t *)hThunk) + thunkSize, 0xCC, 2*THUNK_SIZE - thunkSize);

    // test thunks
#ifdef IMPL_ECC_THUNK
    DECLARE_PFN(ecc_make_key_t, ecc_make_key);
    DECLARE_PFN(ecdh_shared_secret_t, ecdh_shared_secret);
    DECLARE_PFN(cf_curve25519_mul_t, cf_curve25519_mul);
#endif
#ifdef IMPL_SHA_THUNK
    DECLARE_PFN(cf_sha256_init_t, cf_sha256_init);
    DECLARE_PFN(cf_sha256_update_t, cf_sha256_update);
    DECLARE_PFN(cf_sha256_digest_final_t, cf_sha256_digest_final);

    DECLARE_PFN(cf_sha384_init_t, cf_sha384_init);
    DECLARE_PFN(cf_sha384_update_t, cf_sha384_update);
    DECLARE_PFN(cf_sha384_digest_final_t, cf_sha384_digest_final);
#endif
#ifdef IMPL_CHACHA20_THUNK
    DECLARE_PFN(cf_chacha20poly1305_encrypt_t, cf_chacha20poly1305_encrypt);
    DECLARE_PFN(cf_chacha20poly1305_decrypt_t, cf_chacha20poly1305_decrypt);
#endif
#ifdef IMPL_AESGCM_THUNK
    DECLARE_PFN(cf_aesgcm_encrypt_t, cf_aesgcm_encrypt);
    DECLARE_PFN(cf_aesgcm_decrypt_t, cf_aesgcm_decrypt);
#endif

#ifdef IMPL_ECC_THUNK
    uint8_t pubkey[ECC_BYTES+1] = { 0 };
    uint8_t privkey[ECC_BYTES] = { 0 };
    uint8_t secret[ECC_BYTES] = { 0 };
    do {
        _getRandomNumber((uint64_t *)privkey);
    } while (!ecc_make_key(pubkey, privkey));
    pfn_ecc_make_key(pubkey, privkey);
    pfn_ecdh_shared_secret(pubkey, privkey, secret);
    pfn_ecdh_shared_secret(pubkey, privkey, secret);

    pfn_cf_curve25519_mul(secret, privkey, pubkey);
    pfn_cf_curve25519_mul(secret, privkey, pubkey);
#endif
#ifdef IMPL_SHA_THUNK
    cf_sha256_context sha256_ctx = { 0 };
    uint8_t hash256[CF_SHA256_HASHSZ] = { 0 };
    pfn_cf_sha256_init(&sha256_ctx);
    pfn_cf_sha256_update(&sha256_ctx, "123456", 6);
    pfn_cf_sha256_digest_final(&sha256_ctx, hash256);

    cf_sha512_context sha384_ctx = { 0 };
    uint8_t hash384[CF_SHA384_HASHSZ] = { 0 };
    pfn_cf_sha384_init(&sha384_ctx);
    pfn_cf_sha384_update(&sha384_ctx, "123456", 6);
    pfn_cf_sha384_digest_final(&sha384_ctx, hash384);
#endif
    uint8_t key[32] = { 1, 2, 3, 4 };
    uint8_t nonce[12] = { 1, 2, 3, 4 };
    uint8_t tag[16];
    uint8_t aad[] = "header text";
    uint8_t plaintext[] = "this is a test 1234567890";
    uint8_t cyphertext[100] = { 0 };
#ifdef IMPL_CHACHA20_THUNK
    pfn_cf_chacha20poly1305_encrypt(key, nonce, aad, sizeof aad, plaintext, sizeof plaintext, cyphertext, tag);
    pfn_cf_chacha20poly1305_decrypt(key, nonce, aad, sizeof aad, cyphertext, sizeof plaintext, tag, cyphertext);
#endif
#ifdef IMPL_AESGCM_THUNK
    uint8_t *mac = cyphertext + sizeof plaintext;
    pfn_cf_aesgcm_encrypt(cyphertext, mac, plaintext, sizeof plaintext, aad, sizeof aad, nonce, key, sizeof key);
    pfn_cf_aesgcm_decrypt(cyphertext, cyphertext, sizeof plaintext, mac, aad, sizeof aad, nonce, key, sizeof key);
#endif

    // init offsets at beginning of thunk
    int i = 1;
#ifdef IMPL_ECC_THUNK
    ((int *)hThunk)[i++] = ((char *)ecc_make_key - (char *)beginOfThunk);
    ((int *)hThunk)[i++] = ((char *)ecdh_shared_secret - (char *)beginOfThunk);
#ifndef ECC_NO_SIGN
    ((int *)hThunk)[i++] = ((char *)ecdsa_sign - (char *)beginOfThunk);
    ((int *)hThunk)[i++] = ((char *)ecdsa_verify - (char *)beginOfThunk);
#endif
    ((int *)hThunk)[i++] = ((char *)cf_curve25519_mul - (char *)beginOfThunk);
    ((int *)hThunk)[i++] = ((char *)cf_curve25519_mul_base - (char *)beginOfThunk);
#endif
#ifdef IMPL_SHA_THUNK
    ((int *)hThunk)[i++] = ((char *)cf_sha256_init - (char *)beginOfThunk);
    ((int *)hThunk)[i++] = ((char *)cf_sha256_update - (char *)beginOfThunk);
    ((int *)hThunk)[i++] = ((char *)cf_sha256_digest_final - (char *)beginOfThunk);
    ((int *)hThunk)[i++] = ((char *)cf_sha384_init - (char *)beginOfThunk);
    ((int *)hThunk)[i++] = ((char *)cf_sha384_update - (char *)beginOfThunk);
    ((int *)hThunk)[i++] = ((char *)cf_sha384_digest_final - (char *)beginOfThunk);
#endif
#ifdef IMPL_CHACHA20_THUNK
    ((int *)hThunk)[i++] = ((char *)cf_chacha20poly1305_encrypt - (char *)beginOfThunk);
    ((int *)hThunk)[i++] = ((char *)cf_chacha20poly1305_decrypt - (char *)beginOfThunk);
#endif
#ifdef IMPL_AESGCM_THUNK
    ((int *)hThunk)[i++] = ((char *)cf_aesgcm_encrypt - (char *)beginOfThunk);
    ((int *)hThunk)[i++] = ((char *)cf_aesgcm_decrypt - (char *)beginOfThunk);
#endif
    printf("i=%d, needed=0x%02X, allocated=0x%02X\n", i, (i*4 + 15) & -16, ((char *)getContext) - ((char *)beginOfThunk));

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

#ifdef IMPL_ECC_THUNK
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