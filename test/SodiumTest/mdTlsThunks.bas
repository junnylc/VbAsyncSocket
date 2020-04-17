Attribute VB_Name = "mdTlsThunks"
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

'=========================================================================
' API
'=========================================================================

'--- for thunks
Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
'--- for CryptAcquireContext
Private Const PROV_RSA_FULL                 As Long = 1
Private Const CRYPT_VERIFYCONTEXT           As Long = &HF0000000

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Private Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextW" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGenRandom Lib "advapi32" (ByVal hProv As Long, ByVal dwLen As Long, ByVal pbBuffer As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_ECC_GLOB          As String = "////////////////AAAAAAAAAAAAAAAAAQAAAP////9LYNInPjzOO/awU8ywBh1lvIaYdlW967Pnkzqq2DXGWpbCmNhFOaH0oDPrLYF9A3fyQKRj5ea8+EdCLOHy0Rdr9VG/N2hAtsvOXjFrVzPOKxaeD3xK6+eOm38a/uJC409RJWP8wsq584SeF6et+ua8//////////8AAAAA/////5gvikKRRDdxz/vAtaXbtelbwlY58RHxWaSCP5LVXhyrmKoH2AFbgxK+hTEkw30MVXRdvnL+sd6Apwbcm3Txm8HBaZvkhke+78adwQ/MoQwkbyzpLaqEdErcqbBc2oj5dlJRPphtxjGoyCcDsMd/Wb/zC+DGR5Gn1VFjygZnKSkUhQq3JzghGy78bSxNEw04U1RzCmW7Cmp2LsnCgYUscpKh6L+iS2YaqHCLS8KjUWzHGeiS0SQGmdaFNQ70cKBqEBbBpBkIbDceTHdIJ7W8sDSzDBw5SqrYTk/KnFvzby5o7oKPdG9jpXgUeMiECALHjPr/vpDrbFCk96P5vvJ4ccYirijXmC+KQs1l7yORRDdxLztN7M/7wLW824mBpdu16Ti1SPNbwlY5GdAFtvER8VmbTxmvpII/khiBbdrVXhyrQgIDo5iqB9i+b3BFAVuDEoyy5E6+hTEk4rT/1cN9DFVviXvydF2+crGWFjv+sd6ANRLHJacG3JuUJmnPdPGbwdJK8Z7BaZvk4yVPOIZHvu+11YyLxp3BD2WcrHfMoQwkdQIrWW8s6S2D5KZuqoR0StT7Qb3cqbBctVMRg9qI+Xar32buUlE+mBAytC1txjGoPyH7mMgnA7DkDu++x39Zv8KPqD3zC+DGJacKk0eRp9VvggPgUWPKBnBuDgpnKSkU/C/S" & _
                                                "RoUKtycmySZcOCEbLu0qxFr8bSxN37OVnRMNOFPeY6+LVHMKZaiydzy7Cmp25q7tRy7JwoE7NYIUhSxykmQD8Uyh6L+iATBCvEtmGqiRl/jQcItLwjC+VAajUWzHGFLv1hnoktEQqWVVJAaZ1iogcVeFNQ70uNG7MnCgahDI0NK4FsGkGVOrQVEIbDcemeuO30x3SCeoSJvhtbywNGNaycWzDBw5y4pB40qq2E5z42N3T8qcW6O4stbzby5o/LLvXe6Cj3RgLxdDb2OleHKr8KEUeMiE7DlkGggCx4woHmMj+v++kOm9gt7rbFCkFXnGsvej+b4rU3Lj8nhxxpxhJurOPifKB8LAIce4htEe6+DN1n3a6njRbu5/T331um8Xcqpn8AammMiixX1jCq4N+b4EmD8RG0ccEzULcRuEfQQj9XfbKJMkx0B7q8oyvL7JFQq+njxMDRCcxGcdQ7ZCPsu+1MVMKn5l/Jwpf1ns+tY6q2/LXxdYR0qMGURs" ' 1056, 17.4.2020 16:49:09
Private Const STR_ECC_THUNK1        As String = "MB4AAFAhAACAEQAAoBQAANAWAAAwFwAA4BQAAKAXAAAwGAAAcBcAAF3CBADMzMzM6AAAAABYg+g1iwDDzMzMzOgAAAAAWIPoRcPMzMzMzMxVi+yB7KgAAAAzyVaLdRCQiwTOC0TOBHUNQYP5BHLxXovlXcIMAFNXi30MjYVg////V1DoGD0AAI2FYP///1CNReBQ6Og0AACLXQiNReBQU41FoFDopzoAAI1FoFCNRYBQ6Mo0AACNReBQjUWgUOjdPAAAjUWgUI1F4FDosDQAAFZXjUWgUOh1OgAAjUWgUFfomzQAAFaNRaBQ6LE8AACNRaBQVuiHNAAA6DL///9WU1OL+OjYMQAAC8J1C1dT6A0zAACFwHgIV1NT6EE+AADoDP///1ZWVov46LIxAAALwnULV1bo5zIAAIXAeAhXVlboGz4AAOjm/v//VlNWi/joDD4AAAvCdAhXVlbogDEAAFZTjUWgUOjlOQAAjUWgUFPoCzQAAOi2/v//U1NWi/joXDEAAAvCdQtXVuiRMgAAhcB4CFdWVujFPQAA6JD+//9WU1OL+Og2MQAAC8J1C1dT6GsyAACFwHgIV1NT6J89AACLA4PgAYPIAHQd6GD+//9QU1PoCDEAAFOL+OhgOwAAwecfCXsc6wZT6FI7AABTjUWgUOioOwAAjUWgUFbofjMAAOgp/v//i/iNRYBQVlboTD0AAAvCdAhXVlbowDAAAOgL/v//i/iNRYBQVlboLj0AAAvCdAhXVlboojAAAOjt/f//i/iNRYBWUFDoED0AAAvCdAtXjUWAUFDogTAAAI1FgFBTjUWgUOjjOAAAjUWgUFPoCTMAAOi0/f//i/iNReBQU1Do1zwAAAvCdAtXjUXgUFDoSDAAAIsGiQOLRgSLTQyJ" & _
                                                "QwSLRgiJQwiLRgyJQwyLRhCJQxCLRhSJQxSLRhiJQxiLRhyJQxyLAYkGi0EEiUYEi0EIiUYIi0EMiUYMi0EQiUYQi0EUiUYUi0EYiUYYi0EciUYci0XgiQGLReSJQQSLReiJQQiLReyJQQyLRfCJQRCLRfSJQRSLRfhfiUEYi0X8W4lBHF6L5V3CDADMVYvsVot1CDPADx+AAAAAAIsMxgtMxgR1JUCD+ARy8TPSjU4giwELQQR1E0KDwQiD+gRy8LgBAAAAXl3CBAAzwF5dwgQAzMzMzMzMzMzMzMxVi+yB7CQBAABTi10MVlf/dRSLA4lFvItDBIlFwItDCIlFxItDDIlFyItDEIlFzItDFIlF0ItDGIlF1ItDHIlF2ItDIImFfP///4tDJIlFgItDKIlFhItDLIlFiItDMIlFjItDNIlFkItDOIlFlItDPIlFmI2FXP///1CNRZxQjYV8////UI1FvFDoYQcAAIt9EFfoyDgAAIPoAolFFIXAfnYz9ovIg+E/M9IPq86D+SAPQ9Yz8oP5QA9D1sHoBiM0xyNUxwQL8nUFjUYB6wIzwMHgBY2dXP///wPYjU2cA8iNtXz///9T99iJTfxRA/CNfbwD+FZX6BQEAABWV1P/dfzoSQIAAItFFIt9EEiJRRSFwH+Ni10MiweD4AGDyAB1B7gBAAAA6wIzwMHgBY2NXP///wPIjVWcUQPQiU0MjbV8////iVUUUivwjX28K/hWV+i5AwAA6FT7//+JRRCNRZxQjUW8UI1F3FDocDoAAAvCdA3/dRCNRdxQUOjfLQAAVo1F3FCNhRz///9Q6D42AACNhRz///9QjUXcUOheMAAAU41F3FCNhRz///9Q6B02AACNhRz///9QjUXcUOg9M" & _
                                                "AAA6Oj6//9QjUXcUFDo7TIAAI1DIFCNRdxQjYUc////UOjpNQAAjYUc////UI1F3FDoCTAAAFeNRdxQjYUc////UOjINQAAjYUc////UI1F3FDo6C8AAFZX/3UM/3UU6CsBAACNRdxQjYXc/v//UOjrNwAAjYXc/v//UI2FPP///1DouC8AAI2FPP///1CNRZxQjYXc/v//UOhxNQAAjYXc/v//UI1FnFDokS8AAI1F3FCNhTz///9QjYXc/v//UOhKNQAAjYXc/v//UI2FPP///1DoZy8AAI2FPP///1CNhVz///9QjYXc/v//UOgdNQAAjYXc/v//UI2FXP///1DoOi8AAItNCItFnF9eiQGLRaCJQQSLRaSJQQiLRaiJQQyLRayJQRCLRbCJQRSLRbSJQRiLRbiJQRyLhVz///+JQSCLhWD///+JQSSLhWT///+JQSiLhWj///+JQSyLhWz///+JQTCLhXD///+JQTSLhXT///+JQTiLhXj///+JQTxbi+VdwhAAzMzMzMzMzFWL7IPsYFNWV+hS+f//i10Ii/iLdRCNReBTVlDobzgAAAvCdAtXjUXgUFDo4CsAAI1F4FCNRaBQ6JM2AACNRaBQjUXgUOhmLgAAjUXgUFONRaBQ6Cg0AACNRaBQU+hOLgAAjUXgUFaNRaBQ6BA0AACNRaBQVug2LgAA6OH4////dQyLfRRXV4lFCOgBOAAAC8J0Cv91CFdX6HMrAABXjUWgUOgpNgAAjUWgUI1F4FDo/C0AAOin+P//iUUIjUXgU1BQ6Mk3AAALwnQN/3UIjUXgUFDoOCsAAOiD+P//iUUIjUXgVlBQ6KU3AAALwnQN/3UIjUXgUFDoFCsAAOhf+P//U1ZWiUUI6IQ3AAALwn" & _
                                                "QK/3UIVlbo9ioAAFb/dQyNRaBQ6FkzAACNRaBQ/3UM6H0tAADoKPj//4lFCI1F4FBTVuhKNwAAC8J0Cv91CFZW6LwqAABWV41FoFDoITMAAI1FoFBX6EctAADo8vf///91DIvYV1foFjcAAAvCdAhTV1foiioAAItF4IkGi0XkiUYEi0XoiUYIi0XsiUYMi0XwiUYQi0X0iUYUi0X4iUYYi0X8X4lGHF5bi+VdwhAAzMxVi+yB7OAAAABTVlfoj/f//4t1CIv4i10QjUXgVlNQ6Kw2AAALwnQLV41F4FBQ6B0qAACNReBQjYVg////UOjNNAAAjYVg////UI1F4FDonSwAAI1F4FBWjYVg////UOhcMgAAjYVg////UFbofywAAI1F4FBTjYVg////UOg+MgAAjYVg////UFPoYSwAAOgM9////3UMi30UiUUIjUXgV1DoqSkAAAvCdRD/dQiNReBQ6NkqAACFwHgN/3UIjUXgUFDoCDYAAOjT9v///3UMiUUIV1fo9jUAAAvCdAr/dQhXV+hoKQAA6LP2//9WiUUIjUWAU1Do1TUAAAvCdA3/dQiNRYBQUOhEKQAAjUWAUP91DI1FoFDopDEAAI1FoFD/dQzoyCsAAOhz9v//U4lFCI1FgFZQ6BUpAAALwnUQ/3UIjUWAUOhFKgAAhcB4Df91CI1FgFBQ6HQ1AABXjUWgUOiqMwAAjUWgUFPogCsAAOgr9v//iUUIjUWAUFNT6E01AAALwnQK/3UIU1PovygAAOgK9v//U4lFCI1FwFZQ6Cw1AAALwnQN/3UIjUXAUFDomygAAI1FwFBXjYUg////UOj6MAAAjYUg////UFfoHSsAAOjI9f//i10MU1dXiUUI6Oo0AAALwnQK/3U" & _
                                                "IV1foXCgAAI1F4FCNhSD///9Q6AwzAACNhSD///9QjUXAUOjcKgAA6If1//+L+I1FgFCNRcBQUOinNAAAC8J0C1eNRcBQUOgYKAAA6GP1//+L+I1FwFZQjUWAUOiDNAAAC8J0C1eNRYBQUOj0JwAAjUXgUI1FgFCNhSD///9Q6FAwAACNhSD///9QjUWAUOhwKgAA6Bv1//+L+I1FgFNQU+g+NAAAC8J0CFdTU+iyJwAAi0XAiQaLRcSJRgSLRciJRgiLRcyJRgyLRdCJRhCLRdSJRhSLRdiJRhiLRdxfiUYcXluL5V3CEADMzMzMzMzMzMzMVYvsg+wgi00UD1fAU4tdEFaLdQhXi30MZg8TReiLBokDi0YEiUMEi0YIiUMIi0YMiUMMi0YQiUMQi0YUiUMUi0YYiUMYi0YciUMciweJAYtHBIlBBItHCIlBCItHDIlBDItHEIlBEItHFIlBFItHGIlBGItHHIlBHItNGGYPE0XwZg8TRfjHReABAAAAx0XkAAAAAIXJdC+LAYlF4ItBBIlF5ItBCIlF6ItBDIlF7ItBEIlF8ItBFIlF9ItBGIlF+ItBHIlF/I1F4FBXVug+AQAAjUXgUFdW6AP0//+NReBQ/3UUU+gmAQAAX15bi+VdwhQAzMzMzMzMzMzMzMzMzFOLRCQMi0wkEPfhi9iLRCQI92QkFAPYi0QkCPfhA9NbwhAAzMzMzMzMzMzMzMzMzID5QHMVgPkgcwYPpcLT4MOL0DPAgOEf0+LDM8Az0sPMgPlAcxWA+SBzBg+t0NPqw4vCM9KA4R/T6MMzwDPSw8xVi+yLRRBTVot1CI1IeFeLfQyNVng78XcEO9BzC41PeDvxdzA713IsK/i7EAAAACvwixQ4AxCLTDgE" & _
                                                "E0gEjUAIiVQw+IlMMPyD6wF15F9eW13CDACL141IEIveK9Ar2Cv+uAQAAACNdiCNSSAPEEHQDxBMN+BmD9TIDxFO4A8QTArgDxBB4GYP1MgPEUwL4IPoAXXSX15bXcIMAMzMzMzMVYvsg+xgjUWgVv91EFDoDTAAAI1FoFCNReBQ6OAnAACLdQiNReBQVo1FoFDony0AAI1FoFBW6MUnAAD/dRCNReBQjUWgUOiFLQAAjUWgUI1F4FDoqCcAAIt1DI1F4FBWjUWgUOhnLQAAjUWgUFbojScAAF6L5V3CDADMzMzMzMxVi+yD7CBTVot1CDPJV4lN7IEEzgAAAQCLBM6DVM4EAItczgQPrNgQwfsQiUXog/kPdRXHRfwBAAAAi9DHRfAAAAAAiV346yIPV8BmDxNF9ItF+IlF8ItF9GYPE0Xgi1XgiUX8i0XkiUX4g/kPjXkBagAbwPfYD6/HK1X8aiWNNMaLRfgbRfBQUuji/f//i03oA8ET04PoAYPaAAEGi0XsEVYEi3UID6TLEMHhECkMxovPiU3sGVzGBIP5EA+CT////19eW4vlXcIEAMzMzMzMVYvsg+wQU4tdGDPJiU38hdsPhJEAAACLVRBWV4t9DMdF8AEAAACLB4vyK/CJRfQ73g9C84XJdTuLfQgD+Il9+It9DIX2dCwPtkUUi86LffiL2WnAAQEBAcHpAotVEPOri8uLXRiD4QPzqotF9ItN/It9DIXAdQk78g9ETfCJTfwDxjvCdRT/dQj/dSD/VRyLVRDHBwAAAADrAgE3i038K96JXRh1gF9eW4vlXcIcAMxVi+xXi30gi8eD6AB0YoPoAQ+EtQAAAFNWg+gBdHKLXSiNRRSLdSRTVlZqAVD/dRD/dQz/dQjox" & _
                                                "QAAAItNGFNWOE0cdDCNR/6LfRBQUVf/dQz/dQjo9/7//1NWVmoBjUUcUFf/dQz/dQjokgAAAF5bX13CJACNR/+LfRBQUVf/dQz/dQjox/7//15bX13CJAD/dSiLRSSLXRCLfQyLdQhQUGoBjUUUUFNXVuhRAAAA/3Uoi0UkUFBqAY1FHFBTV1boOwAAAF5bX13CJAD/dSiKRRwwRRSLRSRQUGoBjUUUUP91EP91DP91COgSAAAAX13CJADMzMzMzMzMzMzMzMzMVYvsi00Mg+wIi1UkgzkAU4tdFFaLdRBXi30YdG2F/3Rpi8aL3ysBO8cPQtiLAQNFCIldGItdFIlF/ItFGIXAdCOLVfyL8IvLigGNUgGIQv+NSQGD7gF18It1EItNDItVJItFGAEBA9gr+DkxdRz/dQhShf91Bf9VIOsD/1Uci0UMi1UkxwAAAAAAO/5yG2aQU1I7/nUF/1Ug6wP/VRyLVSQr/gPeO/5z54X/dESLTQyLCYvGi3UIK8E7x4vXD0LQA/GJVSSLw4XSdBUPH0QAAIoIjXYBiE7/jUABg+oBdfCLTQyLRSQD2It1EAEBK/h1v19eW4vlXcIgAMzMzMzMzFWL7IHsKAQAAFNWV2pwjYXo/P//x4XY/P//QdsAAGoAUMeF3Pz//wAAAADHheD8//8BAAAAx4Xk/P//AAAAAOhsEgAAi3UMjYVg////ah9WUOgqEgAAikYfg8QYgKVg////+CQ/DECIhX////+Nhdj7////dRBQ6BQgAAAPV8CNtVj+//9mDxOFWP7//429YP7//7keAAAAZg8TRYDzpbkeAAAAZg8TheD+//+NdYDHhVj+//8BAAAAjX2Ix4Vc/v//AAAAAPOluR4AAADHRYABAAAAjb" & _
                                                "Xg/v//x0WEAAAAAI296P7//7v+AAAA86W5IAAAAI212Pv//4292P3///Oli8OKy8H4A4DhB4qEBWD////S6CQBD7bAmYvwjYXY/f//VlCNRYBQ6BMYAABWjYVY/v//UI2F4P7//1Do/xcAAI2F4P7//1CNRYBQjYVY/f//UOgI+v//jYXg/v//UI1FgFBQ6IceAACNhVj+//9QjYXY/f//UI2F4P7//1Do3fn//42FWP7//1CNhdj9//9QUOhZHgAAjYVY/f//UFCNhVj+//9Q6EUSAACNRYBQUI2FWPz//1DoNBIAAI1FgFCNheD+//9QjUWAUOggEgAAjYVY/f//UI2F2P3//1CNheD+//9Q6AYSAACNheD+//9QjUWAUI2FWP3//1DoX/n//42F4P7//1CNRYBQUOjeHQAAjUWAUFCNhdj9//9Q6M0RAACNhVj8//9QjYVY/v//UI2F4P7//1Dosx0AAI2F2Pz//1CNheD+//9QjUWAUOicEQAAjYVY/v//UI1FgFBQ6Pv4//+NRYBQjYXg/v//UFDoehEAAI2FWPz//1CNhVj+//9QjUWAUOhjEQAAjYXY+///UI2F2P3//1CNhVj+//9Q6EkRAACNhVj9//9QUI2F2P3//1DoNREAAFaNhdj9//9QjUWAUOh0FgAAVo2FWP7//1CNheD+//9Q6GAWAACD6wEPiRj+//+NheD+//9QUOhqDQAAjYXg/v//UI1FgFBQ6OkQAACNRYBQ/3UI6A0UAABfXluL5V3CDADMzMzMVYvsg+wgjUXgxkXgCVD/dQwPV8DHRfkAAAAA/3UIDxFF4WbHRf0AAGYP1kXxxkX/AOiq/P//i+VdwggAzMzMzFWL7IPsGFNWu7BrJACB6wBAJAB" & _
                                                "XiV386ETr////dQi5QAAAAI00A4tFCFaLdQiNWGSLQGD34QMDi/qLVQiL2IPXAIPACIPgP4PGZCvIg8IgUWoAagBogAAAAGpAD6TfA1ZSweMDiX346DP6//+Lx4hd78HoGIvLiEXoi8fB6BCIRemLx8HoCIhF6opF+IhF64vHD6zBGMHoGIhN7IvHi8sPrMEQwegQi8OITe0PrPgIwe8IiEXu6KXq//8DRfyLXQhTUFBqCI1F6FBqQFaNQyBQ6Mr6//+LE4vCi3UMwegYiAaLwsHoEIhGAYvCwegIiEYCiFYDi0sEi8HB6BiIRgSLwcHoEIhGBYvBwegIiEYGiE4Hi0sIi8HB6BiIRgiLwcHoEIhGCYvBwegIiEYKiE4Li0sMi8HB6BiIRgyLwcHoEIhGDYvBwegIiEYOiE4Pi0sQi8HB6BiIRhCLwcHoEIhGEYvBwegIiEYSiE4Ti0sUi8HB6BiIRhSLwcHoEIhGFYvBwegIiEYWiE4Xi0sYi8HB6BiIRhiLwcHoEIhGGYvBwegIiEYaiE4bi0sci8HB6BiIRhyLwcHoEIhGHQ9XwIvBwegIiEYeiE4fDxEDXw8RQxBeDxFDIA8RQzAPEUNADxFDUGYP1kNgW4vlXcIIAMxVi+yLRQgPV8APEQAPEUAQDxFAIA8RQDAPEUBADxFAUGYP1kBgxwBn5glqx0AEha5nu8dACHLzbjzHQAw69U+lx0AQf1IOUcdAFIxoBZvHQBir2YMfx0AcGc3gW13CBABVi+zoCOn//7mwayQAgekAQCQAA8GLTQhRUFD/dRCNQWT/dQxqQFCNQSBQ6CD5//9dwgwAzMzMzMzMzMzMzMzMVYvsg+xAjUXAUP91COjuAAAAajCNRcBQ/3UM6HAMAACD" & _
                                                "xAyL5V3CCADMzMzMzMzMVYvsi1UIuTIAAABXM8CL+vOrxwLYngXBx0IEXZ27y8dCCAfVfDbHQgwqKZpix0IQF91wMMdCFFoBWZHHQhg5WQ73x0Ic2OwvFcdCIDELwP/HQiRnJjNnx0IoERVYaMdCLIdKtI7HQjCnj/lkx0I0DS4M28dCOKRP+r7HQjwdSLVHX13CBADMzMzMzMzMzMzMVYvs6Ajo//+5gG0kAIHpAEAkAAPBi00IUVBQ/3UQjYHEAAAA/3UMaIAAAABQjUFAUOga+P//XcIMAMzMzMzMzFWL7IPsGFNWu4BtJACB6wBAJABXiV386LTn////dQi5gAAAAI00A4tFCFaLdQiNmMQAAACLgMAAAAD34QMDi/qLVQiL2IPXAIPAEIPgf4HGxAAAACvIg8JAUWoAagBogAAAAGiAAAAAD6TfA1ZSweMDiX346Jf2///HRegAAAAAx0XsAAAAAOhE5///A0X8i00IUVBQagiNRehQaIAAAABWjXFAVuhm9///i8eIXe/B6BiLy4hF6IvHwegQiEXpi8fB6AiIReqKRfiIReuLxw+swRjB6BiITeyLx4vLD6zBEMHoEIvDiE3tD6z4CMHvCIhF7ujY5v//A0X8i10IU1BQagiNRehQaIAAAACNg8QAAABQVuj39v//iwuLWwSLw4t9DMHoGIlN/IgHi8PB6BCIRwGLw8HoCIhHAovDD6zBGIhfA8HoGIhPBIvDi038D6zBEMHoEIhPBYtN/IvBD6zYCIhHBotFCIhPB8HrCItYCIvLi1AMi8LB6BiIRwiLwsHoEIhHCYvCwegIiEcKi8IPrMEYiFcLwegYiE8Mi8KLyw+swRDB6BCITw2Lww+s0AiIRw6LRQiIXw/B6giLW" & _
                                                "BCLUBSLwsHoGIhHEIvCwegQiEcRi8LB6AiIRxKIVxOLwovLD6zBGMHoGIhPFIvCi8sPrMEQwegQiE8Vi8MPrNAIiEcWi0UIiF8XweoIi1gYi8uLUByLwsHoGIhHGIvCwegQiEcZi8LB6AiIRxqLwg+swRiIVxvB6BiITxyLwovLD6zBEMHoEIhPHYvDD6zQCIhHHotFCIhfH8HqCItYIIvLi1Aki8LB6BiIRyCLwsHoEIhHIYvCwegIiEcii8IPrMEYiFcjwegYiE8ki8KLyw+swRDB6BCITyWLww+s0AiIRyaLRQiIXyfB6giLWCiLUCyLwsHoGIhHKIvCwegQiEcpi8KNdzjB6AiLy4hHKovCD6zBGIhXK4hPLIvLwegYi8IPrMEQwegQiE8ti8MPrNAIiEcui0UIiF8vweoIi1gwi8uLUDSLwsHoGIhHMIvCwegQiEcxi8LB6AiIRzKLwg+swRiIVzOITzSLy8HoGIvCD6zBEMHoEIhPNYvDD6zQCIhHNohfN4t9CMHqCItXPIvCi184i8vB6BiIBovCwegQiEYBi8LB6AiIRgKLwg+swRiIVgOITgSLy8HoGIvCD6zBEMHoEIhOBYvDD6zQCLkyAAAAiEYGweoIM8CIXgfzq19eW4vlXcIIAMzMzMzMzMzMVYvsU4tdDFZXi30ID7ZDGJmLyIvyD6TOCA+2QxnB4QiZC8gL8g+kzggPtkMaweEImQvIC/IPpM4ID7ZDG8HhCJkLyAvyD7ZDHA+kzgiZweEIC/ILyA+2Qx0PpM4ImcHhCAvyC8gPtkMeD6TOCJnB4QgL8gvID7ZDHw+kzgiZweEIC/ILyIl3BIkPD7ZDEJmLyIvyD7ZDEQ+kzgiZweEIC/ILyA+2QxIPpM4Imc"
Private Const STR_ECC_THUNK2        As String = "HhCAvyC8gPtkMTD6TOCJnB4QgL8gvID7ZDFA+kzgiZweEIC8gL8g+kzggPtkMVweEImQvIC/IPpM4ID7ZDFsHhCJkLyAvyD6TOCA+2QxfB4QiZC8gL8olPCIl3DA+2QwiZi8iL8g+kzggPtkMJweEImQvIC/IPtkMKD6TOCJnB4QgL8gvID7ZDCw+kzgiZweEIC/ILyA+2QwwPpM4ImcHhCAvyC8gPtkMND6TOCJnB4QgL8gvID7ZDDg+kzgiZweEIC/ILyA+2Qw8PpM4ImcHhCAvyC8iJdxSJTxAPtgOZi8iL8g+2QwEPpM4ImcHhCAvyC8gPtkMCD6TOCMHhCJkLyAvyD7ZDAw+kzgiZweEIC/ILyA+2QwQPpM4ImcHhCAvyC8gPtkMFD6TOCJnB4QgL8gvID7ZDBg+kzgiZweEIC/ILyA+2QwcPpM4ImcHhCAvIC/KJdxyJTxhfXltdwggAzMzMVYvsg+xgjUXg/3UMUOje/f//M8CLTMXgC0zF5HUOQIP4BHLwM8CL5V3CCACNReBQ6Mvh//+D6IBQ6LIVAACD+AF0E+i44f//g+iAUI1F4FBQ6NogAABqAI1F4FDon+H//4PAQFCNRaBQ6OLk//+NRaBQ6Ink//+FwHWpikXAi00IJAEEAogBjUWgUI1BAVDoDAAAALgBAAAAi+VdwggAzFWL7FaLdQixKFeLfQwPtkcHiEYYD7ZHBohGGYsHi1cE6Mvt//+IRhqxIIsHi1cE6Lzt//+IRhuLD4tHBA+swRiIThyLD8HoGItHBA+swRCITh2LD8HoEItHBA+swQiITh6xKMHoCA+2B4hGHw+2Rw+IRhAPtkcOiEYRi0cIi1cM6Gvt//+IRhKxIItHCItXDOhb7f//iEYTi08" & _
                                                "Ii0cMD6zBGIhOFItPCMHoGItHDA+swRCIThWLTwjB6BCLRwwPrMEIiE4WsSjB6AgPtkcIiEYXD7ZHF4hGCA+2RxaIRgmLRxCLVxToBu3//4hGCrEgi0cQi1cU6Pbs//+IRguLTxCLRxQPrMEYiE4Mi08QwegYi0cUD6zBEIhODYtPEMHoEItHFA+swQiITg6xKMHoCA+2RxCIRg8PtkcfiAYPtkceiEYBi0cYi1cc6KLs//+IRgKxIItHGItXHOiS7P//iEYDi08Yi0ccD6zBGMHoGIhOBItPGItHHA+swRDB6BCITgWLTxiLRxwPrMEIwegIiE4GD7ZHGF+IRgdeXcIIAMzMVYvsg+xgi0UMD1fAU1ZXi30IQFBXx0XgAwAAAMdF5AAAAAAPEUXoZg/WRfjof/v//1eNRaBQ6PUcAACNRaBQjXcgVujIFAAA6HPf//+L2I1F4FBWVuiWHgAAC8J0CFNWVugKEgAAV1aNRaBQ6G8aAACNRaBQVuiVFAAA6EDf//+L+Og53///g8AgUFZW6N4RAAALwnULV1boExMAAIXAeAhXVlboRx4AAFboQQMAAItNDDP/igGLDiQBD7bAg+EBmTvIdQQ7+nQNVujx3v//UFboGh4AAF9eW4vlXcIIAMxVi+yB7KAAAACNhWD/////dQhQ6Aj/////dQyNRaBQ6Kz6//9qAI1FoFCNhWD///9QjUXAUOj24f//jUXAUP91EOg6/f//M8kPH4QAAAAAAItEzcALRM3EdR5Bg/kEcvCLTMXgC0zF5HUOQIP4BHLwM8CL5V3CDAC4AQAAAIvlXcIMAMzMzMzMzMzMzMzMzMxVi+yB7IQBAABTVot1DLkgAAAAV429fP7//7v9AAAA86WJXfiNhXz+" & _
                                                "//9QUFDoXgMAAIP7Ag+EvQEAAIP7BA+EtAEAAItFDI21/P7//w9XwI29BP///4PAEGYPE4X8/v//uTwAAACJRfTzpTPbDx8AjbUE////x0X8BAAAAIv4A/P/tB2A/v///7QdfP7///939P938Oj26f///7QdgP7//wFG+P+0HXz+//8RVvz/d/z/d/jo1+n///+0HYD+//8BBv+0HXz+//8RVgT/dwT/N+i66f///7QdgP7//wFGCP+0HXz+//8RVgz/dwz/dwjom+n//wFGEI1/IBFWFI12IINt/AEPhXb///+LRfSDwwiB+4AAAAAPglP///8z9pBqAGom/3T1gP+09Xz////oXOn//wGE9fz+//9qABGU9QD///9qJv909Yj/dPWE6D3p//8BhPUE////agARlPUI////aib/dPWQ/3T1jOge6f//AYT1DP///2oAEZT1EP///2om/3T1mP909ZTo/+j//wGE9RT///9qABGU9Rj///9qJv909aD/dPWc6ODo//8BhPUc////EZT1IP///4PGBYP+Dw+CVv///42FfP7//7kgAAAAjbX8/v//jb18/v//86VQ6Dfq//+NhXz+//9Q6Cvq//+LXfiD6wGJXfgPiSD+//+LfQiNtXz+//+5IAAAAPOlX15bi+VdwggAzMzMVYvsi0UIi9BWi3UQhfZ0FVeLfQwr+IoMF41SAYhK/4PuAXXyX15dw8zMzMzMzMzMVYvsi00Qhcl0Hw+2RQxWi/FpwAEBAQFXi30IwekC86uLzoPhA/OqX16LRQhdw8zMVYvsgeyAAAAAU1YPV8DHRcABAAAAV41FwMdFxAAAAAC7AQAAAGYP1kXYUA8RRciJXeDHReQAAAAADxFF6GYP1kX46Inb/" & _
                                                "/9QjUXAUOgvDgAAjUXAUOgmGAAAi30IjXD/O/N2aY1F4FCNRYBQ6M8YAACNRYBQjUXgUOiiEAAAM9IzyYvGg+A/D6vCg/ggD0PKM9GD+ECLxg9DysHoBiNUxcAjTMXEC9F0G1eNReBQjUWAUOg5FgAAjUWAUI1F4FDoXBAAAE6D/gF3motd4ItF5IlHBItF6IlHCItF7IlHDItF8IlHEItF9IlHFItF+IlHGItF/IkfiUccX15bi+VdwgQAzMzMzMzMzMzMzMzMzMxVi+yB7BQBAACLRQwPV8BTVle5PAAAAGYPE4Xs/v//jbXs/v//x0X8EAAAAI299P7///Oli00QjZ30/v//g8EQi9MrwolN+IlFDGYPH0QAAIv5x0UQBAAAAIvzDx9EAAD/dBgE/zQY/3f0/3fw6I7m//8BRviLRQwRVvz/dBgE/zQY/3f8/3f46HPm////dwQBBotFDP83EVYE/3QYBP80GOha5v//AUYIi0UMEVYM/3QYBP80GP93DP93COg/5v//AUYQjX8gi0UMEVYUjXYgg20QAXWKi034g8MIg238AQ+Fav///zP2Dx+EAAAAAABqAGom/7T1cP////+09Wz////o+eX//wGE9ez+//9qABGU9fD+//9qJv+09Xj/////tPV0////6NTl//8BhPX0/v//agARlPX4/v//aib/dPWA/7T1fP///+iy5f//AYT1/P7//2oAEZT1AP///2om/3T1iP909YTok+X//wGE9QT///9qABGU9Qj///9qJv909ZD/dPWM6HTl//8BhPUM////EZT1EP///4PGBYP+Dw+CSv///4tVCI217P7//7kgAAAAi/rzpTPJiU34Dx8AgQTKAAABAIsEyoNUygQAi1zKBA" & _
                                                "+s2BDB+xCJRfSD+Q91FcdFDAEAAACL0MdF/AAAAACJXRDrIg9XwGYPE0Xsi0XwiUX8i0XsZg8TReSLVeSJRQyLReiJRRCD+Q+NeQGLTQgbwPfYD6/HK1UMagBqJY00wYtFEBtF/FBS6MDk//+LTfQDwRPTg+gBg9oAAQaLRfgRVgSLVQgPpMsQiX34weEQKQzCi88ZXMIEg/kQD4JM////UugW5v//X15bi+VdwgwAzMzMzMzMzMzMzMzMzFWL7IPsEFNWi3UMV4t9GGoAVmoA/3UU6FTk//9qAFZqAFeJRfCL2uhE5P//agD/dRCJRfSL8moAV+gy5P//agD/dRCJRfxqAP91FIlV+Ogd5P//i/iLRfQD+4PSAAP4E9Y71ncOcgQ7+HMIg0X8AINV+AGLRQgzyQtN8IkIM8kDVfyJeAQTTfhfXolQCIlIDFuL5V3CFADMzMzMzMzMzMxVi+yB7AgBAACNhXj///+5IAAAAFNWi3UMV429eP////OlUOg45f//jYV4////UOgs5f//jYV4////UOgg5f//jb34/v//uwIAAAAPH0QAAIuNeP///4uFfP///4Hp7f8AAImN+P7//4PYAImF/P7//7gIAAAAZmYPH4QAAAAAAIt0B/iLTAf8i5QFeP///4l1+A+szhCLjAV8////g+YBx0QH/AAAAAAr1oPZAIHq//8AAImUBfj+//+D2QCJjAX8/v//D7dN+IlMB/iDwAiD+HhyrIuNaP///4uFbP///4tV8A+swRAPt4Vo////g+EBiYVo////K9HHhWz///8AAAAAi030uAEAAACD2QCB6v9/AACJlXD///+D2QCJjXT///8PrMoQg+IBwfkQK8JQjYX4/v//UI2FeP///1Do3QA" & _
                                                "AAIPrAQ+FBP///4t1CDPSioTVeP///4uM1Xj///+IBFaLhNV8////D6zBCIhMVgFCwfgIg/oQctdfXluL5V3CCADMzMzMzMzMzMzMzMzMVYvsVleLfQgPtgeZi8iL8g+2RwEPpM4ImcHhCAvyC8gPtkcCD6TOCJnB4QgL8gvID7ZHAw+kzgiZweEIC/ILyA+2RwQPpM4ImcHhCAvyC8gPtkcFD6TOCJnB4QgL8gvID7ZHBg+kzgiZweEIC/ILyA+2RwcPpM4ImcHhCAvBC9ZfXl3CBADMzMzMzMzMzMzMVYvsg+wIi0UQSPfQmVOLXQiJRfiLRQyJVfzzD35d+I1LeFYz9mYPbNuNUHg7wXdLO9NyRyvYx0UQEAAAAFdmkIs8GI1ACIt0GPyLSPiLUPwzzyNN+DPWI1X8M/kz8ol8GPiJdBj8MUj4MVD8g20QAXXOX15bi+VdwgwAi9ONSBAr0A8QDPONSSAPEFHQZg/v0WYP29MPKMJmD+/BDxEE84PGBA8QQdBmD+/QDxFR0A8QTArgDxBR4GYP79FmD9vTDyjCZg/vwQ8RRArgDxBB4GYP78IPEUHgg/4QcqVeW4vlXcIMAMzMzMzMzMzMzMzMVYvsg+xsi0UIjVWUU1a7oAAAADP2i0gEiU34i0gIiU30i0gMiU3oi0gQiU38i0gUiU3wi0gYiU3si00Mg8ECiXXcV4s4K9OLQByJfeCJReSJTdiJXQyJVdQPH4AAAAAAg/4QcykPtnH+D7ZB/8HmCAvwD7YBweYIC/APtkEBweYIC/CDwQSJNBqJTdjrVI1eAYPmD41D/YPgD419lI08t4tMhZSLw4PgD4vxwcYPi1SFlIvBwcANM/DB6Qoz8YvCi8rByAfBwQ4zyMHqA41D" & _
                                                "+DPKi10Mg+APA/EDdIWUAzeJN+iZ0///i338i9fByguLz8HBBzPRi8/ByQb31yN97DPRiwwYg8MEi0XwA8ojRfwDzot14DP4i9aJXQzByg2LxsHACgP5A33kM9CLxsHIAjPQi0X4i8gjxjPOI030M8iLReyJReQD0YtF8ItN+IlF7ItF/IlF8ItF6APHiXX4i3XcA/qLVdRGiUX8i0X0iU30i03YiUXoiX3giXXcgfugAQAAD4LX/v//i0UIi034i1X8AUgEi030AUgIAVAQATiLTeiLVfABSAwBUBSLVeyLTeQBUBgBSBz/QGBfXluL5V3CCADMzMzMzMzMzMzMzMxVi+yB7OAAAABTVot1CLugAQAAV4lduIsGiUXsi0YEiUXwi0YMi34IiUXgi0YQiUXUi0YUiUXQi0YYiUW0i0YciUWwi0YgiUXoi0YkiUX0i0YoiUXMi0YsiUXIi0YwiUXEi0Y0iUXAi0Y4iUWsi0Y8i3UMiX3Yjb0g////iUWoM8Ar+4lF3Il9oA8fgAAAAACD+BBzH1boFfz//4vIg8YIi8KJTQyJReSJDB+JRB8E6RMBAACNUAHHRQwAAAAAjUL9g+APi4zFIP///4uExST///+JRfiLwoPgD4lN/I2NIP///4uUxSD///+L+oucxST///+LRdyD4A+JVbzB5xiNBMGLy4lFpIvCD6zICAlFDItFvMHpCAv5i8sPrMgBiX3ki/rR6TPSC9DB5x8xVQwL+YtFvItN5A+s2AczzzFFDItF/MHrBzPLM9uJTeSLTfiL0Q+kwQPB6h3B4AML2YtN+AvQi0X8i/gPrMgTiVW8M9IL0MHpE4tFvDPCwecNi1X8C/mLTfgz3w+sygYzwsHpBotVDDPZi03kA9CLR" & _
                                                "dwTy4PA+YPgDwOUxSD///8TjMUk////i0WkAxCJVQwTSASJEIlN5IlIBOjk0P//i1X0M/+LTeiL2g+kyhfB6wkL+sHhF4tV9AvZi03oiV38i9kPrNESiX34M/8L+cHqEjF9/DP/i03oweMOC9qLVfQxXfiL2Q+s0Q7B4xIL+cHqDjF9/Avai034i1W4M8uLXfyLfej31wMcEBNMEAQjfcSLVfSLRcj30iNF9CNVwDPQiU34i03MI03oi0X4M/mLTfAD3xPCA10ME0XkA12siV38E0WoM9uJRfiLReyL0A+syBzB4gTB6RwL2ItF7AvRi03wi/kPpMEeiVUMM9LB7wIL0cHgHgv4M98xVQwz0otN8Iv5i0XsD6TBGcHvBwvRweAZMVUMC/iLTdgz34tV4Iv5M33sI33UI03sM1XwM/kjVdCLReAjRfCLTcQz0ItFDAPfi334E8KJTayLTcCLVfwDVbSJTagTfbCLTcwDXfyJTcSLTciJTcCLTeiJTcyLTfSJffSLfdSJfbSLfdCJfbCLfdiJfdSLfeCJfdCLfeyJTciLyBNN+ItF3Ild7ECLXbiJfdiDwwiLffCJfeCLfaCJVeiJTfCJRdyJXbiB+yAEAAAPghv9//+LdQiLReyLfdgBBotF4BFOBIvKAX4Ii320EUYMi0XUAUYQi0XQEUYUAX4Yi0WwEUYcAU4gi0X0EUYki0XMAUYoi0XIEUYsi0XEAUYwi0XAEUY0i02sAU44i02oEU48/4bAAAAAX15bi+VdwggAzMzMzMzMzMzMzMzMzMxVi+yLRRBTVot1CI1IeFeLfQyNVng78XcEO9BzC41PeDvxdzA713IsK/i7EAAAACvwixQ4KxCLTDgEG0gEjUAIiVQw+IlMMPyD6w" & _
                                                "F15F9eW13CDACL141IEIveK9Ar2Cv+uAQAAACNdiCNSSAPEEHQDxBMN+BmD/vIDxFO4A8QTArgDxBB4GYP+8gPEUwL4IPoAXXSX15bXcIMAMzMzMzMVYvsi00MU4tdCFaDwxDHRQwEAAAAV4PBAw8fgAAAAAAPtkH+jVsgmY1JCIvwi/oPtkH1D6T3CJnB5ggD8Ilz0BP6iXvUD7ZB95mL8Iv6D7ZB+JkPpMIIweAIA/CJc9gT+ol73A+2QfqZi/CL+g+2QfkPpPcImcHmCAPwiXPgE/qJe+QPtkH8mYvwi/oPtkH7D6T3CJnB5ggD8Ilz6BP6g20MAYl77A+FdP///4tNCF9eW4FheP9/AADHQXwAAAAAXcIIAMzMzMzMzMzMzMzMzFWL7IPsCFOLXQwPV8BWV4t9EIsTi/KLQwSLyGYPE0X4AzcTTwQ78nUGO8h1BOsYO8h3D3IEO/JzCbgBAAAAM9LrC2YPE0X4i0X4i1X8i30IiU8Ei00QiTeLcQgDcwiLSQwTSwwD8BPKO3MIdQU7Swx0IDtLDHcQcgU7cwhzCbgBAAAAM9LrC2YPE0X4i1X8i0X4iU8Mi00QiXcIi3EQA3MQi0kUE0sUA/ATyjtzEHUFO0sUdCA7SxR3EHIFO3MQcwm4AQAAADPS6wtmDxNF+ItV/ItF+IlPFIl3EItLGItbHIlNDItNEItxGAN1DItJHBPLA/ATyjt1DHUEO8t0LDvLdx1yBTt1DHMWiXcYuAEAAACJTxwz0l9eW4vlXcIMAGYPE0X4i1X8i0X4iXcYiU8cX15bi+VdwgwAzMzMzMzMVYvsi00MugMAAABTi10IVivZjUEYV4ldCA8fgAAAAACLNAOLXAMEi3gEiwg733cuciI78XcoO99" & _
                                                "yGncEO/FyFItdCIPoCIPqAXnVX14zwFtdwggAX16DyP9bXcIIAF9euAEAAABbXcIIAMzMzMzMzFWL7IPsEFOLXRC5QAAAAFaLdQgry1eLfQxmD27DiU0QiweLVwSJRfiJVfzzD35N+GYP88hmD9YO6PPX//+LTRCJRfCLRwiJVfSLVwyJRfiJVfzzD35N+GYPbsNmD/PI8w9+RfBmD+vIZg/WTgjovtf//4tNEIlF8ItHEIlV9ItXFIlF+IlV/PMPfk34Zg9uw2YP88jzD35F8GYP68hmD9ZOEOiJ1///i00QiUXwi0cYiVX0i1cciUX4iVX88w9+TfhmD27DZg/zyPMPfkXwZg/ryGYP1k4Y6FTX//9fXluL5V3CDADMzMzMzMzMzMzMzFWL7IPsKFOLXQwPV8BWi3UIV4sDagGJBotDBIlGBItDCIlGCItDDIlGDItDEIlGEItDFIlGFItDGIlGGItDHIlGHItDLIlF5ItDMIlF6ItDNIlF7ItDOIlF8ItDPIlF9I1F2FBQZg8TRdjHReAAAAAA6Jr+//+L+I1F2FBWVujd/P//i0swA/iLUzwzwAtDNIlF6I1F2GoBUIlN5ItLOFDHReAAAAAAiU3siVXwx0X0AAAAAOhX/v//A/iNRdhQVlbomvz//wP4x0XkAAAAAItDIA9XwIlF2ItDJIlF3ItDKIlF4ItDOIlF8ItDPGYPE0XoiUX0jUXYUFZW6GD8//8D+ItLJDPAi1M0C0MoiUXci0MwiUX4M8ALQyyJReCLQziJReiLQzyJRewzwAtDIIlF9I1F2FCJTdiLylZWiU3kiVXw6Bj8//+LUzQD+ItLLDPAC0MwD1fAiUXcM8ALQyCJRfCNRdhQVol9CIt7KFaJTdiJVeDH" & _
                                                "ReQAAAAAZg8TReiJffToVwgAAClFCA9XwItDMLEgi1UMM/+JRdiLQzSJRdyLQziJReCLQzyLWyyJReSLQiCLUiRmDxNF6Oh/1f//C/iNRdhQC9qJffBWiV30VugKCAAAi10MKUUIM8ALQziLSzSLUyyLezCJRdwzwAtDPIlN2ItLIIlF4ItDKIlN5LEg6BjV//8LQySJReiNRdhQVlaJVezHRfAAAAAAiX306LoHAACLfQgr+MdF4AAAAACLQziJRdiLQzyJRdyLQySJReSLQyiJReiLQyyJReyLQzSJRfSNRdhQVlbHRfAAAAAA6HgHAAAr+HkkDx9AAOg7yP//UFZW6OP6//8D+HjvX15bi+VdwggAZg8fRAAAhf91EVboFsj//1DoAPz//4P4AXTc6AbI//9QVlboLgcAACv469rMzMzMzMzMzMzMVYvsi1UMgeyIAAAAM8lmkIsEygtEygR1RkGD+QRy8YtFCMcAAAAAAMdABAAAAADHQAgAAAAAx0AMAAAAAMdAEAAAAADHQBQAAAAAx0AYAAAAAMdAHAAAAACL5V3CDACLAg9XwImFeP///4tCBImFfP///4tCCIlFgItCDIlFhItCEIlFiItCFIlFjItCGIlFkItCHIlFlFaLdRBXM/9mDxNF4GYPE0XoiwaJRZiLRgSJRZyLRgiJRaCLRgyJRaSLRhCJRaiLRhSJRayLRhiJRbCLRhyJRbSNRZhQjYV4////Zg8TRfBQx0XYAQAAAIl93GYPE0W4Zg8TRcBmDxNFyGYPE0XQ6Nb6//+L0IXSD4S6AQAAU2ZmZg8fhAAAAAAAi414////D1fAg+EBZg8TRfiDyQB1L42FeP///1DovgMAAItF2IPgAYPIAA+EtgAAAFaNR" & _
                                                "dhQUOhE+f//i/iL2umoAAAAi0WYg+ABg8gAdSyNRZhQ6IcDAACLRbiD4AGDyAAPhAgBAABWjUW4UFDoDfn//4v4i9rp+gAAAIXSD46MAAAAjUWYUI2FeP///1BQ6GsFAACNhXj///9Q6D8DAACNRbhQjUXYUOgS+v//hcB5C1aNRdhQUOjD+P//jUW4UI1F2FBQ6DUFAACLRdiD4AGDyAB0EVaNRdhQUOif+P//i/iL2usGi138i334jUXYUOjqAgAAC/sPhJIAAACLRfCBTfQAAACAiUXw6YAAAACNhXj///9QjUWYUFDo3wQAAI1FmFDotgIAAI1F2FCNRbhQ6In5//+FwHkLVo1FuFBQ6Dr4//+NRdhQjUW4UFDorAQAAItFuIPgAYPIAHQRVo1FuFBQ6Bb4//+L+Iva6waLXfyLffiNRbhQ6GECAAAL+3QNi0XQgU3UAAAAgIlF0I1FmFCNhXj///9Q6CD5//+L0IXSD4VW/v//i33cW4tNCItF2IkBi0XgiUEIi0XkiUEMi0XoiUEQi0XsiUEUi0XwiXkEiUEYi0X0X4lBHF6L5V3CDADMzMzMzMzMzMzMzFWL7IPsVFMPV8AzyWYPE0XUi0XYVmYPE0XMi13QiUX4i0XUV4t9zIlN7IlF/A8fADP2jUH9g/kED1fAZg8TRfCLVfQPQ/A78Q+HLwEAAIvBiVXoi00QK8aNBMGLTfCJTfSLTeyJReSD/gQPg9gAAAD/cAT/MItFDP908AT/NPCNRaxQ6Cjs//8PEABmD37BDxFFzGYPc9gEA89mD37AiU28E8OJRcA7w3cPcgQ7z3MJuQEAAAAz0usOD1fAZg8TRdyLVeCLTdyLfdQDz4tF2BPQA038iU3EE1X4iVXIDxBFvA" & _
                                                "8RRcw70HcPcgQ7z3MJuAEAAAAzyesOD1fAZg8TRdyLTeCLRdwBRfSLVeiLReQT0YtN7EaLfcyD6AiJVeiJReQ78XcUi13YiV34i13UiV38i13Q6S7///+LRdiLXdCJRfiLRdSJRfyLRfSLdQiJPM6LffyJXM4EQYtd+IlF/IlV+IlN7IP5Bw+Cwv7//4l+OF+JXjxeW4vlXcIMAItF8OvJzMzMzMzMzMzMzMzMzMxVi+yLVQi4AwAAAA8fRAAAiwzCC0zCBHUFg+gBefKNSAGFyXUGM8BdwgQAVot0yviLVMr8i85XM/8LynQQDx8AD6zWAUfR6ovOC8p188HgBgPHX15dwgQAzMzMzMzMzMxVi+yD7AiLRQgPV8BTi9hmDxNF+IPAIDvDdjiLTfhWV4t9/IlNCItw+IPoCIvOi1AED6zRAQtNCNHqC9eJCIv+iVAEwecfx0UIAAAAADvDd9VfXluL5V3CBADMzMzMzMxVi+yD7FgPV8AzyVNmDxNF0ItF1GYPE0XIi1XMi13IVolF8ItF0FeJTdiJRfSJVewz9o1B/YP5BA9XwGYPE0XgD0PwO/EPh2sBAACLfQyLwSvGjQTHi33kiX38i33giUXciX34i/kr/jv3D4cLAQAA/3AE/zCLRQz/dPAE/zTwjUWoUOjY6f//DxAADxFFyItVzDv3czGLTdSL+ovBwegfAUX4i0XQg1X8AA+kwQHB7x8DwAv4M8ALwYlF5ItFyA+kwgEDwOsMi0XUi33QiUXki0XIA8OJRbgTVeyJVbw7Vex3D3IEO8NzCbgBAAAAM8nrDg9XwGYPE0Xoi03si0XoA8cTTeQDRfSJRcATTfCJTcQPEEW4DxFFyDtN5HcPcgQ7x3MJuAEAAAAzyesOD1f" & _
                                                "AZg8TReCLTeSLReABRfiLRdwRTfxGi03Yg+gIi13IiUXcO/F3F4tV1IlV8ItV0IlV9ItVzIlV7On4/v//i0XUi1XMiUXwi0XQiUX0i0X4i338i3UIiRzOi130iVTOBEGLVfCJVeyJRfSJffCJTdiD+QcPgon+//9fiV44iVY8XluL5V3CCACLfeSLReDrw8zMVYvsg+wMU4tdDA9XwFZXi30QixOL8otDBIvIZg8TRfQrNxtPBDvydQY7yHUE6xg7yHIPdwQ78nYJuAEAAAAz0usLZg8TRfSLRfSLVfiLfQiJTwSLTRCJN4tzCIl1+CtxCItLDItdEBtLDCvwi10MG8o7dfh1BTtLDHQgO0sMchB3BTtzCHYJuAEAAAAz0usLZg8TRfSLVfiLRfSJTwyLTRCJdwiLcxCJdfwrcRCLSxSLXRAbSxQr8ItdDBvKO3X8dQU7SxR0IDtLFHIQdwU7cxB2CbgBAAAAM9LrC2YPE0X0i1X4i0X0iU8UiXcQi0sYi/GLfRCLWxyJTQyLTRArcRiLyxtPHCvwi30IG8o7dQx1BDvLdCw7y3IddwU7dQx2Fol3GLgBAAAAiU8cM9JfXluL5V3CDABmDxNF9ItV+ItF9Il3GIlPHF9eW4vlXcIMAAAA" ' 16563, 17.4.2020 16:49:09
Private Const CF_SHA256_HASHSZ      As Long = 32
Private Const CF_SHA256_BLOCKSZ     As Long = 64
Private Const CF_SHA384_HASHSZ      As Long = 48
Private Const CF_SHA384_BLOCKSZ     As Long = 128
Private Const CF_SHA384_CONTEXTSZ   As Long = 200
Private Const LNG_HMAC_INNER_PAD    As Long = &H36
Private Const LNG_HMAC_OUTER_PAD    As Long = &H5C

Private m_uEcc                  As UcsEccThunkData

Private Enum UcsEccPfnIndexEnum
    ucsPfnSecp256r1MakeKey
    ucsPfnSecp256r1SharedSecret
    ucsPfnCurve25519ScalarMultiply
    ucsPfnCurve25519ScalarMultBase
    ucsPfnSha256Init
    ucsPfnSha256Update
    ucsPfnSha256Final
    ucsPfnSha384Init
    ucsPfnSha384Update
    ucsPfnSha384Final
    [_ucsPfnMax]
End Enum

Private Type UcsEccThunkData
    Thunk               As Long
    Glob                As Long
    Pfn(0 To [_ucsPfnMax] - 1) As Long
    EccKeySize          As Long
    HashCtx(0 To CF_SHA384_CONTEXTSZ - 1) As Byte
    HashPad(0 To CF_SHA384_BLOCKSZ - 1) As Byte
    HashFinal(0 To CF_SHA384_HASHSZ - 1) As Byte
End Type

'=========================================================================
' Functions
'=========================================================================

Public Function EccInit() As Boolean
'    Dim baThunk()       As Byte
    Dim lOffset         As Long
    Dim lIdx            As Long
    
    If m_uEcc.Thunk = 0 Then
    With m_uEcc
        .EccKeySize = 32
        '--- prepare thunk/context in executable memory
        .Thunk = pvThunkAllocate(STR_ECC_THUNK1 & STR_ECC_THUNK2)
        .Glob = pvThunkAllocate(STR_ECC_GLOB)
        If .Thunk = 0 Or .Glob = 0 Then
            GoTo QH
        End If
        '--- init pfns from thunk addr + offsets stored at beginning of it
        For lIdx = 0 To UBound(.Pfn)
            Call CopyMemory(lOffset, ByVal UnsignedAdd(.Thunk, 4 * lIdx), 4)
            .Pfn(lIdx) = UnsignedAdd(.Thunk, lOffset)
        Next
        '--- init pfns trampolines
        Call pvPatchProto(AddressOf pvEccCallSecp256r1MakeKey)
        Call pvPatchProto(AddressOf pvEccCallSecp256r1SharedSecret)
        Call pvPatchProto(AddressOf pvEccCallCurve25519Multiply)
        Call pvPatchProto(AddressOf pvEccCallCurve25519MulBase)
        Call pvPatchProto(AddressOf pvHashCallSha256Init)
        Call pvPatchProto(AddressOf pvHashCallSha256Update)
        Call pvPatchProto(AddressOf pvHashCallSha256Final)
        Call pvPatchProto(AddressOf pvHashCallSha384Init)
        Call pvPatchProto(AddressOf pvHashCallSha384Update)
        Call pvPatchProto(AddressOf pvHashCallSha384Final)
        '--- init thunk's first 4 bytes -> global data in C/C++
        Call CopyMemory(ByVal .Thunk, .Glob, 4)
    End With
    End If
    '--- success
    EccInit = True
QH:
End Function

Public Function EccSecp256r1MakeKey(baPrivate() As Byte, baPublic() As Byte) As Boolean
    Const MAX_RETRIES   As Long = 16
    Dim lIdx            As Long
    
    ReDim baPrivate(0 To m_uEcc.EccKeySize - 1) As Byte
    ReDim baPublic(0 To m_uEcc.EccKeySize) As Byte
    For lIdx = 1 To MAX_RETRIES
        pvEccRandomBytes VarPtr(baPrivate(0)), m_uEcc.EccKeySize
        If pvEccCallSecp256r1MakeKey(m_uEcc.Pfn(ucsPfnSecp256r1MakeKey), VarPtr(baPublic(0)), VarPtr(baPrivate(0))) = 1 Then
            Exit For
        End If
    Next
    '--- success (or failure)
    EccSecp256r1MakeKey = (lIdx <= MAX_RETRIES)
End Function

Public Function EccSecp256r1SharedSecret(baPrivate() As Byte, baPublic() As Byte) As Byte()
    Dim baRetVal()      As Byte
    
    Debug.Assert UBound(baPrivate) >= m_uEcc.EccKeySize - 1
    Debug.Assert UBound(baPublic) >= m_uEcc.EccKeySize
    ReDim baRetVal(0 To m_uEcc.EccKeySize - 1) As Byte
    If pvEccCallSecp256r1SharedSecret(m_uEcc.Pfn(ucsPfnSecp256r1SharedSecret), VarPtr(baPublic(0)), VarPtr(baPrivate(0)), VarPtr(baRetVal(0))) = 0 Then
        GoTo QH
    End If
    EccSecp256r1SharedSecret = baRetVal
QH:
End Function

Public Function EccCurve25519MakeKey(baPrivate() As Byte, baPublic() As Byte) As Boolean
    ReDim baPrivate(0 To m_uEcc.EccKeySize - 1) As Byte
    ReDim baPublic(0 To m_uEcc.EccKeySize - 1) As Byte
    pvEccRandomBytes VarPtr(baPrivate(0)), m_uEcc.EccKeySize
    baPrivate(0) = baPrivate(0) And 248
    baPrivate(UBound(baPrivate)) = (baPrivate(UBound(baPrivate)) And 127) Or 64
    pvEccCallCurve25519MulBase m_uEcc.Pfn(ucsPfnCurve25519ScalarMultBase), VarPtr(baPublic(0)), VarPtr(baPrivate(0))
    '--- success
    EccCurve25519MakeKey = True
End Function

Public Function EccCurve25519SharedSecret(baPrivate() As Byte, baPublic() As Byte) As Byte()
    Dim baRetVal()      As Byte
    
    Debug.Assert UBound(baPrivate) >= m_uEcc.EccKeySize - 1
    Debug.Assert UBound(baPublic) >= m_uEcc.EccKeySize - 1
    ReDim baRetVal(0 To m_uEcc.EccKeySize - 1) As Byte
    pvEccCallSecp256r1SharedSecret m_uEcc.Pfn(ucsPfnCurve25519ScalarMultiply), VarPtr(baRetVal(0)), VarPtr(baPrivate(0)), VarPtr(baPublic(0))
    EccCurve25519SharedSecret = baRetVal
End Function

Public Function HashSha256(baInput() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Byte()
    Dim lCtxPtr         As Long
    Dim lPtr            As Long
    Dim baRetVal()      As Byte
    
    With m_uEcc
        lCtxPtr = VarPtr(.HashCtx(0))
        If lSize > 0 Then
            lPtr = VarPtr(baInput(lPos))
        End If
        pvHashCallSha256Init .Pfn(ucsPfnSha256Init), lCtxPtr
        pvHashCallSha256Update .Pfn(ucsPfnSha256Update), lCtxPtr, lPtr, lSize
        ReDim baRetVal(0 To CF_SHA256_HASHSZ - 1) As Byte
        pvHashCallSha256Final .Pfn(ucsPfnSha256Final), lCtxPtr, VarPtr(baRetVal(0))
    End With
    HashSha256 = baRetVal
End Function

Public Function HashSha384(baInput() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Byte()
    Dim lCtxPtr         As Long
    Dim lPtr            As Long
    Dim baRetVal()      As Byte
    
    With m_uEcc
        lCtxPtr = VarPtr(.HashCtx(0))
        If lSize > 0 Then
            lPtr = VarPtr(baInput(lPos))
        End If
        pvHashCallSha384Init .Pfn(ucsPfnSha384Init), lCtxPtr
        pvHashCallSha384Update .Pfn(ucsPfnSha384Update), lCtxPtr, lPtr, lSize
        ReDim baRetVal(0 To CF_SHA384_HASHSZ - 1) As Byte
        pvHashCallSha384Final .Pfn(ucsPfnSha384Final), lCtxPtr, VarPtr(baRetVal(0))
    End With
    HashSha384 = baRetVal
End Function

Public Function HmacSha256(baKey() As Byte, baInput() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Byte()
    Dim lCtxPtr         As Long
    Dim lPtr            As Long
    Dim lIdx            As Long
    
    Debug.Assert UBound(baKey) < CF_SHA256_BLOCKSZ
    With m_uEcc
        lCtxPtr = VarPtr(.HashCtx(0))
        If lSize > 0 Then
            lPtr = VarPtr(baInput(lPos))
        End If
        '-- inner hash
        pvHashCallSha256Init .Pfn(ucsPfnSha256Init), lCtxPtr
        Call FillMemory(.HashPad(0), CF_SHA256_BLOCKSZ, LNG_HMAC_INNER_PAD)
        For lIdx = 0 To UBound(baKey)
            .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_INNER_PAD
        Next
        pvHashCallSha256Update .Pfn(ucsPfnSha256Update), lCtxPtr, VarPtr(.HashPad(0)), CF_SHA256_BLOCKSZ
        pvHashCallSha256Update .Pfn(ucsPfnSha256Update), lCtxPtr, lPtr, lSize
        pvHashCallSha256Final .Pfn(ucsPfnSha256Final), lCtxPtr, VarPtr(.HashFinal(0))
        '-- outer hash
        pvHashCallSha256Init .Pfn(ucsPfnSha256Init), lCtxPtr
        Call FillMemory(.HashPad(0), CF_SHA256_BLOCKSZ, LNG_HMAC_OUTER_PAD)
        For lIdx = 0 To UBound(baKey)
            .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_OUTER_PAD
        Next
        pvHashCallSha256Update .Pfn(ucsPfnSha256Update), lCtxPtr, VarPtr(.HashPad(0)), CF_SHA256_BLOCKSZ
        pvHashCallSha256Update .Pfn(ucsPfnSha256Update), lCtxPtr, VarPtr(.HashFinal(0)), CF_SHA256_HASHSZ
        ReDim baRetVal(0 To CF_SHA256_HASHSZ - 1) As Byte
        pvHashCallSha256Final .Pfn(ucsPfnSha256Final), lCtxPtr, VarPtr(baRetVal(0))
    End With
    HmacSha256 = baRetVal
End Function

Public Function HmacSha384(baKey() As Byte, baInput() As Byte, ByVal lPos As Long, ByVal lSize As Long) As Byte()
    Dim lCtxPtr         As Long
    Dim lPtr            As Long
    Dim lIdx            As Long
    
    Debug.Assert UBound(baKey) < CF_SHA384_BLOCKSZ
    With m_uEcc
        lCtxPtr = VarPtr(.HashCtx(0))
        If lSize > 0 Then
            lPtr = VarPtr(baInput(lPos))
        End If
        '-- inner hash
        pvHashCallSha384Init .Pfn(ucsPfnSha384Init), lCtxPtr
        Call FillMemory(.HashPad(0), CF_SHA384_BLOCKSZ, LNG_HMAC_INNER_PAD)
        For lIdx = 0 To UBound(baKey)
            .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_INNER_PAD
        Next
        pvHashCallSha384Update .Pfn(ucsPfnSha384Update), lCtxPtr, VarPtr(.HashPad(0)), CF_SHA384_BLOCKSZ
        pvHashCallSha384Update .Pfn(ucsPfnSha384Update), lCtxPtr, lPtr, lSize
        pvHashCallSha384Final .Pfn(ucsPfnSha384Final), lCtxPtr, VarPtr(.HashFinal(0))
        '-- outer hash
        pvHashCallSha384Init .Pfn(ucsPfnSha384Init), lCtxPtr
        Call FillMemory(.HashPad(0), CF_SHA384_BLOCKSZ, LNG_HMAC_OUTER_PAD)
        For lIdx = 0 To UBound(baKey)
            .HashPad(lIdx) = baKey(lIdx) Xor LNG_HMAC_OUTER_PAD
        Next
        pvHashCallSha384Update .Pfn(ucsPfnSha384Update), lCtxPtr, VarPtr(.HashPad(0)), CF_SHA384_BLOCKSZ
        pvHashCallSha384Update .Pfn(ucsPfnSha384Update), lCtxPtr, VarPtr(.HashFinal(0)), CF_SHA384_HASHSZ
        ReDim baRetVal(0 To CF_SHA384_HASHSZ - 1) As Byte
        pvHashCallSha384Final .Pfn(ucsPfnSha384Final), lCtxPtr, VarPtr(baRetVal(0))
    End With
    HmacSha384 = baRetVal
End Function

'= private ===============================================================

Private Function pvEccCallSecp256r1MakeKey(ByVal Pfn As Long, ByVal lPubKeyPtr As Long, ByVal lPrivKeyPtr As Long) As Long
    ' int ecc_make_key(uint8_t p_publicKey[ECC_BYTES+1], uint8_t p_privateKey[ECC_BYTES]);
End Function

Private Function pvEccCallSecp256r1SharedSecret(ByVal Pfn As Long, ByVal lPubKeyPtr As Long, ByVal lPrivKeyPtr As Long, ByVal lSecretPtr As Long) As Long
    ' int ecdh_shared_secret(const uint8_t p_publicKey[ECC_BYTES+1], const uint8_t p_privateKey[ECC_BYTES], uint8_t p_secret[ECC_BYTES]);
End Function

Private Function pvEccCallCurve25519Multiply(ByVal Pfn As Long, ByVal lSecretPtr As Long, ByVal lPubKeyPtr As Long, ByVal lPrivKeyPtr As Long) As Long
    ' void cf_curve25519_mul(uint8_t out[32], const uint8_t priv[32], const uint8_t pub[32])
End Function

Private Function pvEccCallCurve25519MulBase(ByVal Pfn As Long, ByVal lPubKeyPtr As Long, ByVal lPrivKeyPtr As Long) As Long
    ' void cf_curve25519_mul_base(uint8_t out[32], const uint8_t priv[32])
End Function

Private Function pvHashCallSha256Init(ByVal Pfn As Long, ByVal lCtxPtr As Long) As Long
    ' void cf_sha256_init(cf_sha256_context *ctx)
End Function

Private Function pvHashCallSha256Update(ByVal Pfn As Long, ByVal lCtxPtr As Long, ByVal lDataPtr As Long, ByVal lSize As Long) As Long
    ' void cf_sha256_update(cf_sha256_context *ctx, const void *data, size_t nbytes)
End Function

Private Function pvHashCallSha256Final(ByVal Pfn As Long, ByVal lCtxPtr As Long, ByVal lHashPtr As Long) As Long
    ' void cf_sha256_digest_final(cf_sha256_context *ctx, uint8_t hash[CF_SHA256_HASHSZ])
End Function

Private Function pvHashCallSha384Init(ByVal Pfn As Long, ByVal lCtxPtr As Long) As Long
    ' void cf_sha384_init(cf_sha384_context *ctx)
End Function

Private Function pvHashCallSha384Update(ByVal Pfn As Long, ByVal lCtxPtr As Long, ByVal lDataPtr As Long, ByVal lSize As Long) As Long
    ' void cf_sha384_update(cf_sha384_context *ctx, const void *data, size_t nbytes)
End Function

Private Function pvHashCallSha384Final(ByVal Pfn As Long, ByVal lCtxPtr As Long, ByVal lHashPtr As Long) As Long
    ' void cf_sha384_digest_final(cf_sha384_context *ctx, uint8_t hash[CF_SHA384_HASHSZ])
End Function

Private Sub pvEccRandomBytes(ByVal lPtr As Long, ByVal lSize As Long)
    Dim hProv           As Long
    
    If CryptAcquireContext(hProv, 0, 0, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0 Then
        Call CryptGenRandom(hProv, lSize, lPtr)
        Call CryptReleaseContext(hProv, 0)
    End If
End Sub

Private Function pvThunkAllocate(sText As String, Optional ByVal Size As Long) As Long
    Static Map(0 To &H3FF) As Long
    Dim baInput()       As Byte
    Dim lIdx            As Long
    Dim lChar           As Long
    Dim lPtr            As Long
    
    pvThunkAllocate = VirtualAlloc(0, IIf(Size > 0, Size, (Len(sText) \ 4) * 3), MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If pvThunkAllocate = 0 Then
        Exit Function
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
    baInput = StrConv(Replace(Replace(sText, vbCr, vbNullString), vbLf, vbNullString) & "===", vbFromUnicode)
    lPtr = pvThunkAllocate
    For lIdx = 0 To UBound(baInput) - 3 Step 4
        lChar = Map(baInput(lIdx + 0)) Or Map(&H100 + baInput(lIdx + 1)) Or Map(&H200 + baInput(lIdx + 2)) Or Map(&H300 + baInput(lIdx + 3))
        Call CopyMemory(ByVal lPtr, lChar, 3)
        lPtr = UnsignedAdd(lPtr, 3)
    Next
End Function

Private Sub pvPatchProto(ByVal Pfn As Long)
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

Private Function pvSetTrue(bValue As Boolean) As Boolean
    bValue = True
    pvSetTrue = True
End Function

Private Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function
