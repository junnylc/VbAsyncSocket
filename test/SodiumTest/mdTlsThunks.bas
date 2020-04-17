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
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Private Declare Function CryptAcquireContext Lib "advapi32" Alias "CryptAcquireContextW" (phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As Long, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGenRandom Lib "advapi32" (ByVal hProv As Long, ByVal dwLen As Long, ByVal pbBuffer As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_ECC_CTX           As String = "////////////////AAAAAAAAAAAAAAAAAQAAAP////9LYNInPjzOO/awU8ywBh1lvIaYdlW967Pnkzqq2DXGWpbCmNhFOaH0oDPrLYF9A3fyQKRj5ea8+EdCLOHy0Rdr9VG/N2hAtsvOXjFrVzPOKxaeD3xK6+eOm38a/uJC409RJWP8wsq584SeF6et+ua8//////////8AAAAA/////w==" ' 15.4.2020 13:27:58
Private Const STR_ECC_THUNK1        As String = "EBQAADAXAACgDgAAwBEAAOgAAAAAWIPoFYsAw8zMzMxVi+yB7KgAAAAzyVaLdRCQiwTOC0TOBHUNQYP5BHLxXovlXcIMAFNXi30MjYVg////V1Do2CwAAI2FYP///1CNReBQ6KgkAACLXQiNReBQU41FoFDoZyoAAI1FoFCNRYBQ6IokAACNReBQjUWgUOidLAAAjUWgUI1F4FDocCQAAFZXjUWgUOg1KgAAjUWgUFfoWyQAAFaNRaBQ6HEsAACNRaBQVuhHJAAA6EL///9WU1OL+OiYIQAAC8J1C1dT6M0iAACFwHgIV1NT6AEuAADoHP///1ZWVov46HIhAAALwnULV1bopyIAAIXAeAhXVlbo2y0AAOj2/v//VlNWi/jozC0AAAvCdAhXVlboQCEAAFZTjUWgUOilKQAAjUWgUFPoyyMAAOjG/v//U1NWi/joHCEAAAvCdQtXVuhRIgAAhcB4CFdWVuiFLQAA6KD+//9WU1OL+Oj2IAAAC8J1C1dT6CsiAACFwHgIV1NT6F8tAACLA4PgAYPIAHQd6HD+//9QU1PoyCAAAFOL+OggKwAAwecfCXsc6wZT6BIrAABTjUWgUOhoKwAAjUWgUFboPiMAAOg5/v//i/iNRYBQVlboDC0AAAvCdAhXVlbogCAAAOgb/v//i/iNRYBQVlbo7iwAAAvCdAhXVlboYiAAAOj9/f//i/iNRYBWUFDo0CwAAAvCdAtXjUWAUFDoQSAAAI1FgFBTjUWgUOijKAAAjUWgUFPoySIAAOjE/f//i/iNReBQU1DolywAAAvCdAtXjUXgUFDoCCAAAIsGiQOLRgSLTQyJQwSLRgiJQwiLRgyJQwyLRhCJQxCLRhSJQxSLRhiJQxiLRhyJQxyLAYkGi0EEiUYE" & _
                                                "i0EIiUYIi0EMiUYMi0EQiUYQi0EUiUYUi0EYiUYYi0EciUYci0XgiQGLReSJQQSLReiJQQiLReyJQQyLRfCJQRCLRfSJQRSLRfhfiUEYi0X8W4lBHF6L5V3CDADMVYvsVot1CDPADx+AAAAAAIsMxgtMxgR1JUCD+ARy8TPSjU4giwELQQR1E0KDwQiD+gRy8LgBAAAAXl3CBAAzwF5dwgQAzMzMzMzMzMzMzMxVi+yB7CQBAABTi10MVlf/dRSLA4lFvItDBIlFwItDCIlFxItDDIlFyItDEIlFzItDFIlF0ItDGIlF1ItDHIlF2ItDIImFfP///4tDJIlFgItDKIlFhItDLIlFiItDMIlFjItDNIlFkItDOIlFlItDPIlFmI2FXP///1CNRZxQjYV8////UI1FvFDoYQcAAIt9EFfoiCgAAIPoAolFFIXAfnYz9ovIg+E/M9IPq86D+SAPQ9Yz8oP5QA9D1sHoBiM0xyNUxwQL8nUFjUYB6wIzwMHgBY2dXP///wPYjU2cA8iNtXz///9T99iJTfxRA/CNfbwD+FZX6BQEAABWV1P/dfzoSQIAAItFFIt9EEiJRRSFwH+Ni10MiweD4AGDyAB1B7gBAAAA6wIzwMHgBY2NXP///wPIjVWcUQPQiU0MjbV8////iVUUUivwjX28K/hWV+i5AwAA6GT7//+JRRCNRZxQjUW8UI1F3FDoMCoAAAvCdA3/dRCNRdxQUOifHQAAVo1F3FCNhRz///9Q6P4lAACNhRz///9QjUXcUOgeIAAAU41F3FCNhRz///9Q6N0lAACNhRz///9QjUXcUOj9HwAA6Pj6//9QjUXcUFDorSIAAI1DIFCNRdxQjYUc////UOipJQAAjYUc////UI1F3" & _
                                                "FDoyR8AAFeNRdxQjYUc////UOiIJQAAjYUc////UI1F3FDoqB8AAFZX/3UM/3UU6CsBAACNRdxQjYXc/v//UOirJwAAjYXc/v//UI2FPP///1DoeB8AAI2FPP///1CNRZxQjYXc/v//UOgxJQAAjYXc/v//UI1FnFDoUR8AAI1F3FCNhTz///9QjYXc/v//UOgKJQAAjYXc/v//UI2FPP///1DoJx8AAI2FPP///1CNhVz///9QjYXc/v//UOjdJAAAjYXc/v//UI2FXP///1Do+h4AAItNCItFnF9eiQGLRaCJQQSLRaSJQQiLRaiJQQyLRayJQRCLRbCJQRSLRbSJQRiLRbiJQRyLhVz///+JQSCLhWD///+JQSSLhWT///+JQSiLhWj///+JQSyLhWz///+JQTCLhXD///+JQTSLhXT///+JQTiLhXj///+JQTxbi+VdwhAAzMzMzMzMzFWL7IPsYFNWV+hi+f//i10Ii/iLdRCNReBTVlDoLygAAAvCdAtXjUXgUFDooBsAAI1F4FCNRaBQ6FMmAACNRaBQjUXgUOgmHgAAjUXgUFONRaBQ6OgjAACNRaBQU+gOHgAAjUXgUFaNRaBQ6NAjAACNRaBQVuj2HQAA6PH4////dQyLfRRXV4lFCOjBJwAAC8J0Cv91CFdX6DMbAABXjUWgUOjpJQAAjUWgUI1F4FDovB0AAOi3+P//iUUIjUXgU1BQ6IknAAALwnQN/3UIjUXgUFDo+BoAAOiT+P//iUUIjUXgVlBQ6GUnAAALwnQN/3UIjUXgUFDo1BoAAOhv+P//U1ZWiUUI6EQnAAALwnQK/3UIVlbothoAAFb/dQyNRaBQ6BkjAACNRaBQ/3UM6D0dAADoOPj//4lFCI1F4F" & _
                                                "BTVugKJwAAC8J0Cv91CFZW6HwaAABWV41FoFDo4SIAAI1FoFBX6AcdAADoAvj///91DIvYV1fo1iYAAAvCdAhTV1foShoAAItF4IkGi0XkiUYEi0XoiUYIi0XsiUYMi0XwiUYQi0X0iUYUi0X4iUYYi0X8X4lGHF5bi+VdwhAAzMxVi+yB7OAAAABTVlfon/f//4t1CIv4i10QjUXgVlNQ6GwmAAALwnQLV41F4FBQ6N0ZAACNReBQjYVg////UOiNJAAAjYVg////UI1F4FDoXRwAAI1F4FBWjYVg////UOgcIgAAjYVg////UFboPxwAAI1F4FBTjYVg////UOj+IQAAjYVg////UFPoIRwAAOgc9////3UMi30UiUUIjUXgV1DoaRkAAAvCdRD/dQiNReBQ6JkaAACFwHgN/3UIjUXgUFDoyCUAAOjj9v///3UMiUUIV1fotiUAAAvCdAr/dQhXV+goGQAA6MP2//9WiUUIjUWAU1DolSUAAAvCdA3/dQiNRYBQUOgEGQAAjUWAUP91DI1FoFDoZCEAAI1FoFD/dQzoiBsAAOiD9v//U4lFCI1FgFZQ6NUYAAALwnUQ/3UIjUWAUOgFGgAAhcB4Df91CI1FgFBQ6DQlAABXjUWgUOhqIwAAjUWgUFPoQBsAAOg79v//iUUIjUWAUFNT6A0lAAALwnQK/3UIU1PofxgAAOga9v//U4lFCI1FwFZQ6OwkAAALwnQN/3UIjUXAUFDoWxgAAI1FwFBXjYUg////UOi6IAAAjYUg////UFfo3RoAAOjY9f//i10MU1dXiUUI6KokAAALwnQK/3UIV1foHBgAAI1F4FCNhSD///9Q6MwiAACNhSD///9QjUXAUOicGgAA6Jf1//+L+I1" & _
                                                "FgFCNRcBQUOhnJAAAC8J0C1eNRcBQUOjYFwAA6HP1//+L+I1FwFZQjUWAUOhDJAAAC8J0C1eNRYBQUOi0FwAAjUXgUI1FgFCNhSD///9Q6BAgAACNhSD///9QjUWAUOgwGgAA6Cv1//+L+I1FgFNQU+j+IwAAC8J0CFdTU+hyFwAAi0XAiQaLRcSJRgSLRciJRgiLRcyJRgyLRdCJRhCLRdSJRhSLRdiJRhiLRdxfiUYcXluL5V3CEADMzMzMzMzMzMzMVYvsg+wgi00UD1fAU4tdEFaLdQhXi30MZg8TReiLBokDi0YEiUMEi0YIiUMIi0YMiUMMi0YQiUMQi0YUiUMUi0YYiUMYi0YciUMciweJAYtHBIlBBItHCIlBCItHDIlBDItHEIlBEItHFIlBFItHGIlBGItHHIlBHItNGGYPE0XwZg8TRfjHReABAAAAx0XkAAAAAIXJdC+LAYlF4ItBBIlF5ItBCIlF6ItBDIlF7ItBEIlF8ItBFIlF9ItBGIlF+ItBHIlF/I1F4FBXVug+AQAAjUXgUFdW6AP0//+NReBQ/3UUU+gmAQAAX15bi+VdwhQAzMzMzMzMzMzMzMzMzFOLRCQMi0wkEPfhi9iLRCQI92QkFAPYi0QkCPfhA9NbwhAAzMzMzMzMzMzMzMzMzID5QHMVgPkgcwYPpcLT4MOL0DPAgOEf0+LDM8Az0sPMgPlAcxWA+SBzBg+t0NPqw4vCM9KA4R/T6MMzwDPSw8xVi+yLRRBTVot1CI1IeFeLfQyNVng78XcEO9BzC41PeDvxdzA713IsK/i7EAAAACvwixQ4AxCLTDgEE0gEjUAIiVQw+IlMMPyD6wF15F9eW13CDACL141IEIveK9Ar2Cv+uAQAAACNdiCN" & _
                                                "SSAPEEHQDxBMN+BmD9TIDxFO4A8QTArgDxBB4GYP1MgPEUwL4IPoAXXSX15bXcIMAMzMzMzMVYvsg+xgjUWgVv91EFDozR8AAI1FoFCNReBQ6KAXAACLdQiNReBQVo1FoFDoXx0AAI1FoFBW6IUXAAD/dRCNReBQjUWgUOhFHQAAjUWgUI1F4FDoaBcAAIt1DI1F4FBWjUWgUOgnHQAAjUWgUFboTRcAAF6L5V3CDADMzMzMzMxVi+yD7CBTVot1CDPJV4lN7IEEzgAAAQCLBM6DVM4EAItczgQPrNgQwfsQiUXog/kPdRXHRfwBAAAAi9DHRfAAAAAAiV346yIPV8BmDxNF9ItF+IlF8ItF9GYPE0Xgi1XgiUX8i0XkiUX4g/kPjXkBagAbwPfYD6/HK1X8aiWNNMaLRfgbRfBQUuji/f//i03oA8ET04PoAYPaAAEGi0XsEVYEi3UID6TLEMHhECkMxovPiU3sGVzGBIP5EA+CT////19eW4vlXcIEAMzMzMzMVYvsgewoBAAAU1ZXanCNhej8///Hhdj8//9B2wAAagBQx4Xc/P//AAAAAMeF4Pz//wEAAADHheT8//8AAAAA6CwLAACLdQyNhWD///9qH1ZQ6OoKAACKRh+DxBiApWD////4JD8MQIiFf////42F2Pv///91EFDohBIAAA9XwI21WP7//2YPE4VY/v//jb1g/v//uR4AAABmDxNFgPOluR4AAABmDxOF4P7//411gMeFWP7//wEAAACNfYjHhVz+//8AAAAA86W5HgAAAMdFgAEAAACNteD+///HRYQAAAAAjb3o/v//u/4AAADzpbkgAAAAjbXY+///jb3Y/f//86WLw4rLwfgDgOEHioQFYP///9LoJAEPt" & _
                                                "sCZi/CNhdj9//9WUI1FgFDoQxAAAFaNhVj+//9QjYXg/v//UOgvEAAAjYXg/v//UI1FgFCNhVj9//9Q6Lj8//+NheD+//9QjUWAUFDo9xAAAI2FWP7//1CNhdj9//9QjYXg/v//UOiN/P//jYVY/v//UI2F2P3//1BQ6MkQAACNhVj9//9QUI2FWP7//1DoBQsAAI1FgFBQjYVY/P//UOj0CgAAjUWAUI2F4P7//1CNRYBQ6OAKAACNhVj9//9QjYXY/f//UI2F4P7//1DoxgoAAI2F4P7//1CNRYBQjYVY/f//UOgP/P//jYXg/v//UI1FgFBQ6E4QAACNRYBQUI2F2P3//1DojQoAAI2FWPz//1CNhVj+//9QjYXg/v//UOgjEAAAjYXY/P//UI2F4P7//1CNRYBQ6FwKAACNhVj+//9QjUWAUFDoq/v//41FgFCNheD+//9QUOg6CgAAjYVY/P//UI2FWP7//1CNRYBQ6CMKAACNhdj7//9QjYXY/f//UI2FWP7//1DoCQoAAI2FWP3//1BQjYXY/f//UOj1CQAAVo2F2P3//1CNRYBQ6KQOAABWjYVY/v//UI2F4P7//1DokA4AAIPrAQ+JGP7//42F4P7//1BQ6CoGAACNheD+//9QjUWAUFDoqQkAAI1FgFD/dQjozQwAAF9eW4vlXcIMAMzMzMxVi+yD7CCNReDGReAJUP91DA9XwMdF+QAAAAD/dQgPEUXhZsdF/QAAZg/WRfHGRf8A6Kr8//+L5V3CCADMzMzMVYvsU4tdDFZXi30ID7ZDGJmLyIvyD6TOCA+2QxnB4QiZC8gL8g+kzggPtkMaweEImQvIC/IPpM4ID7ZDG8HhCJkLyAvyD7ZDHA+kzgiZweEIC/ILyA" & _
                                                "+2Qx0PpM4ImcHhCAvyC8gPtkMeD6TOCJnB4QgL8gvID7ZDHw+kzgiZweEIC/ILyIl3BIkPD7ZDEJmLyIvyD7ZDEQ+kzgiZweEIC/ILyA+2QxIPpM4ImcHhCAvyC8gPtkMTD6TOCJnB4QgL8gvID7ZDFA+kzgiZweEIC8gL8g+kzggPtkMVweEImQvIC/IPpM4ID7ZDFsHhCJkLyAvyD6TOCA+2QxfB4QiZC8gL8olPCIl3DA+2QwiZi8iL8g+kzggPtkMJweEImQvIC/IPtkMKD6TOCJnB4QgL8gvID7ZDCw+kzgiZweEIC/ILyA+2QwwPpM4ImcHhCAvyC8gPtkMND6TOCJnB4QgL8gvID7ZDDg+kzgiZweEIC/ILyA+2Qw8PpM4ImcHhCAvyC8iJdxSJTxAPtgOZi8iL8g+2QwEPpM4ImcHhCAvyC8gPtkMCD6TOCMHhCJkLyAvyD7ZDAw+kzgiZweEIC/ILyA+2QwQPpM4ImcHhCAvyC8gPtkMFD6TOCJnB4QgL8gvID7ZDBg+kzgiZweEIC/ILyA+2QwcPpM4ImcHhCAvIC/KJdxyJTxhfXltdwggAzMzMVYvsg+xgjUXg/3UMUOje/f//M8CLTMXgC0zF5HUOQIP4BHLwM8CL5V3CCACNReBQ6Mvr//+D6IBQ6GIPAACD+AF0E+i46///g+iAUI1F4FBQ6IoaAABqAI1F4FDon+v//4PAQFCNRaBQ6NLu//+NRaBQ6Hnu//+FwHWpikXAi00IJAEEAogBjUWgUI1BAVDoDAAAALgBAAAAi+VdwggAzFWL7FaLdQixKFeLfQwPtkcHiEYYD7ZHBohGGYsHi1cE6Lv3//+IRhqxIIsHi1cE6Kz3//+IRhuLD4tHBA+swRiIThy" & _
                                                "LD8HoGItHBA+swRCITh2LD8HoEItHBA+swQiITh6xKMHoCA+2B4hGHw+2Rw+IRhAPtkcOiEYRi0cIi1cM6Fv3//+IRhKxIItHCItXDOhL9///iEYTi08Ii0cMD6zBGIhOFItPCMHoGItHDA+swRCIThWLTwjB6BCLRwwPrMEIiE4WsSjB6AgPtkcIiEYXD7ZHF4hGCA+2RxaIRgmLRxCLVxTo9vb//4hGCrEgi0cQi1cU6Ob2//+IRguLTxCLRxQPrMEYiE4Mi08QwegYi0cUD6zBEIhODYtPEMHoEItHFA+swQiITg6xKMHoCA+2RxCIRg8PtkcfiAYPtkceiEYBi0cYi1cc6JL2//+IRgKxIItHGItXHOiC9v//iEYDi08Yi0ccD6zBGMHoGIhOBItPGItHHA+swRDB6BCITgWLTxiLRxwPrMEIwegIiE4GD7ZHGF+IRgdeXcIIAMzMVYvsg+xgi0UMD1fAU1ZXi30IQFBXx0XgAwAAAMdF5AAAAAAPEUXoZg/WRfjof/v//1eNRaBQ6KUWAACNRaBQjXcgVuh4DgAA6HPp//+L2I1F4FBWVuhGGAAAC8J0CFNWVui6CwAAV1aNRaBQ6B8UAACNRaBQVuhFDgAA6EDp//+L+Og56f//g8AgUFZW6I4LAAALwnULV1bowwwAAIXAeAhXVlbo9xcAAFboQQMAAItNDDP/igGLDiQBD7bAg+EBmTvIdQQ7+nQNVujx6P//UFboyhcAAF9eW4vlXcIIAMxVi+yB7KAAAACNhWD/////dQhQ6Aj/////dQyNRaBQ6Kz6//9qAI1FoFCNhWD///9QjUXAUOjm6///jUXAUP91EOg6/f//M8kPH4QAAAAAAItEzcALRM3EdR5Bg/kEcvCL" & _
                                                "TMXgC0zF5HUOQIP4BHLwM8CL5V3CDAC4AQAAAIvlXcIMAMzMzMzMzMzMzMzMzMxVi+yB7IQBAABTVot1DLkgAAAAV429fP7//7v9AAAA86WJXfiNhXz+//9QUFDoXgMAAIP7Ag+EvQEAAIP7BA+EtAEAAItFDI21/P7//w9XwI29BP///4PAEGYPE4X8/v//uTwAAACJRfTzpTPbDx8AjbUE////x0X8BAAAAIv4A/P/tB2A/v///7QdfP7///939P938Ojm8////7QdgP7//wFG+P+0HXz+//8RVvz/d/z/d/jox/P///+0HYD+//8BBv+0HXz+//8RVgT/dwT/N+iq8////7QdgP7//wFGCP+0HXz+//8RVgz/dwz/dwjoi/P//wFGEI1/IBFWFI12IINt/AEPhXb///+LRfSDwwiB+4AAAAAPglP///8z9pBqAGom/3T1gP+09Xz////oTPP//wGE9fz+//9qABGU9QD///9qJv909Yj/dPWE6C3z//8BhPUE////agARlPUI////aib/dPWQ/3T1jOgO8///AYT1DP///2oAEZT1EP///2om/3T1mP909ZTo7/L//wGE9RT///9qABGU9Rj///9qJv909aD/dPWc6NDy//8BhPUc////EZT1IP///4PGBYP+Dw+CVv///42FfP7//7kgAAAAjbX8/v//jb18/v//86VQ6Cf0//+NhXz+//9Q6Bv0//+LXfiD6wGJXfgPiSD+//+LfQiNtXz+//+5IAAAAPOlX15bi+VdwggAzMzMVYvsi0UIi9BWi3UQhfZ0FVeLfQwr+IoMF41SAYhK/4PuAXXyX15dw8zMzMzMzMzMVYvsi00Qhcl0Hw+2RQxWi/FpwAEBAQFXi30IwekC8" & _
                                                "6uLzoPhA/OqX16LRQhdw8zMVYvsgeyAAAAAU1YPV8DHRcABAAAAV41FwMdFxAAAAAC7AQAAAGYP1kXYUA8RRciJXeDHReQAAAAADxFF6GYP1kX46Inl//9QjUXAUOjfBwAAjUXAUOjWEQAAi30IjXD/O/N2aY1F4FCNRYBQ6H8SAACNRYBQjUXgUOhSCgAAM9IzyYvGg+A/D6vCg/ggD0PKM9GD+ECLxg9DysHoBiNUxcAjTMXEC9F0G1eNReBQjUWAUOjpDwAAjUWAUI1F4FDoDAoAAE6D/gF3motd4ItF5IlHBItF6IlHCItF7IlHDItF8IlHEItF9IlHFItF+IlHGItF/IkfiUccX15bi+VdwgQAzMzMzMzMzMzMzMzMzMxVi+yB7BQBAACLRQwPV8BTVle5PAAAAGYPE4Xs/v//jbXs/v//x0X8EAAAAI299P7///Oli00QjZ30/v//g8EQi9MrwolN+IlFDGYPH0QAAIv5x0UQBAAAAIvzDx9EAAD/dBgE/zQY/3f0/3fw6H7w//8BRviLRQwRVvz/dBgE/zQY/3f8/3f46GPw////dwQBBotFDP83EVYE/3QYBP80GOhK8P//AUYIi0UMEVYM/3QYBP80GP93DP93COgv8P//AUYQjX8gi0UMEVYUjXYgg20QAXWKi034g8MIg238AQ+Fav///zP2Dx+EAAAAAABqAGom/7T1cP////+09Wz////o6e///wGE9ez+//9qABGU9fD+//9qJv+09Xj/////tPV0////6MTv//8BhPX0/v//agARlPX4/v//aib/dPWA/7T1fP///+ii7///AYT1/P7//2oAEZT1AP///2om/3T1iP909YTog+///wGE9QT///9qABGU9Qj///"
Private Const STR_ECC_THUNK2        As String = "9qJv909ZD/dPWM6GTv//8BhPUM////EZT1EP///4PGBYP+Dw+CSv///4tVCI217P7//7kgAAAAi/rzpTPJiU34Dx8AgQTKAAABAIsEyoNUygQAi1zKBA+s2BDB+xCJRfSD+Q91FcdFDAEAAACL0MdF/AAAAACJXRDrIg9XwGYPE0Xsi0XwiUX8i0XsZg8TReSLVeSJRQyLReiJRRCD+Q+NeQGLTQgbwPfYD6/HK1UMagBqJY00wYtFEBtF/FBS6LDu//+LTfQDwRPTg+gBg9oAAQaLRfgRVgSLVQgPpMsQiX34weEQKQzCi88ZXMIEg/kQD4JM////UugG8P//X15bi+VdwgwAzMzMzMzMzMzMzMzMzFWL7IPsEFNWi3UMV4t9GGoAVmoA/3UU6ETu//9qAFZqAFeJRfCL2ug07v//agD/dRCJRfSL8moAV+gi7v//agD/dRCJRfxqAP91FIlV+OgN7v//i/iLRfQD+4PSAAP4E9Y71ncOcgQ7+HMIg0X8AINV+AGLRQgzyQtN8IkIM8kDVfyJeAQTTfhfXolQCIlIDFuL5V3CFADMzMzMzMzMzMxVi+yB7AgBAACNhXj///+5IAAAAFNWi3UMV429eP////OlUOgo7///jYV4////UOgc7///jYV4////UOgQ7///jb34/v//uwIAAAAPH0QAAIuNeP///4uFfP///4Hp7f8AAImN+P7//4PYAImF/P7//7gIAAAAZmYPH4QAAAAAAIt0B/iLTAf8i5QFeP///4l1+A+szhCLjAV8////g+YBx0QH/AAAAAAr1oPZAIHq//8AAImUBfj+//+D2QCJjAX8/v//D7dN+IlMB/iDwAiD+HhyrIuNaP///4uFbP///4tV8A+swRAPt4V" & _
                                                "o////g+EBiYVo////K9HHhWz///8AAAAAi030uAEAAACD2QCB6v9/AACJlXD///+D2QCJjXT///8PrMoQg+IBwfkQK8JQjYX4/v//UI2FeP///1DoTQAAAIPrAQ+FBP///4t1CDPSioTVeP///4uM1Xj///+IBFaLhNV8////D6zBCIhMVgFCwfgIg/oQctdfXluL5V3CCADMzMzMzMzMzMzMzMzMVYvsg+wIi0UQSPfQmVOLXQiJRfiLRQyJVfzzD35d+I1LeFYz9mYPbNuNUHg7wXdLO9NyRyvYx0UQEAAAAFdmkIs8GI1ACIt0GPyLSPiLUPwzzyNN+DPWI1X8M/kz8ol8GPiJdBj8MUj4MVD8g20QAXXOX15bi+VdwgwAi9ONSBAr0A8QDPONSSAPEFHQZg/v0WYP29MPKMJmD+/BDxEE84PGBA8QQdBmD+/QDxFR0A8QTArgDxBR4GYP79FmD9vTDyjCZg/vwQ8RRArgDxBB4GYP78IPEUHgg/4QcqVeW4vlXcIMAMzMzMzMzMzMzMzMVYvsi0UQU1aLdQiNSHhXi30MjVZ4O/F3BDvQcwuNT3g78XcwO9dyLCv4uxAAAAAr8IsUOCsQi0w4BBtIBI1ACIlUMPiJTDD8g+sBdeRfXltdwgwAi9eNSBCL3ivQK9gr/rgEAAAAjXYgjUkgDxBB0A8QTDfgZg/7yA8RTuAPEEwK4A8QQeBmD/vIDxFMC+CD6AF10l9eW13CDADMzMzMzFWL7ItNDFOLXQhWg8MQx0UMBAAAAFeDwQMPH4AAAAAAD7ZB/o1bIJmNSQiL8Iv6D7ZB9Q+k9wiZweYIA/CJc9AT+ol71A+2QfeZi/CL+g+2QfiZD6TCCMHgCAPwiXPYE/qJe9wPtkH6" & _
                                                "mYvwi/oPtkH5D6T3CJnB5ggD8Ilz4BP6iXvkD7ZB/JmL8Iv6D7ZB+w+k9wiZweYIA/CJc+gT+oNtDAGJe+wPhXT///+LTQhfXluBYXj/fwAAx0F8AAAAAF3CCADMzMzMzMzMzMzMzMxVi+yD7AhTi10MD1fAVleLfRCLE4vyi0MEi8hmDxNF+AM3E08EO/J1BjvIdQTrGDvIdw9yBDvycwm4AQAAADPS6wtmDxNF+ItF+ItV/It9CIlPBItNEIk3i3EIA3MIi0kME0sMA/ATyjtzCHUFO0sMdCA7Swx3EHIFO3MIcwm4AQAAADPS6wtmDxNF+ItV/ItF+IlPDItNEIl3CItxEANzEItJFBNLFAPwE8o7cxB1BTtLFHQgO0sUdxByBTtzEHMJuAEAAAAz0usLZg8TRfiLVfyLRfiJTxSJdxCLSxiLWxyJTQyLTRCLcRgDdQyLSRwTywPwE8o7dQx1BDvLdCw7y3cdcgU7dQxzFol3GLgBAAAAiU8cM9JfXluL5V3CDABmDxNF+ItV/ItF+Il3GIlPHF9eW4vlXcIMAMzMzMzMzFWL7ItNDLoDAAAAU4tdCFYr2Y1BGFeJXQgPH4AAAAAAizQDi1wDBIt4BIsIO993LnIiO/F3KDvfchp3BDvxchSLXQiD6AiD6gF51V9eM8BbXcIIAF9eg8j/W13CCABfXrgBAAAAW13CCADMzMzMzMxVi+yD7BBTi10QuUAAAABWi3UIK8tXi30MZg9uw4lNEIsHi1cEiUX4iVX88w9+TfhmD/PIZg/WDugz6P//i00QiUXwi0cIiVX0i1cMiUX4iVX88w9+TfhmD27DZg/zyPMPfkXwZg/ryGYP1k4I6P7n//+LTRCJRfCLRxCJVfSLVxSJRfiJV" & _
                                                "fzzD35N+GYPbsNmD/PI8w9+RfBmD+vIZg/WThDoyef//4tNEIlF8ItHGIlV9ItXHIlF+IlV/PMPfk34Zg9uw2YP88jzD35F8GYP68hmD9ZOGOiU5///X15bi+VdwgwAzMzMzMzMzMzMzMxVi+yD7ChTi10MD1fAVot1CFeLA2oBiQaLQwSJRgSLQwiJRgiLQwyJRgyLQxCJRhCLQxSJRhSLQxiJRhiLQxyJRhyLQyyJReSLQzCJReiLQzSJReyLQziJRfCLQzyJRfSNRdhQUGYPE0XYx0XgAAAAAOia/v//i/iNRdhQVlbo3fz//4tLMAP4i1M8M8ALQzSJReiNRdhqAVCJTeSLSzhQx0XgAAAAAIlN7IlV8MdF9AAAAADoV/7//wP4jUXYUFZW6Jr8//8D+MdF5AAAAACLQyAPV8CJRdiLQySJRdyLQyiJReCLQziJRfCLQzxmDxNF6IlF9I1F2FBWVuhg/P//A/iLSyQzwItTNAtDKIlF3ItDMIlF+DPAC0MsiUXgi0M4iUXoi0M8iUXsM8ALQyCJRfSNRdhQiU3Yi8pWVolN5IlV8OgY/P//i1M0A/iLSywzwAtDMA9XwIlF3DPAC0MgiUXwjUXYUFaJfQiLeyhWiU3YiVXgx0XkAAAAAGYPE0XoiX306FcIAAApRQgPV8CLQzCxIItVDDP/iUXYi0M0iUXci0M4iUXgi0M8i1ssiUXki0Igi1IkZg8TRejov+X//wv4jUXYUAvaiX3wVold9FboCggAAItdDClFCDPAC0M4i0s0i1Msi3swiUXcM8ALQzyJTdiLSyCJReCLQyiJTeSxIOhY5f//C0MkiUXojUXYUFZWiVXsx0XwAAAAAIl99Oi6BwAAi30IK/jHReAAAAAAi0" & _
                                                "M4iUXYi0M8iUXci0MkiUXki0MoiUXoi0MsiUXsi0M0iUX0jUXYUFZWx0XwAAAAAOh4BwAAK/h5JA8fQADoi9j//1BWVujj+v//A/h4719eW4vlXcIIAGYPH0QAAIX/dRFW6GbY//9Q6AD8//+D+AF03OhW2P//UFZW6C4HAAAr+OvazMzMzMzMzMzMzFWL7ItVDIHsiAAAADPJZpCLBMoLRMoEdUZBg/kEcvGLRQjHAAAAAADHQAQAAAAAx0AIAAAAAMdADAAAAADHQBAAAAAAx0AUAAAAAMdAGAAAAADHQBwAAAAAi+VdwgwAiwIPV8CJhXj///+LQgSJhXz///+LQgiJRYCLQgyJRYSLQhCJRYiLQhSJRYyLQhiJRZCLQhyJRZRWi3UQVzP/Zg8TReBmDxNF6IsGiUWYi0YEiUWci0YIiUWgi0YMiUWki0YQiUWoi0YUiUWsi0YYiUWwi0YciUW0jUWYUI2FeP///2YPE0XwUMdF2AEAAACJfdxmDxNFuGYPE0XAZg8TRchmDxNF0OjW+v//i9CF0g+EugEAAFNmZmYPH4QAAAAAAIuNeP///w9XwIPhAWYPE0X4g8kAdS+NhXj///9Q6L4DAACLRdiD4AGDyAAPhLYAAABWjUXYUFDoRPn//4v4i9rpqAAAAItFmIPgAYPIAHUsjUWYUOiHAwAAi0W4g+ABg8gAD4QIAQAAVo1FuFBQ6A35//+L+Iva6foAAACF0g+OjAAAAI1FmFCNhXj///9QUOhrBQAAjYV4////UOg/AwAAjUW4UI1F2FDoEvr//4XAeQtWjUXYUFDow/j//41FuFCNRdhQUOg1BQAAi0XYg+ABg8gAdBFWjUXYUFDon/j//4v4i9rrBotd/It9+I1F2FD" & _
                                                "o6gIAAAv7D4SSAAAAi0XwgU30AAAAgIlF8OmAAAAAjYV4////UI1FmFBQ6N8EAACNRZhQ6LYCAACNRdhQjUW4UOiJ+f//hcB5C1aNRbhQUOg6+P//jUXYUI1FuFBQ6KwEAACLRbiD4AGDyAB0EVaNRbhQUOgW+P//i/iL2usGi138i334jUW4UOhhAgAAC/t0DYtF0IFN1AAAAICJRdCNRZhQjYV4////UOgg+f//i9CF0g+FVv7//4t93FuLTQiLRdiJAYtF4IlBCItF5IlBDItF6IlBEItF7IlBFItF8Il5BIlBGItF9F+JQRxei+VdwgwAzMzMzMzMzMzMzMxVi+yD7FRTD1fAM8lmDxNF1ItF2FZmDxNFzItd0IlF+ItF1FeLfcyJTeyJRfwPHwAz9o1B/YP5BA9XwGYPE0Xwi1X0D0PwO/EPhy8BAACLwYlV6ItNECvGjQTBi03wiU30i03siUXkg/4ED4PYAAAA/3AE/zCLRQz/dPAE/zTwjUWsUOh48v//DxAAZg9+wQ8RRcxmD3PYBAPPZg9+wIlNvBPDiUXAO8N3D3IEO89zCbkBAAAAM9LrDg9XwGYPE0Xci1Xgi03ci33UA8+LRdgT0ANN/IlNxBNV+IlVyA8QRbwPEUXMO9B3D3IEO89zCbgBAAAAM8nrDg9XwGYPE0Xci03gi0XcAUX0i1Xoi0XkE9GLTexGi33Mg+gIiVXoiUXkO/F3FItd2Ild+Itd1Ild/Itd0Oku////i0XYi13QiUX4i0XUiUX8i0X0i3UIiTzOi338iVzOBEGLXfiJRfyJVfiJTeyD+QcPgsL+//+JfjhfiV48XluL5V3CDACLRfDryczMzMzMzMzMzMzMzMzMVYvsi1UIuAMAAAAPH0QA" & _
                                                "AIsMwgtMwgR1BYPoAXnyjUgBhcl1BjPAXcIEAFaLdMr4i1TK/IvOVzP/C8p0EA8fAA+s1gFH0eqLzgvKdfPB4AYDx19eXcIEAMzMzMzMzMzMVYvsg+wIi0UID1fAU4vYZg8TRfiDwCA7w3Y4i034VleLffyJTQiLcPiD6AiLzotQBA+s0QELTQjR6gvXiQiL/olQBMHnH8dFCAAAAAA7w3fVX15bi+VdwgQAzMzMzMzMVYvsg+xYD1fAM8lTZg8TRdCLRdRmDxNFyItVzItdyFaJRfCLRdBXiU3YiUX0iVXsM/aNQf2D+QQPV8BmDxNF4A9D8DvxD4drAQAAi30Mi8Erxo0Ex4t95Il9/It94IlF3Il9+Iv5K/479w+HCwEAAP9wBP8wi0UM/3TwBP808I1FqFDoKPD//w8QAA8RRciLVcw793Mxi03Ui/qLwcHoHwFF+ItF0INV/AAPpMEBwe8fA8AL+DPAC8GJReSLRcgPpMIBA8DrDItF1It90IlF5ItFyAPDiUW4E1XsiVW8O1Xsdw9yBDvDcwm4AQAAADPJ6w4PV8BmDxNF6ItN7ItF6APHE03kA0X0iUXAE03wiU3EDxBFuA8RRcg7TeR3D3IEO8dzCbgBAAAAM8nrDg9XwGYPE0Xgi03ki0XgAUX4i0XcEU38RotN2IPoCItdyIlF3DvxdxeLVdSJVfCLVdCJVfSLVcyJVezp+P7//4tF1ItVzIlF8ItF0IlF9ItF+It9/It1CIkczotd9IlUzgRBi1XwiVXsiUX0iX3wiU3Yg/kHD4KJ/v//X4leOIlWPF5bi+VdwggAi33ki0Xg68PMzFWL7IPsDFOLXQwPV8BWV4t9EIsTi/KLQwSLyGYPE0X0KzcbTwQ78nUGO8h1B" & _
                                                "OsYO8hyD3cEO/J2CbgBAAAAM9LrC2YPE0X0i0X0i1X4i30IiU8Ei00QiTeLcwiJdfgrcQiLSwyLXRAbSwwr8ItdDBvKO3X4dQU7Swx0IDtLDHIQdwU7cwh2CbgBAAAAM9LrC2YPE0X0i1X4i0X0iU8Mi00QiXcIi3MQiXX8K3EQi0sUi10QG0sUK/CLXQwbyjt1/HUFO0sUdCA7SxRyEHcFO3MQdgm4AQAAADPS6wtmDxNF9ItV+ItF9IlPFIl3EItLGIvxi30Qi1sciU0Mi00QK3EYi8sbTxwr8It9CBvKO3UMdQQ7y3QsO8tyHXcFO3UMdhaJdxi4AQAAAIlPHDPSX15bi+VdwgwAZg8TRfSLVfiLRfSJdxiJTxxfXluL5V3CDAAAAA==" ' 15.4.2020 13:27:58

Private m_uEcc                  As UcsEccThunkData

Private Enum UcsEccPfnIndexEnum
    ucsPfnSecp256r1MakeKey
    ucsPfnSecp256r1SharedSecret
    ucsPfnCurve25519ScalarMultiply
    ucsPfnCurve25519ScalarMultBase
    [_ucsPfnMax]
End Enum

Private Type UcsEccThunkData
    Thunk               As Long
    Glob                As Long
    Pfn(0 To [_ucsPfnMax] - 1) As Long
    KeySize             As Long
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
        .KeySize = 32
        '--- prepare thunk/context in executable memory
        .Thunk = pvThunkAllocate(STR_ECC_THUNK1 & STR_ECC_THUNK2)
        .Glob = pvThunkAllocate(STR_ECC_CTX)
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
    
    ReDim baPrivate(0 To m_uEcc.KeySize - 1) As Byte
    ReDim baPublic(0 To m_uEcc.KeySize) As Byte
    For lIdx = 1 To MAX_RETRIES
        pvEccRandomBytes VarPtr(baPrivate(0)), m_uEcc.KeySize
        If pvEccCallSecp256r1MakeKey(m_uEcc.Pfn(ucsPfnSecp256r1MakeKey), VarPtr(baPublic(0)), VarPtr(baPrivate(0))) = 1 Then
            Exit For
        End If
    Next
    '--- success (or failure)
    EccSecp256r1MakeKey = (lIdx <= MAX_RETRIES)
End Function

Public Function EccSecp256r1SharedSecret(baPrivate() As Byte, baPublic() As Byte) As Byte()
    Dim baRetVal()      As Byte
    
    Debug.Assert UBound(baPrivate) >= m_uEcc.KeySize - 1
    Debug.Assert UBound(baPublic) >= m_uEcc.KeySize
    ReDim baRetVal(0 To m_uEcc.KeySize - 1) As Byte
    If pvEccCallSecp256r1SharedSecret(m_uEcc.Pfn(ucsPfnSecp256r1SharedSecret), VarPtr(baPublic(0)), VarPtr(baPrivate(0)), VarPtr(baRetVal(0))) = 0 Then
        GoTo QH
    End If
    EccSecp256r1SharedSecret = baRetVal
QH:
End Function

Public Function EccCurve25519MakeKey(baPrivate() As Byte, baPublic() As Byte) As Boolean
    ReDim baPrivate(0 To m_uEcc.KeySize - 1) As Byte
    ReDim baPublic(0 To m_uEcc.KeySize - 1) As Byte
    pvEccRandomBytes VarPtr(baPrivate(0)), m_uEcc.KeySize
    baPrivate(0) = baPrivate(0) And 248
    baPrivate(UBound(baPrivate)) = (baPrivate(UBound(baPrivate)) And 127) Or 64
    pvEccCallCurve25519MulBase m_uEcc.Pfn(ucsPfnCurve25519ScalarMultBase), VarPtr(baPublic(0)), VarPtr(baPrivate(0))
    '--- success
    EccCurve25519MakeKey = True
End Function

Public Function EccCurve25519SharedSecret(baPrivate() As Byte, baPublic() As Byte) As Byte()
    Dim baRetVal()      As Byte
    
    Debug.Assert UBound(baPrivate) >= m_uEcc.KeySize - 1
    Debug.Assert UBound(baPublic) >= m_uEcc.KeySize - 1
    ReDim baRetVal(0 To m_uEcc.KeySize - 1) As Byte
    pvEccCallSecp256r1SharedSecret m_uEcc.Pfn(ucsPfnCurve25519ScalarMultiply), VarPtr(baRetVal(0)), VarPtr(baPrivate(0)), VarPtr(baPublic(0))
    EccCurve25519SharedSecret = baRetVal
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
