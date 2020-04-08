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
   Begin VB.TextBox txtResult 
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
      Default         =   -1  'True
      Height          =   348
      Left            =   3864
      TabIndex        =   2
      Top             =   168
      Width           =   1356
   End
   Begin VB.TextBox txtUrl 
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

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
'--- libsodium
Private Declare Function sodium_init Lib "libsodium" () As Long

Private Type UcsParsedUrl
    Protocol        As String
    Host            As String
    Port            As Long
    Path            As String
    User            As String
    Pass            As String
End Type

'=========================================================================
' Events
'=========================================================================

Private Sub Command1_Click()
    Dim uRemote         As UcsParsedUrl
    Dim uCtx            As UcsClientContextType
    Dim sResult         As String
    Dim sError          As String
    
    ' tls13.1d.pw, localhost:44330
    If Not ParseUrl(txtUrl.Text, uRemote, DefProtocol:="https") Then
        MsgBox "Wrong URL", vbCritical
        GoTo QH
    End If
    uCtx = TlsInitClient(uRemote.Host & ":" & uRemote.Port, txtResult)
    sResult = TlsFetchHttp(uCtx, uRemote.Path, sError)
    If LenB(sResult) <> 0 Then
        txtResult.Text = sResult
    End If
    If LenB(sError) <> 0 Then
        MsgBox sError, vbCritical
        GoTo QH
    End If
QH:
End Sub

Private Sub Form_Load()
    If GetModuleHandle("libsodium.dll") = 0 Then
        Call LoadLibrary(App.Path & "\libsodium.dll")
        Call sodium_init
    End If
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        txtResult.Move 0, txtResult.Top, ScaleWidth, ScaleHeight - txtResult.Top
    End If
End Sub

Private Function ParseUrl(sUrl As String, uParsed As UcsParsedUrl, Optional DefProtocol As String) As Boolean
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = "^(?:(.*)://)?(?:(?:([^:]*):)?([^@]*)@)?([A-Za-z0-9\-\.]+)(:[0-9]+)?(.*)$"
        With .Execute(sUrl)
            If .Count > 0 Then
                With .Item(0).SubMatches
                    uParsed.Protocol = .Item(0)
                    uParsed.User = .Item(1)
                    If LenB(uParsed.User) = 0 Then
                        uParsed.User = .Item(2)
                    Else
                        uParsed.Pass = .Item(2)
                    End If
                    uParsed.Host = .Item(3)
                    uParsed.Port = Val(Mid$(.Item(4), 2))
                    If uParsed.Port = 0 Then
                        Select Case LCase$(IIf(LenB(uParsed.Protocol) = 0, DefProtocol, uParsed.Protocol))
                        Case "https"
                            uParsed.Port = 443
                        Case "socks5"
                            uParsed.Port = 1080
                        Case Else
                            uParsed.Port = 80
                        End Select
                    End If
                    uParsed.Path = .Item(5)
                    If LenB(uParsed.Path) = 0 Then
                        uParsed.Path = "/"
                    End If
                End With
                ParseUrl = True
            End If
        End With
    End With
End Function
