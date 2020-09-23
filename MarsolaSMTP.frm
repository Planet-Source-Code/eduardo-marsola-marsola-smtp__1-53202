VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMarsolaSMTP 
   Caption         =   "Marsola SMTP"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCorpo 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "MarsolaSMTP.frx":0000
      Top             =   4320
      Width           =   8055
   End
   Begin VB.TextBox txtAssunto 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Text            =   "Assunto"
      Top             =   3660
      Width           =   8055
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   180
      TabIndex        =   15
      Top             =   5760
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdAnexar 
      Caption         =   "Attachment"
      Height          =   435
      Left            =   2520
      TabIndex        =   7
      Top             =   2760
      Width           =   1755
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   2820
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   2520
      TabIndex        =   14
      Top             =   300
      Width           =   5655
   End
   Begin VB.TextBox txtNomeDe 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Text            =   "Eduardo"
      Top             =   2820
      Width           =   2115
   End
   Begin VB.TextBox txtEmailDe 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "teste@test.com"
      Top             =   2220
      Width           =   2115
   End
   Begin VB.TextBox txtNomePara 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Text            =   "Eduardo Marsola"
      Top             =   1560
      Width           =   2115
   End
   Begin VB.TextBox txtEmailPara 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "teste@test.com"
      Top             =   960
      Width           =   2115
   End
   Begin VB.TextBox txtSMTPServer 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "10.152.229.106"
      Top             =   360
      Width           =   2115
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Send"
      Height          =   495
      Left            =   6900
      TabIndex        =   8
      Top             =   5640
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5220
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label7 
      Caption         =   "Body:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   4080
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "Subject:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3420
      Width           =   2115
   End
   Begin VB.Label Label6 
      Caption         =   "From (nome) - Optional:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2580
      Width           =   2115
   End
   Begin VB.Label Label5 
      Caption         =   "From (e-mail):"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1980
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "To (nome) - Optional:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   2115
   End
   Begin VB.Label Label3 
      Caption         =   "To (e-mail):"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "SMTP Server:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMarsolaSMTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Code created by Eduardo Marsola - edu@marsola.com
'For sending e-mails with attachment useing SMTP server and WINSOCK
'Base64 Encode, Base64 Decode, Mime Version 1.0, etc
'I know there is problem with performance.. it is signed
'If you need more information, just send an e-mail!
'
'
'*********************Enjoy it!!***********************
'Eduardo Marsola
'
'

Option Explicit
Dim Response As String, Reply As Integer
Dim Start As Single, Tmr As Single
Dim sS As String


Sub SendEmail(MailServerName As String, FromName As String, FromEmailAddress As String, ToName As String, ToEmailAddress As String, EmailSubject As String, EmailBodyOfMessage As String, sArquivo() As String)
Dim sComplemento As String
    Winsock1.LocalPort = 0 ' Must Set local port To 0 (Zero) or you can only send 1 e-mail per program start
    
    'create the message complemente (after DATA)
    sComplemento = FU_Complemento(FromName, FromEmailAddress, ToName, ToEmailAddress, EmailSubject, EmailBodyOfMessage, sArquivo)
    
    If Winsock1.State = sckClosed Then ' Check To see if socet is closed
        Winsock1.Protocol = sckTCPProtocol ' Set protocol For sending
        Winsock1.RemoteHost = MailServerName ' Set the server address
        Winsock1.RemotePort = 25 ' Set the SMTP Port
        Winsock1.Connect ' Start connection
        WaitFor ("220")
        
        Call FU_Envia("HELO " & Winsock1.LocalHostName & vbCrLf)
        WaitFor ("250")
        
        Call FU_Envia("mail from:" + Chr$(32) + FromEmailAddress + vbCrLf)  ' Get who's sending E-Mail address
        WaitFor ("250")
        
        Call FU_Envia("rcpt to:" + Chr$(32) + ToEmailAddress + vbCrLf)  ' Get who mail is going to
        WaitFor ("250")
        
        Call FU_Envia("data" + vbCrLf)
        WaitFor ("354")
        
        
        Call FU_Envia(sComplemento)
        Call FU_Envia(vbCrLf & "." & vbCrLf)
        WaitFor ("250")
        Call FU_Envia("quit" + vbCrLf)
        
        WaitFor ("221")
        Winsock1.Close
    Else
        MsgBox (Str(Winsock1.State))
    End If
    
End Sub


Sub WaitFor(ResponseCode As String)
    Start = Timer ' Time Event so won't Get stuck In Loop

    While Len(Response) = 0
        Tmr = -(Start - Timer)
        DoEvents ' Let System keep checking For incoming response **IMPORTANT**
            If Tmr > 50 Then ' Time In seconds To wait
                MsgBox "SMTP service error, timed out While waiting For response", 64
                Exit Sub
            End If
        Wend


        While Left(Response, 3) <> ResponseCode


            DoEvents


                If Tmr > 50 Then
                    MsgBox "SMTP service error, impromper response code. Code should have been: " + ResponseCode + " Code recieved: " + Response, 64
                    Exit Sub
                End If
            Wend
            Response = "" ' Sent response code To blank **IMPORTANT**
        End Sub






Private Sub cmdAnexar_Click()
CommonDialog1.Filter = "*.*"
CommonDialog1.ShowOpen

If Len(Trim(CommonDialog1.filename)) > 0 Then
    List1.AddItem CommonDialog1.filename
End If
End Sub


Private Sub cmdEnviar_Click()
Dim sArquivo() As String
Dim iContador1 As Integer
For iContador1 = 0 To List1.ListCount - 1
    ReDim Preserve sArquivo(iContador1)
    sArquivo(iContador1) = List1.List(iContador1)
Next
If Len(Dir("C:\EDU.LOG")) > 0 Then Kill "C:\EDU.LOG"
Call SendEmail(txtSMTPServer, txtNomeDe, txtEmailDe, txtNomePara, txtEmailPara, txtAssunto, txtCorpo, sArquivo())


End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Winsock1.GetData Response, vbString ' Check For incoming response *IMPORTANT*
    GravaLog Response
    
End Sub

Function FU_Envia(sTexto As String)
    GravaLog (sTexto)
    Debug.Print sTexto
    Winsock1.SendData sTexto
    
End Function

Public Sub GravaLog(ByVal sTexto As String)
    Open "C:\EDU.LOG" For Append As #2
    Print #2, Format(Now, "DD/MM/YYYY-HH:MM:SS - ") & sTexto
    Close #2
    
    
End Sub

Function FU_Complemento(FromName As String, _
    FromEmailAddress As String, _
    ToName As String, _
    ToEmailAddress As String, _
    EmailSubject As String, _
    EmailBodyOfMessage As String, _
    sArquivo() As String) As String

Dim sTexto As String
Dim iQuantArquivos As Integer
Dim lTot As Long
sS = ""
sS = sS & "From: " & Chr$(34) & FromName & Chr$(34) & " <" & FromEmailAddress & ">" & vbCrLf
sS = sS & "To: " & Chr$(34) & ToName & Chr$(34) & " <" & ToEmailAddress & ">" & vbCrLf
sS = sS & "Subject: " & EmailSubject & vbCrLf
sS = sS & "Date: " & Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & " -0300" & vbCrLf
sS = sS & "MIME-Version: 1.0" & vbCrLf
sS = sS & "Content-Type: multipart/mixed;" & vbCrLf
sS = sS & vbTab & "boundary=" & Chr$(34) & "----=_NextPart_000_000C_01C21C67F3F2CCA0" & Chr$(34) & vbCrLf
sS = sS & vbCrLf
sS = sS & "This is a multi-part message in MIME format." & vbCrLf
sS = sS & vbCrLf
sS = sS & "------=_NextPart_000_000C_01C21C67F3F2CCA0" & vbCrLf
sS = sS & "Content-Type: text/plain;" & vbCrLf
sS = sS & vbTab & "charset=" & Chr$(34) & "iso-8859-1" & Chr$(34) & vbCrLf
sS = sS & "Content-Transfer-Encoding: 7bit" & vbCrLf
sS = sS & vbCrLf
sS = sS & EmailBodyOfMessage & vbCrLf
sS = sS & vbCrLf

For iQuantArquivos = 0 To UBound(sArquivo)


    sS = sS & "------=_NextPart_000_000C_01C21C67F3F2CCA0" & vbCrLf
    sS = sS & "Content-Type: application/octet-stream;" & vbCrLf
    sS = sS & vbTab & "name=" & Chr$(34) & Dir(sArquivo(iQuantArquivos)) & Chr$(34) & vbCrLf
    sS = sS & "Content-Transfer-Encoding: base64" & vbCrLf
    sS = sS & "Content-Disposition: attachment;" & vbCrLf
    sS = sS & vbTab & "filename=" & Chr$(34) & Dir(sArquivo(iQuantArquivos)) & Chr$(34) & vbCrLf
    sS = sS & vbCrLf
    sTexto = Base64Enc01(sArquivo(iQuantArquivos))
    Dim y As Long
    lTot = Len(sTexto) / 74
    ProgressBar1.Min = 1
    ProgressBar1.Max = lTot + 1
    
    'There is problem with performance on thes FOR below!!!
    'I would love a solution, but I couldn't spend more time looking for!!!
    'Eduardo Marsola
    For y = 0 To lTot
        DoEvents
        sS = sS & Mid(sTexto, (y * 74) + 1, 74) & vbCrLf
        ProgressBar1.Value = y + 1
    Next
    

Next
    sS = sS & vbCrLf
    sS = sS & "------=_NextPart_000_000C_01C21C67F3F2CCA0--" & vbCrLf
FU_Complemento = sS
End Function



 Function Base64Enc01(sArquivo) As String
' by Nobody, 20011204
  Static Enc() As Byte
  Dim nLen As Integer

  Dim b() As Byte, Out() As Byte, i&, j&, L&
  If (Not Val(Not Enc)) = 0 Then 'Null-Ptr = not initialized
    Enc = StrConv("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", vbFromUnicode)
  End If
    Dim iFile As Integer
    Dim s As String
    Dim TheBytes() As Byte
    ReDim TheBytes(FileLen(sArquivo))
    iFile = FreeFile
    Open sArquivo For Binary Access Read Lock Read Write As #iFile 'Input As #iFile 'Binary Access Read Lock Read Write As #iFile

    s = Input(LOF(iFile), #iFile)
    Close #iFile
  
  L = Len(s): b = StrConv(s, vbFromUnicode)
  ReDim Preserve b(0 To (UBound(b) \ 3) * 3 + 2)
  ReDim Preserve Out(0 To (UBound(b) \ 3) * 4 + 3)
  For i = 0 To UBound(b) - 1 Step 3
    Out(j) = Enc(b(i) \ 4): j = j + 1
    Out(j) = Enc((b(i + 1) \ 16) Or (b(i) And 3) * 16): j = j + 1
    Out(j) = Enc((b(i + 2) \ 64) Or (b(i + 1) And 15) * 4): j = j + 1
    Out(j) = Enc(b(i + 2) And 63): j = j + 1
    DoEvents
  Next i
  For i = 1 To i - L: Out(UBound(Out) - i + 1) = 61: Next i
  Base64Enc01 = StrConv(Out, vbUnicode)
End Function


