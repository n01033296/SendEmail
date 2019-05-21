VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   2145
   ClientTop       =   2085
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.Frame Send 
      Caption         =   "SendEmailTest"
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   3855
      Begin VB.CommandButton SendEmail 
         Caption         =   "Send"
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Command1_Click()
        
End Sub

'NOTE: INCLUDE EASendMailObj Reference
Private Sub SendEmail_Click()
    Dim oSmtp As New EASendMailObjLib.mail
    oSmtp.LicenseCode = "TryIt"
    
    'Set from Address
    oSmtp.FromAddr = "sameplemail@gmail.com"
    
    'Add Recipients
    oSmtp.AddRecipientEx "samplerecipient@gmail.com", 0
    
    'Add Subject
    oSmtp.Subject = "Sample Mail from VB 6.0"
    
    'Add Body
    oSmtp.BodyText = "This is the test email Body, Do not reply"
    
    'Server Address and Port
    oSmtp.ServerAddr = "smtp.gmail.com"
    oSmtp.ServerPort = 587
    'Detects SSL/TLS Automatically
    oSmtp.SSL_init
    
    'Authentication for User and Password
    oSmtp.UserName = "gmailsample.gmail.com"
    oSmtp.Password = "Password"
    
    'Check if Email has been sent
    MsgBox "started to Send Email..."
    If oSmtp.SendMail() = 0 Then
        MsgBox "email was sent Sucessfully!"
    Else
        MsgBox "Failed to send Email with the Following Error:" & oSmtp.GetLastErrDescription()
    End If
    
End Sub



