VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Welcome 
   Caption         =   "Excelliarmus"
   ClientHeight    =   9735.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12240
   OleObjectBlob   =   "Welcome.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' sub to check if credentials are OK
Private Sub btnWelcomeLogin_Click()
    Dim resp As String
    resp = ModWelcome.login(Welcome.inputWelcomeEmail, Welcome.inputWelcomePassword)
    If resp = "loginok" Then
        Unload Me
    End If
End Sub

' sub to sign up user (requires e-mail, names and passwords)
Private Sub btnWelcomeSignUp_Click()
    Call ModWelcome.signup(Welcome.inputSignupEmail, Welcome.inputSignupFname, Welcome.inputSignupLname, Welcome.inputSignupPassword, Welcome.inputSignupConfirm)
End Sub

' sub to display the Login frame
Private Sub lblWelcomeLogin_Click()
    Welcome.frmWelcomeLogin.Visible = True
    Welcome.frmWelcomeSignup.Visible = False
End Sub

' sub to display the Signup frame
Private Sub lblWelcomeSignUp_Click()
    Welcome.frmWelcomeLogin.Visible = False
    Welcome.frmWelcomeSignup.Visible = True
End Sub
