Attribute VB_Name = "ModWelcome"
' function returns if email and password are ok (connects to supabase and gets basic data)
Function login(email As String, password As String)
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    url = "https://aeawfxrqshwuxazdckoi.supabase.co/auth/v1/token?grant_type=password"
    xmlhttp.Open "POST", url, False
    xmlhttp.setRequestHeader "apikey", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImFlYXdmeHJxc2h3dXhhemRja29pIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTY3OTQxMzI3MywiZXhwIjoxOTk0OTg5MjczfQ.wLyIVFYOFkIbaIJVG9r1iH2FCdM8wem3ClDnhuwfOIQ"
    xmlhttp.setRequestHeader "Content-Type", "application/json"
    xmlhttp.Send "{" & Chr(34) & "email" & Chr(34) & ":" & Chr(34) & email & Chr(34) & "," & Chr(34) & "password" & Chr(34) & ":" & Chr(34) & password & Chr(34) & "}"
    If InStr(xmlhttp.responseText, "error") Then
        'MsgBox "error"
        Welcome.lblWelcomeLoginMessage.Caption = "Error" & vbNewLine & "Wrong credentials. Please try again."
        Welcome.lblWelcomeLoginMessage.Visible = True
        Exit Function
        login = "loginnotok"
    Else
        Set json = JsonConverter.ParseJson(xmlhttp.responseText)
        uid = json("user")("id")
        getUserDetails (uid)
        Welcome.lblWelcomeLoginMessage.Caption = "Success" & vbNewLine & "Logged in ! Redirecting to the app ..."
        Welcome.lblWelcomeLoginMessage.BorderColor = &HFF00&
        Welcome.lblWelcomeLoginMessage.ForeColor = &HFF00&
        Welcome.lblWelcomeLoginMessage.Visible = True
        Application.Wait (Now + TimeValue("00:00:02"))
        DoEvents
        Welcome.Hide
        UserForm1.Show
        login = "loginok"
    End If
End Function

' sub to sign up user (requires email, names, passwords)
Sub signup(email As String, first_name As String, last_name As String, password As String, confirm As String)
    If Not (IsValidEmail(email)) Then
        Welcome.lblWelcomeSignupMessage.Caption = "Error" & vbNewLine & "Invalid email."
        Welcome.lblWelcomeSignupMessage.Visible = True
    ElseIf Len(first_name) < 2 Then
        Welcome.lblWelcomeSignupMessage.Caption = "Error" & vbNewLine & "First name too short."
        Welcome.lblWelcomeSignupMessage.Visible = True
    ElseIf Len(last_name) < 2 Then
        Welcome.lblWelcomeSignupMessage.Caption = "Error" & vbNewLine & "Last name too short."
        Welcome.lblWelcomeSignupMessage.Visible = True
    ElseIf Len(password) < 2 Then
        Welcome.lblWelcomeSignupMessage.Caption = "Error" & vbNewLine & "Password too short."
        Welcome.lblWelcomeSignupMessage.Visible = True
    ElseIf password <> confirm Then
        Welcome.lblWelcomeSignupMessage.Caption = "Error" & vbNewLine & "Passwords do not match."
        Welcome.lblWelcomeSignupMessage.Visible = True
    Else
        Dim xmlhttp As Object
        Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        Dim json As Object
        url = "https://aeawfxrqshwuxazdckoi.supabase.co/auth/v1/signup"
        xmlhttp.Open "POST", url, False
        xmlhttp.setRequestHeader "apikey", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImFlYXdmeHJxc2h3dXhhemRja29pIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTY3OTQxMzI3MywiZXhwIjoxOTk0OTg5MjczfQ.wLyIVFYOFkIbaIJVG9r1iH2FCdM8wem3ClDnhuwfOIQ"
        xmlhttp.setRequestHeader "Content-Type", "application/json"
        xmlhttp.Send "{" & Chr(34) & "email" & Chr(34) & ":" & Chr(34) & email & Chr(34) & "," & Chr(34) & "password" & Chr(34) & ":" & Chr(34) & password & Chr(34) & "}"
        If InStr(xmlhttp.responseText, "phone") Then
            Set json = JsonConverter.ParseJson(xmlhttp.responseText)
            Welcome.lblWelcomeSignupMessage.Caption = "Success" & vbNewLine & "Please confirm your e-mail."
            Welcome.lblWelcomeSignupMessage.BorderColor = &HFF00&
            Welcome.lblWelcomeSignupMessage.ForeColor = &HFF00&
            Welcome.lblWelcomeSignupMessage.Visible = True
            Call registerToSupabase(json("id"), first_name, last_name)
        Else
            Welcome.lblWelcomeSignupMessage.Caption = "Error" & vbNewLine & "Could not register. Please try again"
            Welcome.lblWelcomeSignupMessage.Visible = True
        End If
    End If
End Sub

' sub to add to supabase database names and uuid from uuid from auth api
Sub registerToSupabase(id As String, fname As String, lname As String)
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    url = "https://aeawfxrqshwuxazdckoi.supabase.co/rest/v1/users"
    xmlhttp.Open "POST", url, False
    xmlhttp.setRequestHeader "apikey", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImFlYXdmeHJxc2h3dXhhemRja29pIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTY3OTQxMzI3MywiZXhwIjoxOTk0OTg5MjczfQ.wLyIVFYOFkIbaIJVG9r1iH2FCdM8wem3ClDnhuwfOIQ"
    xmlhttp.setRequestHeader "Authorization", "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImFlYXdmeHJxc2h3dXhhemRja29pIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTY3OTQxMzI3MywiZXhwIjoxOTk0OTg5MjczfQ.wLyIVFYOFkIbaIJVG9r1iH2FCdM8wem3ClDnhuwfOIQ"
    xmlhttp.setRequestHeader "Content-Type", "application/json"
    xmlhttp.setRequestHeader "Prefer", "return=minimal"
    xmlhttp.Send "{" & Chr(34) & "id" & Chr(34) & ":" & Chr(34) & id & Chr(34) & "," & Chr(34) & "first_name" & Chr(34) & ":" & Chr(34) & fname & Chr(34) & "," & Chr(34) & "last_name" & Chr(34) & ":" & Chr(34) & lname & Chr(34) & "}"
End Sub

' function to check if email field is valid
' inspired from internet
Function IsValidEmail(sEmailAddress As String) As Boolean
    'Code from Officetricks
    'Define variables
    Dim sEmailPattern As String
    Dim oRegEx As Object
    Dim bReturn As Boolean
    
    'Use the below regular expressions
    sEmailPattern = "^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$" 'or
    sEmailPattern = "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"
    
    'Create Regular Expression Object
    Set oRegEx = CreateObject("VBScript.RegExp")
    oRegEx.Global = True
    oRegEx.IgnoreCase = True
    oRegEx.Pattern = sEmailPattern
    bReturn = False
    
    'Check if Email match regex pattern
    If oRegEx.test(sEmailAddress) Then
        'Debug.Print "Valid Email ('" & sEmailAddress & "')"
        bReturn = True
    Else
        'Debug.Print "Invalid Email('" & sEmailAddress & "')"
        bReturn = False
    End If

    'Return validation result
    IsValidEmail = bReturn
End Function

' sub to diplay in Account section names and creation date of user
Sub getUserDetails(uid As String)
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    url = "https://aeawfxrqshwuxazdckoi.supabase.co/rest/v1/users?id=eq." & uid & "&select=*"
    xmlhttp.Open "GET", url, False
    xmlhttp.setRequestHeader "apikey", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImFlYXdmeHJxc2h3dXhhemRja29pIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTY3OTQxMzI3MywiZXhwIjoxOTk0OTg5MjczfQ.wLyIVFYOFkIbaIJVG9r1iH2FCdM8wem3ClDnhuwfOIQ"
    xmlhttp.setRequestHeader "Authorization", "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImFlYXdmeHJxc2h3dXhhemRja29pIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTY3OTQxMzI3MywiZXhwIjoxOTk0OTg5MjczfQ.wLyIVFYOFkIbaIJVG9r1iH2FCdM8wem3ClDnhuwfOIQ"
    xmlhttp.Send
    'MsgBox (xmlhttp.responseText)
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    'MsgBox json(1)("first_name")
    UserForm1.lblAccountFname.Caption = json(1)("first_name")
    UserForm1.lblAccountLname.Caption = json(1)("last_name")
    UserForm1.lblAccountDate.Caption = json(1)("created_at")
End Sub
