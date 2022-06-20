set args = wscript.Arguments

WScript.ECHO EMail ( args(0), _
                    args(1), _
                    args(2), _
                    args(3), _
                    args(4), _
                    args(5), _
                    args(6), _
                    args(7) )

Function EMail( myFrom, myTo, mySubject, myTextBody, myHTMLBody, myAttachment, mySMTPServer, mySMTPPort )
' This function sends an e-mail message using CDOSYS
'
' Arguments:
' myFrom       = Sender's e-mail address ("John Doe jdoe@mydomain.org" or jdoe@mydomain.org)
' myTo         = Receiver's e-mail address ("John Doe jdoe@mydomain.org" or jdoe@mydomain.org)
' mySubject    = Message subject (optional)
' myTextBody   = Actual message (text only, optional)
' myHTMLBody   = Actual message (HTML, optional)
' myAttachment = Attachment as fully qualified file name, either string or array of strings (optional)
' mySMTPServer = SMTP server (IP address or host name)
' mySMTPPort   = SMTP server port (optional, default 25)
'
' Returns:
' status message
'
' Written by Rob van der Woude
' http://www.robvanderwoude.com

    ' Standard housekeeping
    Dim i, objEmail

    ' Use custom error handling
    On Error Resume Next

    ' Create an e-mail message object
    Set objEmail = CreateObject( "CDO.Message" )

    ' Fill in the field values
    With objEmail
        .From     = myFrom
        .To       = myTo
        ' Other options you might want to add:
        ' .Cc     = ...
        ' .Bcc    = ...
        .Subject  = mySubject
        .TextBody = myTextBody
        '.HTMLBody = myHTMLBody
        If IsArray( myAttachment ) Then
            For i = 0 To UBound( myAttachment )
                .AddAttachment Replace( myAttachment( i ), "\", "\\" ),"",""
            Next
        ElseIf myAttachment <> "" Then
            .AddAttachment Replace( myAttachment, "\", "\\" ),"",""
        End If
        If mySMTPPort = "" Then
            mySMTPPort = 25
        End If
        With .Configuration.Fields
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = mySMTPServer
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = mySMTPPort
            .Update
        End With
        ' Send the message
        .Send
    End With
    ' Return status message
    If Err Then
        EMail = "Error sending install notice to " & myto & " " & Err.Number & ": " & Err.Description
        Err.Clear
    Else
        EMail = "Install notice sent to " & myTo
    End If

    ' Release the e-mail message object
    Set objEmail = Nothing
    ' Restore default error handling
    On Error Goto 0
End Function
