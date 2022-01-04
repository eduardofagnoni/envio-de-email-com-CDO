<%

    Set objCDOSYSMail = Server.CreateObject("CDO.Message")
    Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")

    objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
    objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30

    objCDOSYSCon.Fields.update

    Set objCDOSYSMail.Configuration = objCDOSYSCon
    objCDOSYSMail.From = "faleconosco@bluetree.com.br"
    objCDOSYSMail.To = "eduardo@brancozulu.com.br"
    objCDOSYSMail.Subject = "Teste de envio"

    objCDOSYSMail.HtmlBody = "Mensagem enviada no teste..."

    objCDOSYSMail.Send

    set objCDOSYSMail = nothing
    set objCDOSYSCon = nothing

    response.write("Fim do arquivo... Mensagem enviada.")

%>