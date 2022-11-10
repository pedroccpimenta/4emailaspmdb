<!DOCTYPE HTML>
<HEAD>	
<meta charset="utf-8"/>
<meta http-equiv="Content-Type" content="text/html;charset=ISO-8859-1">
</HEAD>	
<%	
	dibug=1
	'response.end
	
	Const cdoSendUsingMethod        = "http://schemas.microsoft.com/cdo/configuration/sendusing"
    Const cdoSendUsingPort          = 2
    Const cdoSMTPServer             = "http://schemas.microsoft.com/cdo/configuration/smtpserver"
    Const cdoSMTPServerPort         = "http://schemas.microsoft.com/cdo/configuration/smtpserverport"
    Const cdoSMTPConnectionTimeout  = "http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"
    Const cdoSMTPAuthenticate       = "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
   

    Const cdoBasic                  = 1
    Const cdoSendUserName           = "http://schemas.microsoft.com/cdo/configuration/sendusername"
    Const cdoSendPassword           = "http://schemas.microsoft.com/cdo/configuration/sendpassword"
    Const cdoSMTPUseSSL             = "http://schemas.microsoft.com/cdo/configuration/smtpusessl"
    Const cdoSMTPUseTLS             = "http://schemas.microsoft.com/cdo/configuration/smtpusetls"

    Dim objConfig  ' As CDO.Configuration
    Dim objMessage ' As CDO.Message
    Dim Fields     ' As ADODB.Fields
    Dim DB
    Set DB=Server.CreateObject("ADODB.Connection")

    srv = request.servervariables("SERVER_NAME")
    response.write "<srv>"&srv&"</srv>"

	bdad=Server.Mappath("./db/emails.mdb")
	on error resume next 
	 
	DB.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & bdad)
	
	if err<>0 then
		response.write "<errdb>" & err & " - " & err.description & "</errdb>"
		response.flush
		response.end
	else
		response.write "Base de dados operacional<br>"
		response.write "<errdb>None</errdb>"		
	end if

	' Retrieve das variáveis
	nome=request.querystring("n")
	email=request.querystring("e")
	mensagem=request.querystring("m")
	response.write "<args>" + nome+"|"+ "|"+email + "|"+mensagem + "</args>"

	dim qa
	set qa = Server.CreateObject("ADODB.Recordset")
	
	Set objConfig = Server.CreateObject("CDO.Configuration")
    Set Fields = objConfig.Fields

	With Fields
		.Item(cdoSendUsingMethod)       = cdoSendUsingPort

		.Item(cdoSMTPServer)            = "smtp.gmail.com"					
		.Item(cdoSMTPServerPort)        = 465
		'.Item(cdoSMTPServerPort)        = 587
		.item (cdoSMTPUseSSL )=true

		.Item(cdoSendUserName)          = "pedroccpimenta@gmail.com"
		.Item(cdoSendPassword)          = "*password*"
		   
        .Item(cdoSMTPConnectionTimeout) = "30"
        .Item(cdoSMTPAuthenticate) = 1
        '.Item(cdoSMTPUseSSL)       = 1
        '.Item(cdoSMTPUseTLS)       = 1
        .Update
	End With
               
    Set objMessage = Server.CreateObject("CDO.Message")
    Set objMessage.Configuration = objConfig
	
	nometoi=left(nome, len(nome)-1)
	nometoi=right(nometoi, len(nometoi)-1)
	nometoi=replace(nometoi, "'", "´")	
	mensagemtoi=replace(mensagem, "'", "´")
		
	str="<body><font face=Lato size=3>Dear <b>"+nometoi+"</b><br>Thank you for your message:<hr color=lime>" + mensagem + "<hr color=lime>Pedro Pimenta<hr><font size=2>Disclaimer:<br>blá blá blá....</body>"
			
	
	response.write "<br>STR EMAIL<span style='background:yellow'>"+str+"</span>"

	' Full SQL, sem err.sending.mail
	str=replace(str, "'", "´") 
	sql =     "insert into registo (ip, data, hora,  email, nota)"
	sql=sql& " values ('"+Request.ServerVariables("REMOTE_ADDR")+"','"&date&"','"&time&"',"
	sql=sql& email & "," & mensagemtoi &")"


	response.write "<sql>"& sql & "</sql><hr red>"

	Err.clear	
	on error resume next
	
	tosend="<html><head><meta http-equiv='Content-Type' content='text/html;charset=ISO-8859-1'></head><body>"
	tosend=tosend+str
	tosend=tosend+"</html>"
	nometos=left(nome, len(nome)-1)
	nometos=right(nometos, len(nometos)-1)
	
	Err=0	
	on error resume next

	response.write "To:"
	response.write replace(email, "'", "")
    With objMessage
		.To 		= replace(email, "'", "")
		.Cc = "pedroccpimenta@gmail.com"
		.From     	= "pedroccpimenta@gmail.com" 
		.HTMLbody 	= tosend
        .Subject  	= "Thank you for your message."
		.Send
	End With		
	
	' fim do envio do email propriamente dito ...
	' Preparação do comando sql
	
	if Err<>0 then
		response.write "<br><br>Err on sending email:" & Err & " - " & err.description
		erro=Err & " - " & err.description		
	else
		response.write "<br> no error on sending email" 
		erro="-"
	end if
	
	response.write "<hr><sql>"&sql & "</sql>"
	
	err.clear
	db.execute sql
	
	if Err<>0 then
		response.write "<sqlerr>Err SQL:" & Err & " - " & err.description & "</sqlerr>"
	else
		response.write "<sqlerr>-</sqlerr>" 
		response.write "<br>SEM ErROs sql"
	end if
	

		sql2="insert into accoes (data, ip,  enviodemail, email, erremail)"
		sql2=sql2&" values ('"& now() &"', '"&ip&"',  'sim', '"&str&"','"&erro&"')"
	
	response.write "<sqla>"+sql2+"</sqla>"
	
	db.execute sql2
	
	Set objMessage.Configuration = Nothing 
	Set objMessage = Nothing
              	
	Set objConfig = Nothing
    Set Fields = Nothing
	
%>	