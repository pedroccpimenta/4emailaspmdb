<!DOCTYPE HTML>
<HEAD>	
<meta charset="utf-8"/>
<meta http-equiv="Content-Type" content="text/html;charset=ISO-8859-1">

<TITLE>Form2BD&email</TITLE>

<script>

function prep(x)
{	h=document.getElementById(x);
	hs=h.value.replace(/'/gi, "`");
	hs.replace(/&/g, "%26");
	return hs;
}


function submete()
{	email=prep('email');
		
	if ((email.length==0))  
	{	alert("email cannot be blank!");	
		return;
	}
	
	str="./bdem.asp?n="+"'"+prep('nome')+"'"+'&e='+"'"+email+"'&m='"+prep('mensagem')+"'"
	
	console.log(str)
			
	var xmlhttp;
	
	if (window.XMLHttpRequest) // code for IE7+, Firefox, Chrome, Opera, Safari
	{	xmlhttp=new XMLHttpRequest();
	}
	else // code for IE6, IE5
	{	xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
	}

	xmlhttp.onreadystatechange=function()
	{
		if (xmlhttp.readyState==4 && xmlhttp.status==200)
		{	tok=xmlhttp.responseText;
			// alert('tok:'+tok);
			document.getElementById("resposta").innerHTML=tok
			h2=document.getElementById("resposta");
							
			var p=[];
			p.push('errdb');
			p.push('args');
			p.push('sql');
			p.push('sqla');
			
			p.push('SQERR');
			p.push('DEBUG');
		
			h2.innerHTML="&lterrdb>"+tok.substring(tok.indexOf("<errdb>"),tok.indexOf("</errdb>")) +"&lt/errdb>";
		
		for(i=1;i<p.length;i++)
		{	h2.innerHTML=h2.innerHTML+"<br>"+"&lt"+p[i]+">"+tok.substring(tok.indexOf("<"+p[i]+">"),tok.indexOf("</"+p[i]+">")) +"&lt/"+p[i]+">" ;
		}
		
		if (tok.substring(tok.indexOf("<SQERR>"),tok.indexOf("</SQERR>")) !='0')
		{	document.getElementById('resposta').style.display='';
		}
		else
		{	document.getElementById('resposta').style.background='whitesmoke';
		}
		
		//document.getElementById(campo+k).style.background='white';
		document.getElementById('status').innerHTML=document.getElementById('status').innerHTML+'<br>sav(ed)...';		
		}
	}
	
	try
	{	// alert(str);
		xmlhttp.open("POST",str,false);			
		xmlhttp.send();
	}
	catch(err)
	{	alert(err.message);
	}
	
}
</script>

</HEAD>
<BODY style='font-family: lato;'><center>


<table border=1 cellpadding=0 cellspacing=0 style='width:50%'><tr><td align=center>

	<table border=1 style='width:50%;'>
		<tr><td>Name<td><input type=text  id=nome style='width:96%;'>
		<tr><td>email<td><input type=text  id=email style='width:96%;' >
		<tr><td>message<td><textarea id=mensagem style='width:97%;height:70px;font-size:12px;color:navy;'></textarea>
	</table>

	<input id=sutao type=submit value='Submit' title='Click to submit'  onclick='submete();'></span>

	</table>
<hr color=lime>
<div id=resposta></div>
<div id=status></div>

</BODY></HTML>
