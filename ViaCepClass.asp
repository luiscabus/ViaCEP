<%
'
' ViaCep Consulta API Class, 1.0 Dec, 03 - 2019.
' Written by Luis Alberto Cabus
' https://github.com/luiscabus/ViaCEP
' https://linkedin.com/in/lapcs
' API Provider: https://viacep.com.br/
' License MIT 
' 

Option Explicit

Class ViaCep
	Private CepNumber
	Private CepPattern
	Private ViaCepUrl
	Private ViaCepFormato
	

	Private Sub Class_Initialize
		CepNumber = ""
		CepPattern = "^\d{5}-\d{3}$"
		ViaCepUrl = "https://viacep.com.br/ws/"
		ViaCepFormato = "json"
	End Sub


	Public Function busca(pcep)
		Call validate_cep(pcep)

		'Connects to ViaCep and retrieves the information
		Dim xmlHttp, xmlHttpResponse
		Set xmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
		  Call xmlHttp.open("GET", ViaCepUrl&"/"&CepNumber&"/"&ViaCepFormato&"/", false)
		  Call xmlHttp.SetRequestHeader("Content-Type", "text/html")
		  Call xmlHttp.setRequestHeader("CharSet", "UTF-8")
		  Call xmlHttp.Send()
		  xmlHttpResponse = xmlHttp.responseText

		'If response status isn't 200, returns message
		If xmlhttp.Status<>200 Then
			busca = "Serviço indisponível. Status != 200."
			Exit Function
		End If

		Set xmlhttp = Nothing

		busca = xmlHttpResponse
	End Function


	Private Sub validate_cep(pcep)
		If IsObject(pcep) Then 'Validate if input is not an object
			Err.Raise vbObjectError + 1000, "Information Class", _
			"Invalid format for CEP. Must be in #####-### format."
			Exit Sub
		End If

		Dim objRegExp
		Set objRegExp = New regexp
		objRegExp.Pattern = CepPattern

		If objRegExp.Test(pcep) Then 'Make sure it matches the pattern 
			CepNumber = pcep 'Set property CepNumber value if input is correct
		Else
			Err.Raise vbObjectError + 1000, "Information Class", _
			"Invalid format for CEP. Must be in #####-#### format."
		End If

	  	Set objRegExp = Nothing
	End Sub


	Public Property Let formato(pformato)
		If IsObject(pformato) then
			Err.Raise vbObjectError + 1000, "Information Class", _
			"Invalid format option. Must be JSON, JSONP, XML, PIPED or QUERTY."
			Exit property
		End If

		Select Case pformato
			Case "JSON"
				ViaCepFormato = "json"
			Case "JSONP"
				ViaCepFormato = "jsonp"
			Case "XML"
				ViaCepFormato = "xml"
			Case "PIPED"
				ViaCepFormato = "piped"
			Case "QUERTY"
				ViaCepFormato = "querty"
			Case Else
				Err.Raise vbObjectError + 1000, "Information Class", _
				"Invalid format option. Must be JSON, JSONP, XML, PIPED or QUERTY."
		End Select
	End Property


	' Public Property Let Cep(pcep)
	' 	If IsObject(pcep) then
	' 		Err.Raise vbObjectError + 1000, "Information Class", _
	' 		"Invalid format for CEP.  Must be in #####-### format."
	' 		Exit property
	' 	End If

	' 	Dim objRegExp
	' 	Set objRegExp = New regexp
	' 	objRegExp.Pattern = "^\d{5}-\d{3}$"

	' 	If objRegExp.Test(pcep) Then 'Make sure the pattern fits
	' 		CepNumber = pcep
	' 	Else
	' 		Err.Raise vbObjectError + 1000, "Information Class", _
	' 		"Invalid format for CEP.  Must be in #####-#### format."
	' 	End If

	'   	Set objRegExp = Nothing
	' End Property

	' Public Property Get Cep()
	' 	Cep = CepNumber
	' End Property

End Class
%>


<!DOCTYPE html>
<html>
<head>
	<title>ViaCep</title>
	<!-- <meta charset="utf-8"> -->
	<meta charset="iso-8859">
</head>
<body>

	<%
	'Examples of use
	Dim buscaCep, endereco 'Declare variables
	Set buscaCep = new ViaCep 'Initialize class

	endereco = buscaCep.busca("57038-740") 'Stores the return of the method into a variable
	Response.Write "<pre>" & endereco & "</pre><br>"


	buscaCep.formato = "XML" 'Method to set the format of the return
	endereco = buscaCep.busca("14015-130") 'Stores the return of the method into a variable
	Response.Write "<textarea rows='20' cols='40'>" & endereco & "</textarea><br>"


	'Exemplo de erro:
	' buscaCep.Cep = "11111"

	Set buscaCep = Nothing
	%>

</body>
</html>
