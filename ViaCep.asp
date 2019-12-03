<!--#include file="ViaCepClass.asp" -->
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
