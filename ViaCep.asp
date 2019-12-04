<!--#include file="ViaCepClass.asp" -->
<!DOCTYPE html>
<html>
<head>
	<title>ViaCep</title>
	<meta charset="iso-8859">
</head>
<body>

	<%
	'Exemplos de uso
	Dim buscaCep, endereco, cep 'Declara as variáveis
	Set buscaCep = new ViaCep 'Inicializa a classe

	cep = 57038740

	endereco = buscaCep.buscar(cep) 'Guarda o retorno do método numa variável
	Response.Write "<pre>" & endereco & "</pre><br>" 'Imprime o resultado na tela


	buscaCep.formato = "xml" 'Método que modifica o formato de retorno
	endereco = buscaCep.buscar(cep) 'Guarda o retorno do método numa variável
	Response.Write "<textarea rows='20' cols='40'>" & endereco & "</textarea><br>"


	'Exemplo de erro:
	' buscaCep.Cep = "11111"

	Set buscaCep = Nothing
	%>

</body>
</html>
