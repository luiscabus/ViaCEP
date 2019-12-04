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
		CepPattern = "^\d{8}$"
		ViaCepUrl = "https://viacep.com.br/ws/"
		ViaCepFormato = "json"
	End Sub


	Public Function buscar(pcep)
		Call validar_cep(pcep)
 
		Dim xmlHttp, xmlHttpResponse
		Set xmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP") 'Conecta via GET ao Webservice ViaCep
		  Call xmlHttp.open("GET", ViaCepUrl&"/"&CepNumber&"/"&ViaCepFormato&"/", false)
		  Call xmlHttp.Send()

		If xmlhttp.Status<>200 Then 'Valida o Status 200 de retorno
			buscar = "Serviço indisponível. Status != 200."
			Exit Function
		End If

		xmlHttpResponse = xmlHttp.responseText

		Set xmlhttp = Nothing 

		buscar = xmlHttpResponse 'Se não houve erros, retorna a resposta do webservice
	End Function


	Private Sub validar_cep(pcep)
		If IsObject(pcep) Then 'Validate if input is not an object
			Err.Raise vbObjectError + 1000, "ViaCep Class", _
			"O CEP deve ser informado com 8 dígitos, sem pontuação. Ex.:57036000"
			Exit Sub
		End If

		Dim objRegExp
		Set objRegExp = New regexp
		objRegExp.Pattern = CepPattern

		If objRegExp.Test(pcep) Then 'Make sure it matches the pattern 
			CepNumber = pcep 'Set property CepNumber value if input is correct
		Else
			Err.Raise vbObjectError + 1000, "ViaCep Class", _
			"O CEP deve ser informado com 8 dígitos, sem pontuação. Ex.:57036000"
		End If

	  	Set objRegExp = Nothing
	End Sub


	Public Property Let formato(pformato)
		If IsObject(pformato) Then
			Err.Raise vbObjectError + 1000, "ViaCep Class", _
			"Formato inválido. Opções são json, jsonp, xml, piped or querty."
			Exit property
		End If

		Select Case pformato 'Opções disponíveis
			Case "json"
				ViaCepFormato = "json"
			Case "jsonp"
				ViaCepFormato = "jsonp"
			Case "xml"
				ViaCepFormato = "xml"
			Case "piped"
				ViaCepFormato = "piped"
			Case "querty"
				ViaCepFormato = "querty"
			Case Else
				Err.Raise vbObjectError + 1000, "Information Class", _
				"Formato inválido. Opções são json, jsonp, xml, piped or querty."
		End Select
	End Property
End Class
%>
