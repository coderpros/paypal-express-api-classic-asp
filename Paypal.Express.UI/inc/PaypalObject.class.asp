<!--#include file="jsonObject.class.asp" -->
<%
Class PaypalObject
	Private i_clientToken
	Dim JSON

	Public Sub Class_Initialize
		Set JSON = New JsonObject
	End Sub

	Public Sub Class_Terminate
		Set JSON = Nothing
	End Sub 

	Public Property Get ClientApiKey
		ClientToken = i_clientToken
	End Property
	
	Public Property Let ClientApiKey(value)
		i_clientToken = value
	End Property
	
	' Gets the access token from PayPal.
	Public Function GetAccessToken()
		Dim oHttp: Set oHttp = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")

		With oHttp
			Call .setOption(3, SslClientCertFriendlyName)
			
			Call .open("POST", PayPalApiEndPoint & "oauth2/token", False)

			Call .setRequestHeader("Accept", "application/json")
			Call .setRequestHeader("Accept-Language", "en_US")
			Call .setRequestHeader("Authorization", "Basic " & PayPalEncodedAuth)
		End With

		oHttp.send("grant_type=client_credentials")

		' Deserialize the json response into an object.
		Set GetAccessToken = DeserializeJson(oHttp.responseText)
	End Function

	Public Function GetClientToken(accessToken)
		Dim oHttp: Set oHttp = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")

		With oHttp
			Call .setOption(3, SslClientCertFriendlyName)
			
			Call .open("POST", PayPalApiEndPoint & "identity/generate-token", False)

			Call .setRequestHeader("Accept", "application/json")
			Call .setRequestHeader("Accept-Language", "en_US")
			Call .setRequestHeader("Authorization", "Bearer " & accessToken)
		End With

		oHttp.send()

		' Deserialize the json response into an object.
		Set GetClientToken = DeserializeJson(oHttp.responseText)
	End Function
	Private Function DeserializeJson(jsonString)
		Dim jsonArr
		Set jsonArr = JSON.parse(jsonString)
		
		Set DeserializeJson = jsonArr
		Set jsonArr = nothing
	End Function
End Class
%>