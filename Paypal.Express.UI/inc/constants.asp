<%
    Const PayPalApiClientId = "XXXXXX"
    Const PayPalApiSecret = "XXXXXX"
    
    ' The below is VERY important. It is a **base64 encoded** copy of PaypalApiClientId:PayPalApiSecret (the : is necessary).
    Const PayPalEncodedAuth = "XXXXXX"
    Const PaypalApiEndPoint = "https://api.sandbox.paypal.com/v1/"
    
    'The path and friendly name of the CLIENT (or client/server) ssl cert. Either obtained from your host or from Start => "Manage Computer Certificates" => Personal (could possibly be elsewhere)
    Const SslClientCertFriendlyName = "LOCAL_MACHINE\My\ServerXMLHTTP"
    Const Locale = "en-us"
    Const PaypalDebug = False
%>