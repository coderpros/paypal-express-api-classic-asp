﻿<%
    Const PayPalApiClientId = "XXXXXX"
    Const PayPalApiSecret = "XXXXXX"
    
    ' The below is VERY important. It is a **base64 encoded** copy of PaypalApiClientId:PayPalApiSecret (the : is necessary).
    Const PayPalEncodedAuth = "QWVYME02bmM2N2ItWGJOb290cjVFdktmWmNieUNCSW94YnNKNElNTHlORGRJc2NPbWZGTUlReXdLUTlleTFhTm8wZmRpVmRNUm11ZXoyY246RUtpSnBJVjhjWmt5dmNBVGlkTUNDNGRITFRNU2VGRTJoM2VnOV9hbGhoOEtQZXZkTmhBeGJHR21Va2o2aWNRRV9mN3RJNUpLTERtbGc1V3o="
    Const PaypalApiEndPoint = "https://api.sandbox.paypal.com/v1/"
    
    Const SslClientCertFriendlyName = "LOCAL_MACHINE\My\ServerXMLHTTP" 'The path and friendly name. Either obtained from your host or from Start => "Manage Computer Certificates" => Personal (could possibly be elsewhere)
%>