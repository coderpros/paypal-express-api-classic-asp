<%
    Const PayPalApiClientId = "ARU_T7OE10p7U0xTQbPvQGxGGi2684vk5cLVGXnLaTBdDIhZbZ4UXNLSfCUTMDJawRfQqbDeAtE429WC"
    Const PayPalApiSecret = "EHr51LkGmwaNh1ag2iemBCY-wYabUfsxvWu0_fs_6gvaYwuFIUPuUgTkuByLsDI3Hq4rGPBZhSYNhSdY"
    ' The below is VERY important. It is a **base64 encoded** copy of PaypalApiClientId:PayPalApiSecret (the : is necessary).
    Const PayPalEncodedAuth = "QVJVX1Q3T0UxMHA3VTB4VFFiUHZRR3hHR2kyNjg0dms1Y0xWR1huTGFUQmRESWhaYlo0VVhOTFNmQ1VUTURKYXdSZlFxYkRlQXRFNDI5V0M6RUhyNTFMa0dtd2FOaDFhZzJpZW1CQ1ktd1lhYlVmc3h2V3UwX2ZzXzZndmFZd3VGSVVQdVVnVGt1QnlMc0RJM0hxNHJHUEJaaFNZTmhTZFk="
    Const PaypalApiEndPoint = "https://api.sandbox.paypal.com/v1/"
     'The path and friendly name of the CLIENT (or client/server) ssl cert. Either obtained from your host or from Start => "Manage Computer Certificates" => Personal (could possibly be elsewhere)
    Const SslClientCertFriendlyName = "LOCAL_MACHINE\My\ServerXMLHTTP"
    Const Locale = "en-us"
    Const PaypalDebug = False
%>