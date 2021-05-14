<%@  language="vbscript" %>
<% Option Explicit %>
<!--#include file="inc/adovbs.inc"-->
<!--#include file="inc/constants.asp"-->
<!--#include file="inc/common.asp"-->
<!--#include file="inc/PaypalObject.class.asp"-->
<%
    Dim idx, itemCount, total
    Dim aryItemName, aryItemPrice, aryItemQty
    Dim accessToken, clientToken
    Dim debug: debug = False

    itemCount = 0
    total = 0

    SetLCID("en-gb") 'Manually set the locale. *NB: If locale is not available on the server it WILL explode.

    If InStr(1, Request.Form("ItemName"), ",") > 0 Then
        aryItemName = Split(Request.Form("ItemName"), ",")
        aryItemPrice = Split(Request.Form("ItemPrice"), ",")
        aryItemQty = Split(Request.Form("ItemQuantity"), ",")

        itemCount = UBound(aryItemName)
    End If

    Dim paypal: Set paypal = New PaypalObject
    
    accessToken = paypal.GetAccessToken()("access_token")
        
    clientToken = paypal.GetClientToken(accessToken)("client_token")

    If debug = True Then 
        Response.Write "Access Token: " & accessToken & "<br />"
        Response.Write "Client Token: " & clientToken & "<br />"
    End If
%>
<!doctype html>

<html lang="en">
<head>
    <meta charset="utf-8">

    <title>My Cart ~ Paypal Express Example Implementation</title>
    <meta name="description" content="A Paypal express example implementation with JavaScript and Classic ASP">
    <meta name="author" content="coderPro.net">

    <link rel="stylesheet" href="Content/styles.css?v=1.0">

    <script src="https://www.paypal.com/sdk/js?debug=true&intent=capture&components=buttons&integration-date=2021-05-09&client-id=<%= PayPalApiClientId %>" data-client-token="<%= clientToken %>"></script>
    <!-- 
        Developer Notes: 
        1) Turn off debug before moving to prod.
        2) Integration-Date stops the paypal js sdk from automatically updating to the latest and potentially breaking the implementation.
    -->
</head>

<body>
    <form method="post" action="cart.asp">
        <section id="store-container">
            <h1>My Cart</h1>
            <table border="1">
                <thead>
                    <tr>
                        <th>Item</th>
                        <th>Unit Price</th>
                        <th>Quantity</th>
                        <th>Subtotal</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                        
                    If itemCount > 0 Then
                        For idx = 0 To itemCount
                            Dim subTotal
                            
                            subTotal = CDbl(aryItemPrice(idx)) * CDbl(aryItemQty(idx))
                            total = total + subTotal
                    %>
                    <tr>
                        <td><%= aryItemName(idx) %></td>
                        <td><%= FormatCurrency(aryItemPrice(idx)) %></td>
                        <td><%= aryItemQty(idx) %></td>
                        <td><%= FormatCurrency(subTotal) %></td>
                    </tr>
                    <%
                        Next
                    %>

                    <%Else%>
                    <tr>
                        <td><%= Request.Form("ItemName") %></td>
                        <td><%= Request.Form("ItemPrice")%></td>
                        <td><%= Request.Form("ItemQty")%></td>
                    </tr>
                    <%End If %>
                </tbody>
                <tfoot>
                    <tr>
                        <td colspan="3" style="text-align: right; font-weight: bold;">Total:
                        </td>
                        <td><%= FormatCurrency(total) %></td>
                    </tr>
                </tfoot>
            </table>
            <br />
            <div id="paypal-button-container"></div>
        </section>
    </form>

    <script src="Scripts/jquery-3.6.0.js"></script>
    <script src="Scripts/scripts.js"></script>

    <script>
        paypal.Buttons({
            style: {
                layout: 'vertical'
            },
            createOrder: function (data, actions) {
                return actions.order.create({
                    purchase_units: [{
                        items: [
                        <%For idx = 0 To itemCount%>
                        {
                            unit_amount: {
                                currency_code: "USD",
                                value: "<%= trim(aryItemPrice(idx)) %>"
                            },
                            quantity: "<%= trim(aryItemQty(idx)) %>",
                            name: "<%= trim(aryItemName(idx)) %>"
                        },
                        <%Next%>
                        ],
                        amount: {
                            value: "<%= trim(total) %>",
                            currency_code: "USD",
                            breakdown: {
                                item_total: {
                                    currency_code: "USD",
                                    value: "<%= trim(total) %>"
                                }
                            }
                        }
                    }]	
                });
            },
            onApprove: function (data, actions) {
                return actions.order.capture().then(function (details) {
                    window.location.href = '/success.html';
                });
            },
            onCancel: function (data, actions) {
                return actions.redirect();
            },
            onError: function (err) {
                // Show an error page here, when an error occurs
            }
        }).render("#paypal-button-container");
    </script>
</body>
</html>
