<%@  language="vbscript" %>
<% Option Explicit %>

<!--#include file="inc/adovbs.inc"-->
<!--#include file="inc/constants.asp"-->
<!--#include file="inc/common.asp"-->
<!--#include file="inc/PaypalObject.class.asp"-->
<%
    SetLCID(Locale) 'Manually set the locale. *NB: Make sure the server supports it.
%>
<!doctype html>

<html lang="en">
<head>
    <meta charset="utf-8">

    <title>Store Simulator ~ Paypal Express Example Implementation</title>
    <meta name="description" content="A Paypal express example implementation with JavaScript and Classic ASP">
    <meta name="author" content="coderPro.net">

    <link rel="stylesheet" href="Content/styles.css?v=1.0">
</head>
<body>
    <form method="post" action="cart.asp">
        <div id="store-container">
            <h1>Store Simulator</h1>
            <table>
                <thead>
                    <tr>
                        <th>Item</th>
                        <th>Unit Price</th>
                        <th>Quantity</th>
                    </tr>
                </thead>
                <tbody>
                    <tr style="display: none;">
                        <td>
                            <input type="text" id="ItemName" name="ItemName" maxlength="255" required />
                        </td>
                        <td>
                            <input type="number" step="0.01" id="ItemPrice" name="ItemPrice" required />
                        </td>
                        <td>
                            <input type="number" step="1" id="ItemQuantity" name="ItemQuantity" required />
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <input type="text" id="ItemName" name="ItemName" maxlength="255" value="Test Item 1" required />
                        </td>
                        <td>
                            <input type="number" step="0.01" id="ItemPrice" name="ItemPrice" value="25.00" required />
                        </td>
                        <td>
                            <input type="number" step="1" id="ItemQuantity" name="ItemQuantity" value="1" required />
                        </td>
                    </tr>
                </tbody>
                <tfoot>
                    <tr>
                        <td>
                            <input type="button" name="AddButton" id="AddButton" value="Add Row" class="btn btn-secondary" />
                        </td>
                        <td>
                            <input type="button" name="RemoveButton" id="RemoveButton" value="Remove Row" class="btn btn-secondary" />
                        </td>
                        <td>
                            <input type="submit" name="SubmitButton" id="SubmitButton" value="Submit Form" class="btn btn-primary" />
                        </td>
                    </tr>
                </tfoot>
            </table>
        </div>
    </form>

    <script src="Scripts/jquery-3.6.0.min.js"></script>
    <script src="Scripts/scripts.js"></script>
</body>
</html>
