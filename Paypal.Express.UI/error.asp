<%@  language="vbscript" %>
<% Option Explicit %>
<!--#include file="inc/adovbs.inc"-->
<!--#include file="inc/constants.asp"-->
<!--#include file="inc/common.asp"-->
<!--#include file="inc/PaypalObject.class.asp"-->
<!doctype html>

<html lang="en">
<head>
    <meta charset="utf-8">

    <title>My Cart ~ Paypal Express Example Implementation</title>
    <meta name="description" content="A Paypal express example implementation with JavaScript and Classic ASP">
    <meta name="author" content="coderPro.net">

    <link rel="stylesheet" href="Content/styles.css?v=1.0">
</head>

<body>
        <section id="container">
            <h1>Error</h1>
            <h2><%= Request.QueryString("e") %></h2>
          
        </section>

    <script src="Scripts/jquery-3.6.0.min.js"></script>
    <script src="Scripts/scripts.js"></script>
</body>
</html>