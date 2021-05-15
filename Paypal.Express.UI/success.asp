<%@  language="vbscript" %>
<% Option Explicit %>
<!--#include file="inc/adovbs.inc"-->
<!--#include file="inc/constants.asp"-->
<!--#include file="inc/common.asp"-->
<!--#include file="inc/PaypalObject.class.asp"-->
<%
    Dim orderId, orderDate, intent, status, payerId, payerEmail
    Dim items, quantities

    orderId = Request.Form("orderId")
    orderDate = Request.Form("orderDate")
    intent = Request.Form("intent")
    status = Request.Form("status")
    payerId = Request.Form("payerId")
    payerEmail = Request.Form("payerEmail")

    items = Split(Request.Form("items"), ",")
    quantities = Split(Request.Form("quantitites"), ",")

    'Pseudo code:
    '   INSERT INTO tblOrders(orderId, orderDate, status, payerId, payerEmail) VALUES(?, ?, ?, ?, ?)
    '   FOR i = LBound(items) To UBound(items)
    '       INSERT INTO tblOrderLine(orderId, item, quantity) VALUE(?, ?, ?)
    '   NEXT
%>