<!-- #include file="data\config.asp" -->

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
</head>

<%
    Set rs =server.createobject("adodb.recordset")

    	sql1="update outstore set Weight=net*pack1/1000"
		conn.Execute(sql1)


%>


<body>

</body>

</html>
