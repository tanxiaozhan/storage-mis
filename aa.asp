<!-- #include file="data\config.asp" -->

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
</head>

<%
txt1=request("t1")
txt2=""
if txt1<>"" then
	txt2=gs(txt1,4)
end if


%>



<body>

<form method="POST" action="aa.asp">
  <p><input type="text" name="T1" size="20"></p>
  <p><input type="text" name="T2" size="20" value="<%=txt2%>"></p>
  <p><input type="submit" value="提交" name="B1"><input type="reset" value="全部重写" name="B2"></p>
</form>

</body>

</html>
