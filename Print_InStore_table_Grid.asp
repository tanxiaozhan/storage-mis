<%
ClientID=Request("ClientID")
Date1=Request("PrintDate")
BillNo=Request("BillNo")

%>
<script language="javascript">
  		 window.open('Print_InStore_Table.asp?BillNo=<% = BillNo %>&ClientID=<%=ClientID %>&PrintDate=<% =Date1 %>','打印入库通知单','toobar=yes,menubar=yes,scrollbars=yes,resizable=yes') 
</script>

<%
	if session("varClientID")="All" then
		url="InStore_Print_All_Frame.asp"
	else
		url="InStore_Print_Frame.asp"
	end if	
%>


<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
<meta http-equiv=refresh content="0; url=<%=url%>">

</head>

<body>

</body>

</html>
