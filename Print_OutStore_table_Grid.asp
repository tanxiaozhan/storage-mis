<%
BillNo=Request("BillNo")
Date1=Request("PrintDate")
ClientID=Request("ClientID")



''''''''''''''''''''''''''
if Session("PrintChoice_maozhong")="print_buxianshimaozhong" then
XIANSHI="OFF"
else
XIANSHI="ON"
end if
'''''''''''''''''''''''''''''''''

%>


<%
   if Session("PrintChoice")="OutStoreTable" then
%>

		<script language="javascript">
  				 window.open('Print_OutStore_Table.asp?ClientID=<% =ClientID %>&BillNo=<%=BillNo %>&PrintDate=<% =Date1 %>&XIANSHI=<%=XIANSHI%>','打印出库通知单','toobar=yes,menubar=yes,scrollbars=yes,resizable=yes') 
		</script>
<%
	else
%>
		<script language="javascript">
  				 window.open('Print_OutStore_Table_Receive.asp?ClientID=<% =ClientID %>&BillNo=<%=BillNo %>&PrintDate=<% =Date1 %>&XIANSHI=<%=XIANSHI%>','打印出库通知单','toobar=yes,menubar=yes,scrollbars=yes,resizable=yes') 
		</script>

<%
	end if
	
	
	if session("varClientID")="All" then
		url="OutStore_Print_All_Frame.asp"
	else
		url="OutStore_Print_Frame.asp"
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
