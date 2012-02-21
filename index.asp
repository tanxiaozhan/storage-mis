

 <SCRIPT LANGUAGE ="VBSCRIPT">
<!--
	Sub b1_Click()
		my.submit		
	End Sub



-->
</Script>


<script language="javascript">
   function openwin_all()
     {
       window.navigate('index.asp?body=all','仓库管理系统','scrollbars=yes,resizable=yes') 
     }   
	 
	 
	   function openwin_onwork()
     {
       window.navigate('index.asp?body=onwork','仓库管理系统','scrollbars=yes,resizable=yes') 
     }   
	  
	 
	 
	 
	 
</script>









<%

		''''''''''''''''''''''''''''''''''''''START修改毛重
		if Request("body")="all"  then
		       		  Session("PrintChoice_body")="print_allbody" 
		
	    else
	
			   	        Session("PrintChoice_body")="print_onworkbody"
	    end if
	''''''''''''''''''''''''''''''''''''''END

%>







<%
    if request("op")="reset" then
       session("username")="" 
       session("userright")=""
    end if
    
    dim conn
    dim  connectstr
    set  conn=server.createobject("adodb.connection")
    connectstr="provider=microsoft.jet.oledb.4.0; jet oledb:database password=office;data source=" & server.mappath("data\SMIS.mdb")
    conn.Open connectstr 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>仓库管理系统登录</title>
</head>

<script language="javascript">
   function openwin()
     {
       window.open('Client_list.htm','仓库管理系统','scrollbars=yes,resizable=yes') 
     }   
</script>

<body bgcolor="#007E23" topmargin="0" leftmargin="0" background="image/bg5.jpg" >
<table border="0" width="100%" height="100%">
<tr><td align="center">
<table border="0" width="804" height="400" bgcolor="#007E23">

  <tr>
    <td width="680" height="20" align="center" colspan="3" bgcolor="#007E23">
    
    
    <p align="left"><img border="0" src="image/bg3.jpg" width="600" height="20">
    
    
    </td>
  </tr>

  <tr>
    <td width="207" height="312" align="center" bgcolor="#007E23">
    
      <p align="left"><img border="0" src="image/bg2.jpg" width="300" height="364"></p>
    
    
    </td>
    <td width="269" height="312" align="center" bgcolor="#007E23">
    <form name="my" action="login.asp">
    <table style="border: 1 double #FFFFFF" bgcolor="#339933">
    
    <tr>
    <td>
      <p align="left"><font color="#FFFFFF"><b>用户名：</b></font></td>
    <td bgcolor="#339933"><select size="1" name="us">
     <%
         sql="select * from userInfo"
         Set rs =server.createobject("adodb.recordset")
         rs.Open sql,conn,3,3
         while not rs.eof
		 
		 
		 
		 
		
	if Request("body")<>"all" then 	 
		 
           Select Case rs("user")
             Case "陈妃英"    ' rs.movenext
             Case "王辉"     ' rs.movenext
             Case "王月强"   ' rs.movenext
			  Case "陈娟"   ' rs.movenext
			   Case "单文凤"   ' rs.movenext
			    Case "傅桃嫦"   ' rs.movenext 
				 Case "黄菊"   ' rs.movenext
				  Case "黄永权"   ' rs.movenext
				   Case "罗敏洁"   ' rs.movenext
				    Case "潘钊雄"   ' rs.movenext
					 Case "彭婵"   ' rs.movenext
					  Case "文娥"   ' rs.movenext
					   Case "徐华"   ' rs.movenext
					    Case "禤葵英"   ' rs.movenext	
                         Case "杨广镇"   ' rs.movenext
                           Case "郑焕姬"   ' rs.movenext
                             Case "周凤萍"   ' rs.movenext
                               Case "周焕华"   ' rs.movenext  
	                             Case "周锦霞"   ' rs.movenext 
                                    Case "周永培"   ' rs.movenext 
  		 
                       Case Else        response.write "<option>" & rs("user") & "</option>"
           End Select

	else
                               response.write "<option>" & rs("user") & "</option>"
							   
	end if				   
							   
	
		 
		 
		 
		 
 '           response.write "<option>" & rs("user") & "</option>"
		

		
			
			
			
            rs.movenext
         wend
         rs.close
         set rs=nothing
         conn.close
         set conn=nothing
     %>
      </select></td>
    </tr>
    
    <tr>
    <td>
      　</td>
    <td bgcolor="#339933">　</td>
    </tr>
    
    <tr>
    <td>
      <p><font color="#FFFFFF"><b>密&nbsp; 码：</b></font></td>  
    <td bgcolor="#339933"><input type="password" name="mm" size="20" maxlength="50"></td>
    </tr>
    <tr>
    <td>　</td>
    <td bgcolor="#339933">
    &nbsp;    
    </td>
    </tr>
    <tr>
    <td>
      <p align="left"><input type="submit" name="dl" value="登 录"></td>
    <td bgcolor="#339933">
    <table><tr>
        <td>
          <p><input type="button" name="gm" value="修改密码" onclick="javascript:location.href='changepwd.asp?usern='+document.my.us.value">&nbsp; 
          </p>
        </td>
        <td>
          <p align="right">&nbsp; <input type="reset"  name="re" value="重 置"></p>
        </td>
    </tr></table>    
    </td>
    </tr>
    </table>

	    <input type="radio"  value="V6" <%if Session("PrintChoice_body")="print_onworkbody" then %> checked <% end if %> name="R2" onclick="openwin_onwork()"></b><font color="#ffffff">在职员工</font><b> 
        <input type="radio" name="R2" <%if Session("PrintChoice_body")="print_allbody" then %> checked <% end if %>  value="V5" onclick="openwin_all()"></b><font color="#ffffff">全体员工</font>&nbsp;&nbsp;
  
	
	
	
	
    </form>
    </td>
    <td width="204" height="312" align="center">
    
      <p align="center">
    
      <applet archive="image/AnLake.jar" code="AnLake.class" width="204" height="312">         
        <param name="credits" value="Applet by Fabio Ciucci (www.anfyteam.com)">         
        <param name="res" value="1">         
        <param name="image" value="image/bg1.jpg">         
        <param name="wavspeed" value="30">         
        <param name="perspective" value="10">         
        <param name="farwaving" value="200">         
        <param name="wind" value="10">         
        <param name="windvarval" value="5">         
        <param name="windvarmin" value="5">         
        <param name="windvarmax" value="15">         
        <param name="halfreflect" value="NO">         
        <param name="overimg" value="NO">         
        <param name="overimgX" value="0">         
        <param name="overimgY" value="0">         
        <param name="textscroll" value="NO">         
        <param name="regcode" value="NO">         
        <param name="reglink" value="NO">         
        <param name="regnewframe" value="NO">         
        <param name="regframename" value="_blank">         
        <param name="statusmsg" value="AnLake applet">         
        <param name="memdelay" value="1000">         
        <param name="priority" value="6">         
        <param name="MinSYNC" value="2">
        Sorry, your browser doesn't support Java(tm).          
      </applet>
    
    
    </td>
  </tr>

  <tr>
    <td width="680" height="20" align="center" colspan="3">
    
      <table border="0" width="117%">
        <tr>
          <td width="100%"><img border="0" src="image/bg4.jpg" width="600" height="20" align="right"></td>
        </tr>
      </table>
    
    
    </td>
  </tr>
</table>
</table>
</body>
</html>