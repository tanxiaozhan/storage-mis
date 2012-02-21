<%
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
<title>修改用户密码</title>
</head>

<body bgcolor="#007E23" topmargin="0" leftmargin="0" background="image/bg5.jpg">
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
    <form name="my" action="changepwd1.asp">
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
            response.write "<option>" & rs("user") & "</option>"
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
      <p><font color="#FFFFFF"><b>旧密码：</b></font></td> 
    <td bgcolor="#339933"><input type="password" name="oldp" size="20" maxlength="50"></td>
    </tr>
    <tr>
    <td>
      <p><font color="#FFFFFF"><b>新密码：</b></font></td> 
    <td bgcolor="#339933"><input type="password" name="newp" size="20" maxlength="50"></td>
    </tr>
    <tr>
    <td>
      <p><font color="#FFFFFF"><b>确认密码：</b></font></td> 
    <td bgcolor="#339933"><input type="password" name="rnewp" size="20" maxlength="50"></td>
    </tr>
    <tr>
    <td>　</td>
    <td bgcolor="#339933">
    &nbsp;    
    </td>
    </tr>
    <tr>
    <td>
      <p align="left"><input type="submit" name="dl" value="修改密码"></td>
    <td bgcolor="#339933">
    <table><tr>
        <td>
          <p align="left"><input type="button" name="gm" value="返 回" onclick="javascript:history.back()">&nbsp; 
          </p>
        </td>
        <td>
          <p align="left"><input type="reset"  name="re" value="重 置"></p>
        </td>
    </tr></table>    
    </td>
    </tr>
    </table>
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