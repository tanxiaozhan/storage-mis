<!-- #include file="data\config.asp" -->
<html>

<head>
<STYLE type=text/css>A:link {
	COLOR: #000000; TEXT-DECORATION: none
}
A:active {
	COLOR: #000000; TEXT-DECORATION: none
}
A:visited {
	COLOR: #000000; TEXT-DECORATION: none
}
A:hover {
	COLOR: #4455aa; TEXT-DECORATION: underline
}
BODY {
	SCROLLBAR-FACE-COLOR: #dee3e7; FONT-SIZE: 12px; SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; SCROLLBAR-SHADOW-COLOR: #dee3e7; COLOR: #000000; SCROLLBAR-3DLIGHT-COLOR: #d1d7dc; SCROLLBAR-ARROW-COLOR: #006699; SCROLLBAR-TRACK-COLOR: #efefef; FONT-FAMILY: 宋体; SCROLLBAR-DARKSHADOW-COLOR: #98aab1
}
FONT {
	LINE-HEIGHT: normal
}
TD {
	FONT-SIZE: 12px; LINE-HEIGHT: 15px; FONT-FAMILY: 宋体
}
TH {
	FONT-WEIGHT: bold; FONT-SIZE: 12px; BACKGROUND-IMAGE: url('Skins/Default/css/default/bg1.gif'); COLOR: white; BACKGROUND-COLOR: #4455aa
}
TD.TableTitle2 {
	BACKGROUND-COLOR: #e4e8ef
}
TD.TableBody1 {
	LINE-HEIGHT: normal; BACKGROUND-COLOR: #ffffff
}
TD.TableBody2 {
	LINE-HEIGHT: normal; BACKGROUND-COLOR: #e4e8ef
}
TD.TopDarkNav {
	BACKGROUND-IMAGE: url('Skins/Default/css/default/topbg.gif')
}
TD.TopLighNav {
	BACKGROUND-IMAGE: url('Skins/Default/css/default/bottombg.gif')
}
TD.TopLighNav1 {
	BACKGROUND-IMAGE: url('Skins/Default/css/default/tabs_m_tile.gif')
}
TD.TopLighNav2 {
	BACKGROUND-COLOR: #ffffff
}
.tableBorder1 {
	BORDER-RIGHT: 1px; BORDER-TOP: 1px; BORDER-LEFT: 1px; WIDTH: 98%; BORDER-BOTTOM: 1px; BACKGROUND-COLOR: #6595d6
}
.tableBorder2 {
	BORDER-RIGHT: #dedede 1px solid; BORDER-TOP: #dedede 1px solid; BORDER-LEFT: #dedede 1px solid; WIDTH: 98%; BORDER-BOTTOM: #dedede 1px solid; BACKGROUND-COLOR: #efefef
}
#TableTitleLink A:link {
	COLOR: #ffffff; TEXT-DECORATION: none
}
#TableTitleLink A:visited {
	COLOR: #ffffff; TEXT-DECORATION: none
}
#TableTitleLink A:active {
	COLOR: #ffffff; TEXT-DECORATION: none
}
#TableTitleLink A:hover {
	COLOR: #ffffff; TEXT-DECORATION: underline
}
INPUT {
	FONT-SIZE: 12px; COLOR: #000000; LINE-HEIGHT: 15px; FONT-FAMILY: Tahoma,Verdana,"宋体"
}
SELECT {
	FONT-SIZE: 12px; COLOR: #000000; LINE-HEIGHT: 15px; FONT-FAMILY: Tahoma,Verdana,"宋体"
}
TEXTAREA {
	FONT-SIZE: 12px; COLOR: #000000; LINE-HEIGHT: 15px; FONT-FAMILY: Tahoma,Verdana,"宋体"
}
OPTION {
	FONT-SIZE: 12px; COLOR: #000000; LINE-HEIGHT: 15px; FONT-FAMILY: Tahoma,Verdana,"宋体"
}
.normalTextSmall {
	FONT-SIZE: 11px; COLOR: #000000; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif
}
.menuskin {
	border:1px solid #666666; BACKGROUND-IMAGE: url('Skins/Default/dvmenubg3.gif'); VISIBILITY: hidden; BACKGROUND-REPEAT: repeat-y; POSITION: absolute; BACKGROUND-COLOR: #efefef; font-style:normal; font-variant:normal; font-weight:normal; font-size:12px; font-family:Verdana
}
.menuskin A {
	PADDING-RIGHT: 10px; PADDING-LEFT: 25px; BEHAVIOR: url(inc/noline.htc); COLOR: black; TEXT-DECORATION: none
}
#mouseoverstyle {
	BORDER-RIGHT: #597db5 1px solid; PADDING-RIGHT: 0px; BORDER-TOP: #597db5 1px solid; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; MARGIN: 2px; BORDER-LEFT: #597db5 1px solid; PADDING-TOP: 0px; BORDER-BOTTOM: #597db5 1px solid; BACKGROUND-COLOR: #c9d5e7
}
#mouseoverstyle A {
	COLOR: black
}
.menuitems {
	PADDING-RIGHT: 1px; PADDING-LEFT: 1px; PADDING-BOTTOM: 1px; MARGIN: 2px; WORD-BREAK: keep-all; PADDING-TOP: 1px
}
A.navlink:link {
	COLOR: #000000; TEXT-DECORATION: none
}
A.navlink:visited {
	COLOR: #000000; TEXT-DECORATION: none
}
A.navlink:hover {
	COLOR: #003399; TEXT-DECORATION: none
}
.BrightClass {
	BACKGROUND-COLOR: #d7d7d7
}
DIV.quote {
	BORDER-RIGHT: #cccccc 1px solid; PADDING-RIGHT: 5px; BORDER-TOP: #cccccc 1px solid; PADDING-LEFT: 5px; BACKGROUND: #f3f3f3; PADDING-BOTTOM: 5px; MARGIN: 5px 20px; BORDER-LEFT: #cccccc 1px solid; LINE-HEIGHT: normal; PADDING-TOP: 5px; BORDER-BOTTOM: #cccccc 1px solid
}
DIV.HtmlCode {
	BORDER-RIGHT: #cccccc 1px solid; PADDING-RIGHT: 5px; BORDER-TOP: #cccccc 1px solid; PADDING-LEFT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 14px; BACKGROUND: #fdfddf; PADDING-BOTTOM: 5px; MARGIN: 5px 20px; BORDER-LEFT: #cccccc 1px solid; LINE-HEIGHT: normal; PADDING-TOP: 5px; BORDER-BOTTOM: #cccccc 1px solid; FONT-STYLE: oblique; FONT-FAMILY: Tahoma
}
</STYLE>
</head>

<body>
<%
  
  sql="select * from UnitInfo"
  
  Set rs =server.createobject("adodb.recordset")
  
  rs.Open sql,conn,3,3
 %>
  <TABLE class=tableborder1 cellSpacing=1 cellPadding=1 width="400">
  <TR> 
     <TH width="300"><font color="#ffffff">计量单位名称</font></TH>                                                                                    
     <TH width=30>&nbsp;&nbsp;&nbsp;&nbsp;</TH>
  </TR> 

<%
  'set objRS1=server.createobject("adodb.recordset")    
  while not rs.eof
      
       Response.Write "<TR>"
      'Response.Write "<TD bgcolor="e4e8ef"><IMG SRC=" & Fig(I - 1) & "></TD>"
      '显示留言的作者、主题及时间，至于内容则设计为连接至ShowNews.asp的超链接
      'strSQL = "SELECT * FROM ClientName WHERE Id=" & rs("id") 
      'objRS1.Open strSQL,conn,3,3 
      
      Response.Write "<TD bgcolor='#e4e8ef'>" & rs("Unitname") & "</td><td bgcolor='#e4e8ef'><A HREF='deljl.asp?jln=" & rs("Unitname") & "'>删除</A></TD></TR>"
      
      'objRS1.close                      
      rs.MoveNext
      
       
  wend
     
  response.write "</table>"
  
  'objRS1.close
  'set objRS1=nothing
  
  rs.close
  set rs=nothing
  
  conn.close
  set conn=nothing 

%>
</table>
</body>
</html>