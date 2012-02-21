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
  <TABLE class=tableborder1 cellSpacing=1 cellPadding=1>
  <TR> 
     <TH width=60>出仓</TH>
     <TH width="70"><a href="OutStore_Browse_Frame.asp?Sort1=BillNo"><font color="#ffffff">进仓单号</font></a></TH> 
     <TH width="90"><a href="OutStore_Browse_Frame.asp?Sort1=ContractID"><font color="#ffffff">合同号</font></a></TH> 
     <TH width="250"><a href="OutStore_Browse_Frame.asp?Sort1=GoodsName"><font color="#ffffff">品 名</font></a></TH>                                                                          
     <TH width="110"><a href="OutStore_Browse_Frame.asp?Sort1=Specs"><font color="#ffffff">规 格</font></a></TH>                                                                          
     <TH width="60"><a href="OutStore_Browse_Frame.asp?Sort1=AllNumber"><font color="#ffffff">捆 数</font></a></TH>                                                                                          
     <TH width="60"><a href="OutStore_Browse_Frame.asp?Sort1=RemainNumber"><font color="#ffffff">结余捆数</font></a></TH>                                                                                          
     <TH width="60"><a href="OutStore_Browse_Frame.asp?Sort1=PackNumber"><font color="#ffffff">件 数</font></a></TH>    
     <TH width="60"><a href="OutStore_Browse_Frame.asp?Sort1=RemainPack"><font color="#ffffff">结余件数</font></a></TH>    
     <TH width="60"><a href="OutStore_Browse_Frame.asp?Sort1=Weight"><font color="#ffffff">总重(吨)</font></a></TH>                                                                                          
     <TH width="100"><a href="OutStore_Browse_Frame.asp?Sort1=RemainWeight"><font color="#ffffff">结余总重(吨)</font></a></TH> 
     <TH width="50"><a href="OutStore_Browse_Frame.asp?Sort1=Place1"><font color="#ffffff">仓 位</font></a></TH> 
	   <TH width="60"><a href="OutStore_Browse_All_Frame.asp?Sort1=Place2"><font color="#ffffff">库  </font></a></TH> 
     <TH width=110><a href="OutStore_Browse_Frame.asp?Sort1=date1"><font color="#ffffff">日 期</font></a></TH>
  </TR> 

<%
  set rs=server.createobject("adodb.recordset")
  
  
  if Session("varClientID") = "-" or Session("varClientID") = "" then   '未选择客户前,不显示, 停止执行  
		response.end

  else
  		
  		sql ="Select * from InStore where RemainWeight> 0.01 AND ClientID = " & Session("varClientID")  & " "   '显示选择客户的仓储
  
	   if request("Sort1") ="" then
  			session("SortKey")="date1"
  			session("SortDirect")="asc"
	  	    sql =sql & " order by date1 asc,time1 asc " 
  	   else
  
			if request("Sort1")= session("SortKey") then
	
  				if session("SortDirect") = "asc" then
  					session("SortDirect") = "desc"
  				else
  			    	session("SortDirect") ="asc"
			    end if


 	  		else

  				session("SortKey") = request("Sort1")

  		    end if

		    sql=sql & "order by " & session("SortKey") & " " & session("SortDirect")

  	    end if

  end if

  
  
  rs.Open sql, conn,1,1
  
  while not rs.eof
      
       Response.Write "<TR>"
      
      Response.Write "<td bgcolor='#e4e8ef'><A HREF='OutStore_Fill.asp?RecNo=" & rs("RecNo") & "&ClientID=" & rs("ClientID") & "'>" & _
                     "<p align='center'><img border='0' src='image/OutStore.gif' width='16' height='16'></p></A></TD>" & _  
                     "<TD bgcolor='#ffffff' align='center'>" & rs("BillNo") & "</td>" & _ 
                     "<td bgcolor='#e4e8ef' align='center'>" & rs("ContractID") & "</td>"  
					 
					 
%>					 
	 
<%					 
	'	if    instr(rs("GoodsName"),"高光")> 1 and session("username")="区福全" then
         	if    instr(rs("GoodsName"),"白底白")> 1  then
		 response.write   "<td bgcolor='#CCFFFF' align='center'><font color='#3300FF'>" & rs("GoodsName") & "</font></td>" 
        else
   
             	if instr(rs("GoodsName"),"亚光")> 1 then
                   response.write   "<td bgcolor='#ffffff' align='center'><font color='#008B8B'>" & rs("GoodsName") & "</font></td>" 
		else
		
			    if instr(rs("GoodsName"),"双铜")> 1 or instr(rs("GoodsName"),"铜版")>1 then
                   response.write   "<td bgcolor='#ffffff' align='center'><font color='#B23AEE'>" & rs("GoodsName") & "</font></td>" 
        else
            	if instr(rs("GoodsName"),"双胶")> 1 or instr(rs("GoodsName"),"胶版印刷")>1  then		
                   response.write   "<td bgcolor='#ffffff' align='center'><font color='#008B00'>" & rs("GoodsName") & "</font></td>" 
		else
		        if instr(rs("GoodsName"),"灰底白")> 1 or instr(rs("GoodsName"),"白板")> 1 then		
                   response.write   "<td bgcolor='#6B6B6B' align='center'><font color='#ffffff'>" & rs("GoodsName") & "</font></td>" 
		else
				        if instr(rs("GoodsName"),"白卡")> 1 then		
                   response.write   "<td bgcolor='#ffffff' align='center'><font color='#FF6600'>" & rs("GoodsName") & "</font></td>" 
		else
                   response.write   "<td bgcolor='#ffffff' align='center'>" & rs("GoodsName") & "</td>" 
        end if
	    end if
	    end if
	    end if
		end if
		end if
		
	'	灰底白
		
%>

<%					 
					 
   response.write    "<td bgcolor='#e4e8ef' align='center'>" & rs("Specs") & "</td>" & _
                     "<td bgcolor='#ffffff' align='right'>" & rs("AllNumber") & "</td>" & _
                     "<td bgcolor='#e4e8ef' align='right'>" & rs("RemainNumber") & "</td>" & _
                     "<td bgcolor='#ffffff' align='right'>" & rs("PackNumber") & "</TD>" & _
                     "<td bgcolor='#e4e8ef' align='right'>" & rs("RemainPack") & "</td>" & _
                     "<td bgcolor='#ffffff' align='right'>" & gs(cdbl(rs("Weight")),4) & "</td>" & _
                     "<td bgcolor='#e4e8ef' align='right'>" & gs(cdbl(rs("RemainWeight")),4) & _
                     "<td bgcolor='#ffffff' align='center'>" & rs("Place1") & "</td>" & _
					  "<td bgcolor='#ffffff' align='center'>" & rs("Place2") & "</td>" & _
                     "<td bgcolor='#e4e8ef' align='center'>" & rs("Date1") & "</td></TR>"

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