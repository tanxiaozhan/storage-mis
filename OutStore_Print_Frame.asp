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
	FONT-WEIGHT: bold; FONT-SIZE: 12px; BACKGROUND-IMAGE: url(Skins/Default/css/default/bg1.gif); COLOR: white; BACKGROUND-COLOR: #4455aa
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
	BACKGROUND-IMAGE: url(Skins/Default/css/default/topbg.gif)
}
TD.TopLighNav {
	BACKGROUND-IMAGE: url(Skins/Default/css/default/bottombg.gif)
}
TD.TopLighNav1 {
	BACKGROUND-IMAGE: url(Skins/Default/css/default/tabs_m_tile.gif)
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
	BORDER-RIGHT: #666666 1px solid; BORDER-TOP: #666666 1px solid; BACKGROUND-IMAGE: url(Skins/Default/dvmenubg3.gif); VISIBILITY: hidden; FONT: 12px Verdana; BORDER-LEFT: #666666 1px solid; BORDER-BOTTOM: #666666 1px solid; BACKGROUND-REPEAT: repeat-y; POSITION: absolute; BACKGROUND-COLOR: #efefef
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
     <TH width=40>打印</TH>
     <TH width="100"><a href="OutStore_Print_Frame.asp?Sort1=BillNo"><font color="#ffffff">出仓单号</font></a></TH> 
     <TH width="110"><a href="OutStore_Print_Frame.asp?Sort1=PickUpID"><font color="#ffffff">提单号</font></a></TH>                                                                          
     <TH width="190"><a href="OutStore_Print_Frame.asp?Sort1=Crop"><font color="#ffffff">提货单位</font></a></TH>                                                                          
     <TH width="210"><a href="OutStore_Print_Frame.asp?Sort1=GoodsName"><font color="#ffffff">品 名</font></a></TH>                                                                          
     <TH width="110"><a href="OutStore_Print_Frame.asp?Sort1=Specs"><font color="#ffffff">规 格</font></a></TH>                                                                          
     <TH width="60"><a href="OutStore_Print_Frame.asp?Sort1=Number1"><font color="#ffffff">捆 数</font></a></TH>                                                                                          
     <TH width="80"><a href="OutStore_Print_Frame.asp?Sort1=Pack1"><font color="#ffffff">件　数</font></a></TH>    
     <TH width="100"><a href="OutStore_Print_Frame.asp?Sort1=Weight"><font color="#ffffff">总重(吨)</font></a></TH>                                                                                          
     <TH width="50"><a href="OutStore_Print_Frame.asp?Sort1=Place1"><font color="#ffffff">备注</font></a></TH> 
     <TH width=110><a href="OutStore_Print_Frame.asp?Sort1=Date2"><font color="#ffffff">日　期</font></a></TH>
  </TR> 

<!-- #include file="selectdate.asp" -->

<%  
  
  sql1 = "and date2 >= #" & StartDate & "#" & " and date2 <= #" & EndDate & "#" 


  set rs=server.createobject("adodb.recordset")
  
  
  if Session("varClientID") = "-" or Session("varClientID") = "" then   '未选择客户前,不显示, 停止执行代码  
	  Response.End
  end if
  
  sql ="Select * from OutStore where ClientID = " & Session("varClientID")  & " " & sql1    '显示选择客户的出仓记录
 
 if request("Sort1") ="" then
  		session("SortKey")="Date2"
  		session("SortDirect")="asc"
  	    sql =sql & "order by date2 asc,time1 asc "
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
  
  rs.Open sql, conn,1,1
  
  while not rs.eof
      
      Response.Write "<td bgcolor='#e4e8ef'><A HREF='Print_OutStore_Table_Grid.asp?ClientID=" & rs("ClientID") & "&BillNo=" & rs("BillNo") & "&PrintDate=" & rs("Date2") &  "'>" & _
                     "<p align='center'><img border='0' src='image/Print.gif' width='16' height='16'></p></A></TD>"  & _
                     "<TD bgcolor='#ffffff' align='center'>" & rs("BillNo") & "</td>" & _
                     "<td bgcolor='#e4e8ef' align='center'>" & rs("PickupID") & "</td>" & _
                     "<td bgcolor='#ffffff' align='center'>" & rs("Crop") & "</td>" & _
                     "<td bgcolor='#e4e8ef' align='center'>" & rs("GoodsName") & "</td>" & _
                     "<td bgcolor='#ffffff' align='center'>" & rs("Specs") & "</td>" & _
                     "<td bgcolor='#e4e8ef' align='right'>" & rs("Number1") & "</td>" & _
                     "<td bgcolor='#ffffff' align='right'>" & rs("Pack1") & "</TD>" & _
                     "<td bgcolor='#e4e8ef'align='right'>" & gs(cdbl(rs("Weight")),4) & "</td>" & _
                     "<td bgcolor='#ffffff'>" & rs("Remark") & "</td>" & _
                     "<td bgcolor='#e4e8ef'>" & rs("Date2") & "</td></TR>"

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