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
     <TH width="280"><a href="InStore_Print_All_Frame.asp?Sort1=ClientID"><font color="#ffffff">存货货主</font></a></TH> 
     <TH width="70"><a href="InStore_Print_All_Frame.asp?Sort1=BillNo"><font color="#ffffff">单 号</font></a></TH> 
     <TH width="70"><a href="InStore_Print_All_Frame.asp?Sort1=ContractID"><font color="#ffffff">合同号</font></a></TH> 
     <TH width="110"><a href="InStore_Print_All_Frame.asp?Sort1=GoodsName"><font color="#ffffff">品名</font></a></TH>                                                                          
     <TH width="110"><a href="InStore_Print_All_Frame.asp?Sort1=Specs"><font color="#ffffff">规格</font></a></TH>                                                                          
     <TH width="40"><a href="InStore_Print_All_Frame.asp?Sort1=AllNumber"><font color="#ffffff">捆数</font></a></TH>                                                                                        
     <TH width="60"><a href="InStore_Print_All_Frame.asp?Sort1=PackNumber"><font color="#ffffff">件数</font></a></TH>                                                                                        
     <TH width="120"><a href="InStore_Print_All_Frame.asp?Sort1=Weight"><font color="#ffffff">总重(吨)</font></a></TH>                                                                                        
     <TH width=110><a href="InStore_Print_All_Frame.asp?Sort1=Date1"><font color="#ffffff">日期</font></a></TH>
  </TR> 


<!-- #include file="selectdate.asp" -->

<%  

  if session("varDisplayZero") = "ON" then
  	  sql1 = "and date1 >= #" & StartDate & "#" & " and date1 <= #" & EndDate & "#" 
  else
  	  sql1 = "and InStore.RemainWeight > 0 and date1 >= #" & StartDate & "#" & " and date1 <= #" & EndDate & "#" 
  end if

  set rs=server.createobject("adodb.recordset")
  
  if session("varBillNo")<>"" then
	  sql ="Select ClientName,BillNo,ClientID,ContractID,GoodsName,Specs,AllNumber,PackNumber,Weight,Date1 " & _
    	   "from ClientInfo,InStore " & _
	       "where ClientInfo.ID=InStore.ClientID and RemainWeight > 0  and BillNo='"  & session("varBillNo") & "'"
  
  else
  		if session("varContractID")<>"" then
			  sql ="Select ClientName,BillNo,ClientID,ContractID,GoodsName,Specs,AllNumber,PackNumber,Weight,Date1 " & _
			       "from ClientInfo,InStore " & _
			       "where ClientInfo.ID=InStore.ClientID and RemainWeight > 0  and ContractID='" & session("varContractID") & "'" 
  		else
  		
			  sql ="Select ClientName,BillNo,ClientID,ContractID,GoodsName,Specs,AllNumber,PackNumber,Weight,Date1 " & _
			       "from ClientInfo,InStore " & _
			       "where ClientInfo.ID=InStore.ClientID   " & sql1 

			  if request("Sort1") ="" then
			  		session("SortKey")="ClientID"
			  		session("SortDirect")="asc"
			  	    sql =sql & "order by InStore.ClientID asc, BillNo asc, Date1 asc,time1 asc"
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

   end if

  
  
  rs.Open sql, conn,1,1
  
  while not rs.eof
      
       Response.Write "<TR>"
       
      Response.Write "<td bgcolor='#e4e8ef'><A HREF='Print_InStore_Table_Grid.asp?BillNo=" & rs("BillNo") & "&ClientID=" & rs("ClientID") & "&PrintDate=" & rs("Date1") & "'>" & _
					 "<p align='center'><img border='0' src='image/Print.gif' width='16' height='16'></p></A></TD>" & _				    
					 "<TD bgcolor='#ffffff'>" & rs("ClientName") & "</td>" & _
					 "<TD bgcolor='#e4e8ef' align='center'>" & rs("BillNo") & "</td>" & _
					 "<TD bgcolor='#ffffff' align='center'>" & rs("ContractID") & "</td>" & _
					 "<td bgcolor='#e4e8ef' align='center'>" & rs("GoodsName") & "</td>" & _
					 "<td bgcolor='#ffffff' align='center'>" & rs("Specs") & "</td>" & _
                     "<td bgcolor='#e4e8ef' align='right'>" & rs("AllNumber") & "</td>" & _
                     "<td bgcolor='#ffffff' align='right'>" & rs("PackNumber") & "</td>" & _
                     "<td bgcolor='#e4e8ef' align='right'>" & gs(cdbl(rs("Weight")),4) & "</td>" & _
                     "<td bgcolor='#ffffff' align='center'>" & rs("Date1") & "</td></TR>"

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