<%
  if session("username")="" or session("userright")="" then 
     response.write "你还没有登录系统！请【</b></font><a href='index.asp'>登录</a>】"
     response.end
  end if
  'csng()
%>
<%
function jiami(strf)
   
   randomize
   tmp=""
   
   sz=(rnd * 25) + 90'65
   for i=1 to len(strf)
       tmp=tmp & chr((asc(mid(strf,i,1)) xor sz))
   next
   
   jiami=tmp & chr(sz)
end function


function jiemi(strf)
   
   sz=asc(right(strf,1))
   
   tmp=""
   
   for i=1 to len(strf)-1
       tmp=tmp & chr((asc(mid(strf,i,1)) xor sz))
   next
   
   jiemi=tmp
      
end function

%>
<%
    dim conn
    dim  connectstr
    set  conn=server.createobject("adodb.connection")
    'pathstr=split(server.mappath("chat.mdb"),"\")
    'pathstr(0) & "\" pathstr(1) & "data\" & "chat.mdb"
    connectstr="provider=microsoft.jet.oledb.4.0; jet oledb:database password=office;data source=" & server.mappath("data\SMIS.mdb")
    'Response.Write connectstr
    'Response.End 
    conn.Open connectstr 
 %>
 
 <%
 function gs(data,nlen)
	n=0.5
	
 	data=int(cdbl(data)*10^nlen+n)/10^nlen
 	if data=int(data) then
 		tempdata=cstr(data) + ".0000"
 	else
 		tempdata=cstr(data)
 	end if
 	
 	data1=split(trim(tempdata),".")
	if cdbl(data)<1 then 	
		data1(0)="0"
	end if
	ll=len(trim(data1(1)))
	for i = 1 to nlen - ll
		data1(1)=data1(1) + "0"
	
	next 	
		
 	gs=data1(0)  + "." + left(data1(1),4)
	


 end function
 
 %>
 
 
 
 <script language=vbscript>
 <!--
 
 function gs1(data,nlen)
		
	n=0.5
	
 	data=int(data*10^nlen+n)/10^nlen
 	if data=int(data) then
 		tempdata=cstr(data) + ".0000"
 	else
 		tempdata=cstr(data)
 	end if
 	
 	data1=split(tempdata,".")
	if data<1 then 	
		data1(0)="0"
	end if
	ll=len(trim(data1(1)))
	for i = 1 to nlen - ll
		data1(1)=data1(1) + "0"
	
	next 	
		
 	gs1=data1(0)  + "." + left(data1(1),4)
	


 end function
 
 -->
 </script>