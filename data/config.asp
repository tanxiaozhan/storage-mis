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
function gs(dt,w)
	if cdbl(dt)<0.00001 then
		dt=0
	end if
	
   if w=2 then 
      dt1=cdbl(dt)+0.005
   else 
      dt1=cdbl(dt)+0.00005
   end if
   
   if dt1<1 and dt1>=0 then 
      tmp= cstr(dt1)
   else
      if dt1>=1 then
        tmp=cstr(dt1)
      else
        tmp="0.0"
      end if  
   end if   
   
   tmp1=split(tmp,".")
   
   on error resume next
   
   gs=tmp1(0) & "." & left(tmp1(1),w)
   
end function


function show1(tmpstr)
  if not isnull(tmpstr) then
    'tmpp=""
     
    'k=0
     
    ' if len(tmpstr)>10 then
    '    kk=10
    ' else
    '    kk=len(tmpstr)
    ' end if   
     
    ' for i=1 to kk    
    '     tmpp=tmpp & mid(tmpstr,i,1)
    '     k=k+1
    '     if k=5 and i<len(tmpstr) then 
    '        tmpp=tmpp & "<br>"
    '        k=0
    '     end if
    ' next
    ' show1=tmpp
     show1=left(trim(tmpstr),3)
  else
     show1=""   
  end if
  
end function
 
 %>
 
 
 <%
 	'十二个月份的天数
	dim days(12)
	days(1)=31
	days(2)=29 '平年另外作处理
	days(3)=31
	days(4)=30
	days(5)=31
	days(6)=30
	days(7)=31
	days(8)=31
	days(9)=30
	days(10)=31
	days(11)=30
	days(12)=31

'判断闰年函数
function isLeadYear(testyear)    

 if ( testYear mod 4) = 0 then
  		if (testYear mod 100) > 0 or (testYear mod 400)=0 then
  			 isLeadYear =true     '闰年
  			 exit function
  		end if
  end if
  
  isLeadYear=false     '平年
  
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