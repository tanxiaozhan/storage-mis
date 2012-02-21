<%
function jiami(strf)
   
   randomize
   tmp=""
   
   sz=(rnd * 25) + 97
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

function gs(dt,w)
   if w=2 then 
      dt1=cdbl(dt)+0.005
   else 
      dt1=cdbl(dt)+0.00005
   end if
   
   if dt1<1 and dt1>=0 then 
      tmp="0" & cstr(dt1)
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


%>

<%
  
  response.write gs(0.03398,2)
  response.write "<br>"
  
  response.write gs(0,4)
  response.write "<br>"
  
  response.write gs(cdbl(0.25),2)
  response.write "<br>"
    response.write gs(cdbl(0.25),2)
  response.write "<br>"
    response.write gs(cdbl(0.25),2)
  response.write "<br>"
    response.write gs(cdbl(0.25),2)
  response.write "<br>"
    response.write gs(cdbl(0.25),2)
  response.write "<br>"
    response.write gs(cdbl(0.25),2)
  response.write "<br>"
    response.write gs(cdbl(0.25),2)
  response.write "<br>"
    response.write gs(cdbl(0.25),2)
  response.write "<br>"
    response.write gs(cdbl(0.25),2)
  response.write "<br>"
  
    response.write gs(1111.1,4)
  response.write "<br>"
    
  response.write "<br><br><br>"
  cc="32.569"
 
  a="5566"
  b=jiami(a)
  response.write a & "加密后:" & b
  response.write "<br>"
  response.write b & "解密后" & jiemi(b)
  
  
  fromDate=#2/1/2000#
  toDate=#2/9/2000# 
  response.write "<br>" & DateDiff("d",fromDate,toDate)
  
  t1=#9:30:30#
  t2=#8:30:30#
  if t1<t2 then response.write "sdfsdf"
%>