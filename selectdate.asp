<%

  if isLeadYear( session("varStartYear") ) then
  		days(2)=29    '闰年
  else
  		days(2)=28    '平年
  end if
  
  if Session("varStartDay") > days( Session("varStartMonth") ) then
  		Session("varStartDay") = days(Session("varStartMonth"))
  end if
  
  
  if isLeadYear( session("varEndYear") ) then
  		days(2)=29    '闰年
  else
  		days(2)=28    '平年
  end if
  
  if Session("varEndDay") > days( Session("varEndMonth") ) then
  		Session("varEndDay") = days(Session("varEndMonth"))
  end if
  
  
  StartDate =CStr( Session("varStartYear") ) & "-" & CStr( Session("varStartMonth") ) & "-" & CStr( Session("varStartDay") )
  EndDate =CStr( Session("varEndYear") ) & "-" & CStr( Session("varEndMonth") ) & "-" & CStr( Session("varEndDay") )
  
%>