<!-- #include file="data\config.asp" -->
 
 <SCRIPT LANGUAGE ="VBSCRIPT">
  <!-- 
    Sub ValidateData()
        set ThisForm = Document.InStoreForm 
     
        If trim(ThisForm.tGoodsName.Value) = Empty Then
            Window.Alert "品名不能空”！"
            Exit Sub
        End If

  	    if instr(1,trim(ThisForm.tGoodsName.Value),"'")>0 or instr(1,trim(ThisForm.tGoodsName.Value),chr(34))>0 then
        	Window.Alert "“品名”中不能含有 ' 和 '' ！"
          	ThisForm.tGoodsName.focus()
          	Exit Sub
        end if

        If trim(ThisForm.tSpecs.Value) = Empty Then
            Window.Alert "规格不能空”！"
            Exit Sub
        End If

  	    if instr(1,trim(ThisForm.tSpecs.Value),"'")>0 or instr(1,trim(ThisForm.tSpecs.Value),chr(34))>0 then
        	Window.Alert "“规格”中不能含有 ' 和 '' ！"
          	ThisForm.tSpecs.focus()
          	Exit Sub
        end if

        If trim(ThisForm.tNumber.Value) = Empty Then
            Window.Alert "捆数不能为空”！"
            Exit Sub
        End If
 
  	    if instr(1,trim(ThisForm.tNumber.Value),"'")>0 or instr(1,trim(ThisForm.tNumber.Value),chr(34))>0 then
        	Window.Alert "“捆数”中不能含有 ' 和 '' ！"
          	ThisForm.tNumber.focus()
          	Exit Sub
        end if
        
		if trim(ThisForm.tAllPack.value) = Empty then
            Window.Alert "总件数不能为空”！"
            Exit Sub
        End If

  	    if instr(1,trim(ThisForm.tAllPack.Value),"'")>0 or instr(1,trim(ThisForm.tAllPack.Value),chr(34))>0 then
        	Window.Alert "“总件数”中不能含有 ' 和 '' ！"
          	ThisForm.tAllPack.focus()
          	Exit Sub
        end if

  	    if instr(1,trim(ThisForm.tPlace1.Value),"'")>0 or instr(1,trim(ThisForm.tPlace1.Value),chr(34))>0 then
        	Window.Alert "“仓位1”中不能含有 ' 和 '' ！"
          	ThisForm.tPlace1.focus()
          	Exit Sub
        end if
  	    if instr(1,trim(ThisForm.tPlace2.Value),"'")>0 or instr(1,trim(ThisForm.tPlace2.Value),chr(34))>0 then
        	Window.Alert "“仓位2”中不能含有 ' 和 '' ！"
          	ThisForm.tPlace2.focus()
          	Exit Sub
        end if
  	    if instr(1,trim(ThisForm.tPlace3.Value),"'")>0 or instr(1,trim(ThisForm.tPlace3.Value),chr(34))>0 then
        	Window.Alert "“仓位3”中不能含有 ' 和 '' ！"
          	ThisForm.tPlace3.focus()
          	Exit Sub
        end if
  	    if instr(1,trim(ThisForm.tremark.Value),"'")>0 or instr(1,trim(ThisForm.tremark.Value),chr(34))>0 then
        	Window.Alert "“备注”中不能含有 ' 和 '' ！"
          	ThisForm.tremark.focus()
          	Exit Sub
        end if
        
  	    if instr(1,trim(ThisForm.tdeliver.Value),"'")>0 or instr(1,trim(ThisForm.tdeliver.Value),chr(34))>0 then
        	Window.Alert "“送货单位”中不能含有 ' 和 '' ！"
          	ThisForm.tdeliver.focus()
          	Exit Sub
        end if
  	    if instr(1,trim(ThisForm.tconsignee.Value),"'")>0 or instr(1,trim(ThisForm.tconsignee.Value),chr(34))>0 then
        	Window.Alert "“收货人”中不能含有 ' 和 '' ！"
          	ThisForm.tconsignee.focus()
          	Exit Sub
        end if
  	    if instr(1,trim(ThisForm.tdriver.Value),"'")>0 or instr(1,trim(ThisForm.tdriver.Value),chr(34))>0 then
        	Window.Alert "“司机”中不能含有 ' 和 '' ！"
          	ThisForm.tdriver.focus()
          	Exit Sub
        end if
  	    if instr(1,trim(ThisForm.tNumberPlate.Value),"'")>0 or instr(1,trim(ThisForm.tNumberPlate.Value),chr(34))>0 then
        	Window.Alert "“车号”中不能含有 ' 和 '' ！"
          	ThisForm.tNumberPlate.focus()
          	Exit Sub
        end if
		
        If trim(ThisForm.tPackWeight.Value) = Empty Then
            Window.Alert "包重不能为空”！"
            Exit Sub
        End If

        If trim(ThisForm.tWeight.Value) = Empty Then
            Window.Alert "总重不能为空”！"
            Exit Sub
        End If

        If trim(ThisForm.tGross.Value) = Empty Then
            Window.Alert "毛重不能为空”！"
            Exit Sub
        End If

        If trim(ThisForm.tNet.Value) = Empty Then
            Window.Alert "净重不能为空”！"
            Exit Sub
        End If

    	if not IsNumbers(ThisForm.tAllPack.value) then
    		Window.Alert "总包数必须为数字!"
    		Exit Sub
    	End If

    	if not IsNumbers(ThisForm.tGross.value) then
    		Window.Alert "毛重必须为数字!"
    		Exit Sub
    	End If

    	if not IsNumbers(ThisForm.tNet.value) then
    		Window.Alert "净重必须为数字!"
    		Exit Sub
    	End If

		if cdbl(ThisForm.tGross.value) < cdbl(ThisForm.tNet.value) then
			Window.Alert "毛重应大于或等于净重"
			Exit Sub
		End If
		
		'OnPackChange    '计算总包数

		'CountWeight     '计算总重量
		
		ThisForm.tAllPack.disabled=false
		
		ThisForm.Submit

    End Sub
    

   
    Sub OnR1V1()		'总包数Disabled

    	Document.InStoreForm.tAllPack.disabled=true
    	Document.InStoreForm.tPackNumber.disabled = false
    	

    End Sub
    
	Sub OnR1V2()

    	Document.InStoreForm.tPackNumber.disabled = true
    	Document.InStoreForm.tAllPack.disabled = False
		
	End Sub    

    Sub OnPackChange()    '计算总包数
    	set Form1=Document.InStoreForm
    	If not IsNumbers(Form1.tPackNumber.value) then
    		Window.Alert "包数必须为数字!"
    		Exit Sub
    	End IF
    	if (InStoreForm.tPackNumber.value<> Empty) and (InStoreForm.tNumber.value <>Empty) then
    		InStoreForm.tAllPack.value = InStoreForm.tNumber.value * InStoreForm.tPackNumber.value
    	end if
    	    
    End Sub
    
    Sub CountWeight()      '计算总重量
    	set Form1=Document.InStoreForm
    	if not IsNumbers(Form1.tPackWeight.value) then
    		Window.Alert "包重必须为数字!"
    		Exit Sub
    	end if
		if trim(Form1.tPackWeight.value) = Empty then
			' Window.Alert "包重不能为空!"
			Exit Sub
		end if
		if trim(Form1.tAllPack.value)=Empty then
			Window.Alert "总包数不能为空!"
			Exit Sub
		end if
	   	Form1.tWeight.value = gs1(Form1.tPackWeight.value * Form1.tAllPack.value /1000,4)
    End Sub
    

    Sub CheckNumber()   '检查件数是否为数字
    	if not IsNumbers(InStoreForm.tNumber.value) then
    		Window.Alert "件数必须为数字!"
    		Exit Sub
    	End If
    	if (InStoreForm.tPackNumber.value<> Empty) and (InStoreForm.tNumber.value <>Empty) then
    		InStoreForm.tAllPack.value = InStoreForm.tNumber.value * InStoreForm.tPackNumber.value
    	end if
    End Sub
    

    Sub CheckAllPack()   '检查总包数是否为数字
    	if not IsNumbers(InStoreForm.tAllPack.value) then
    		Window.Alert "总包数必须为数字!"
    		Exit Sub
    	End If
    	if instoreform.tPackNumber.value = empty then
    		instoreform.tPackNumber.value=gs1( instoreform.tAllPack.value / instoreform.tNumber.value , 2 )
    	end if
    	
    	if instoreform.tpacknumber.value<> empty then
    		instoreform.tnumber.value=gs1(instoreform.tallpack.value / instoreform.tpacknumber.value,4)
	    end if
    	
    End Sub
    
    
    Sub CheckWeight()   '检查总重是否为数字
    	if not IsNumbers(InStoreForm.tWeight.value) then
    		Window.Alert "总重必须为数字!"
    		Exit Sub
    	End If
    	if (InStoreForm.tWeight.value<>Empty) and (InStoreForm.tAllPack.value<>Empty) then
    		InStoreForm.tPackWeight.value = gs1( InStoreForm.tWeight.value * 1000 / InStoreForm.tAllPack.value,4)
    		InStoreForm.tNet.value = InStoreForm.tPackWeight.value
    	end if
    	
    End Sub



    '检测是否为数字
    Function IsNumbers(str) 
    	for i=1 to len(str)
    		ch = mid(str,i,1) 
    		if not ( (ch>="0" and ch<="9") or ch="." ) then
    			IsNumbers = false
    			exit function
    		end if
    		
    	next
    	
    		IsNumbers= true
    		
    End Function

	Sub OnDeleteClick()
		dim choice	
		Choice = msgbox ("删除后将无法恢复，确认删除该记录吗?",vbYESNO,"删除记录")
		if Choice = vbNO then
			Exit Sub
		end if
		window.location.href="InStore_Add.asp?Operate=Edit&ButtonName=Delete&RecNo=" & "<%=Request("RecNo")%>"
	
	End Sub


  -->
  </SCRIPT>


<HTML>
<HEAD>
<TITLE>进仓管理</TITLE>
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<STYLE type=text/css>
<!--
A:link {text-decoration:none; color:"#0000ff"}
A:visited {text-decoration:none}
A:hover {text-decoration:underline; color:"#FF0033"}

body { font-family:"Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 9pt}

p {  font-family:"Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 9pt}

td {  font-family: "Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 9pt}
-->

</STYLE>
</HEAD>


<body>

		<%    Set rs =server.createobject("adodb.recordset")
			
		    sql = "select * from InStore where RecNo = "  & request("RecNo")
			rs.open sql ,conn,1,1
			if rs.eof then
				Response.Redirect "InStore_Browse_Frame.asp"
				response.end
			end if
			if rs("operater") = Session("UserName") then     '要修改的记录是否为当前操作员录入的
				OperaterMatch=1
			else
				OperaterMatch=0
			end if
			tempYear=year(rs("date1"))
			tempMonth=month(rs("date1"))
			tempDay=day(rs("date1"))
			
			

		%>
		

<TABLE border=0 cellPadding=0 cellSpacing=0 width="591">
  <TR>
    <TD width="589" valign="top" height="57"> 
      <form name="InStoreForm" method="POST" action="InStore_Add.asp?Operate=Edit&ButtonName=Save">
        <hr color="#AFE0FA">
        
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><b><font size="3" face="宋体">　　单　号                                           
        <input type="text" name="tBillNo" size="12" value="<% =rs("BillNo") %>"><font color="#FF0000">*</font>　日　期                                              
        <select size="1" name="DYear">
        <% for i=tempYear-5 to tempYear+5 %>  
        		<option <% if i=tempYear then %> selected <% end if %> ><%=i%></option>
        <% next %>
                	
        </select>年<select size="1" name="DMonth">  
        <% for i=1 to 12   
        		if i<10 then
        			tempvalue="0" + trim(CStr(i))
        		else
        			tempvalue=trim(CStr(i))
        		end if
        		
        %>
        		<option value="<% = tempvalue %>" <% if i=tempMonth then %> selected <% end if %> ><% =i %></option>
        <% next %>
        
        </select>月<select size="1" name="DDay">  
        <% for i=1 to 31%>  
               <option <% if i=tempDay then %> selected <% end if %> ><% =i %></option>
        <% next %>
               </select>日</font></b></p>  

        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><b><font size="3" face="宋体">
        　　货　主 <select size="1" name="D1" tabindex="1" >                                                                                                                                                                    
		
		<%	  
			rs.close
			sql = "select * from ClientInfo  " 
			rs.open sql,conn,1,1
			if rs.eof then
				response.write "客户数据错误！"
				Response.End
			end if
			do while not rs.eof
		%>			
	       		 <OPTION value = "<% = rs("ID") %>" <%if rs("ID") = clng(request("ClientID")) then %> selected <%end if%> > <% = rs("ClientName") %>  </Option>
		<% 
				rs.movenext
			loop   
			rs.close
		    sql = "select * from InStore where RecNo = "  & request("RecNo")
			rs.open sql ,conn,1,1
			
			
		%>
        
        </select><font color="#FF0000">*</font></font></b>  
        
        </p>

        
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><b><font size="3" face="宋体">　　合同号                                           
        <input type="text" name="tContractID" size="20" maxlength="40" tabindex="2" value='<% =rs("ContractID") %>'></font>
        
        </p>
        
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="宋体">　　品　名                                           
        <input type="text" name="tGoodsName" size="24" maxlength="40" tabindex="3" value=<% = rs("GoodsName") %>><font color="#FF0000">*</font>　　&nbsp;&nbsp;&nbsp;&nbsp;                                           
        规　格 <input type="text" name="tSpecs" size="19" tabindex="4" value=<% =rs("Specs") %> maxlength="40"><font color="#FF0000">*</font></font></p>                                           
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="宋体">　　包　装                                           
        <input type="text" name="tPack" size="18" maxlength="40" tabindex="5" value=<% = rs("Pack") %>>　　　</font></p>
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="宋体">　　捆　数                                           
        <input type="text" name="tNumber" size="9" onfocusout="CheckNumber()" maxlength="40" tabindex="6" value=<% = rs("AllNumber") %>><font color="#FF0000">*</font> 
        <input type="radio" value="V1" checked name="R1" onclick="OnR1V1()" tabindex="7">每捆件数<input type="text" name="tpackNumber" maxlength="40" size="10" onfocusout="OnPackChange()" tabindex="8" <% if rs("AllNumber")>0 then %> value = <% = gs(cdbl(rs("PackNumber") / rs("AllNumber")),2) %> <% end if %> ><font color="#FF0000">*</font>                                           
        <input type="radio" name="R1" value="V2" onClick="OnR1V2()" tabindex="9" maxlength="40">总件数<input type="text" name="tAllPack" disabled onfocusout="CheckAllPack()" size="10" tabindex="10" value=<% = rs("PackNumber") %> ><font color="#FF0000">*</font></font></p>                                                                                                          
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="宋体">　每件重量                                           
        <input type="text" name="tPackWeight" size="14" onfocusout="CountWeight()" tabindex="11" value=<% = rs("PackWeight") %> >公斤<font color="#FF0000">*</font>　&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                                           
        总　重                                               
        <input type="text" name="tWeight" size="16" onfocusout="CheckWeight()" tabindex="12" value=<% = gs(cdbl(rs("Weight")),4) %> >吨<font color="#FF0000">*</font></font></p>
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="宋体">　　毛　重 
        <!--webbot bot="Validation" s-data-type="Number"
        s-number-separators=",."
        s-validation-constraint="Greater than or equal to"
        s-validation-value="0" --><input type="text" name="tGross" size="14" tabindex="13" value=<% = rs("Gross") %> >公斤<font color="#FF0000">*</font>　&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        　净　重&nbsp;<!--webbot bot="Validation" s-data-type="Number"
        s-number-separators=",."
        s-validation-constraint="Greater than or equal to"
        s-validation-value="0" --><input type="text" name="tNet" size="16" tabindex="14" value=<% = rs("Net") %> >公斤<font color="#FF0000">*</font></font></p>                                     
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="宋体">　仓堆位置                                           
        <input type="text" name="tPlace1" tabindex="15" size="14" value=<% = rs("Place1") %> >仓&nbsp;&nbsp;&nbsp;&nbsp;                                           
        <input type="hidden" name="tRecNo"  size="14" value=<% = rs("RecNo") %> > 
        <input type="text" name="tPlace2" maxlength="40" size="12" tabindex="16" value=<% = rs("Place2") %> >库&nbsp;&nbsp;                                           
        &nbsp; <input type="text" name="tPlace3" maxlength="40" size="12" tabindex="17" value=<% = rs("Place3") %> >位</font></p>                                              
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="宋体">　　备　注 
        <input type="text" name="tRemark"  maxlength="40" size="60" tabindex="18" value=<% = rs("Remark") %> ></font></p>    
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="宋体">
        　送货单位 <input type="text" name="tDeliver" maxlength="40" size="20" value="<% =rs("Deliver") %>" >　　&nbsp; 
        　收货人 <input type="text" name="tConsignee" size="20" value="<% =rs("Consignee")%>" maxlength="40"></font></p>                                               
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="宋体">　　司　机                                           
        <input type="text" name="tDriver" maxlength="40" size="20" value="<% =rs("Driver") %>" >　　&nbsp;&nbsp;&nbsp;                                           
        车　号 <input type="text" name="tNumberPlate" size="20" value="<% =rs("NumberPlate")%>" maxlength="40"></font></p>                                               
        <hr color="#AFE0FA">
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2">&nbsp                                                       
      <input type="checkbox" name="C3" value="ON"></b><font face="宋体" size="2"><b><font color="#FF0000">保护出库单 </font></b><br>  &nbsp;&nbsp;&nbsp;  【每件重量不同的货物一定要选择打勾，否则只要保存了(即使是什么数据都没更改)，所有已经出库的出库单重量都会变 平均重量*件数,每件重量一样的和从来没出过库的请不要打勾】囧</font><b>
	 	 <hr color="#AFE0FA">
	 <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 

	 <input type="button" value="保 存" name="btSave" <%if OperaterMatch=0 then %> disabled = true <%end if%> OnClick="ValidateData()" style="font-family: 宋体; font-size: 12pt; font-weight: bold" tabindex="20" maxlength="40">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                                                           
        <input type="button" value="删 除" name="btDelete" style="font-family: 宋体; font-size: 12pt; font-weight: bold" OnClick="OnDeleteClick()" <%if OperaterMatch=0 then %> disabled=true <%end if%>>　&nbsp;                                                            
        &nbsp;&nbsp;&nbsp;<input type="reset" value="返 回" name="btCancel"  style="font-weight: bold; font-family: 宋体; font-size: 12pt" tabindex="21" OnClick="vbscript:history.back()">&nbsp;&nbsp;&nbsp;&nbsp;</p>         
  		<font color="#FF0000">

        <% 
		   if OperaterMatch=0 then
		   		response.write("<br>&nbsp;&nbsp;&nbsp;&nbsp;该进仓记录由用户 " & rs("operater") & " 创建，您没有修改此数据的权限。")
		  %>           
				<script languge=vbscript>
					' window.alert "该进仓记录由其他用户创建，您没有修改此数据的权限。"
				</script>
		 <%
		   end if

	       rs.close 
           set rs=nothing
           conn.close
           set conn=nothing 

           
        %>
        
		</font>
        </b>
      </form>
      </td>
  </tr>
</table>
</html>