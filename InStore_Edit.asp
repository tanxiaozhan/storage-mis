<!-- #include file="data\config.asp" -->
 
 <SCRIPT LANGUAGE ="VBSCRIPT">
  <!-- 
    Sub ValidateData()
        set ThisForm = Document.InStoreForm 
     
        If trim(ThisForm.tGoodsName.Value) = Empty Then
            Window.Alert "Ʒ�����ܿա���"
            Exit Sub
        End If

  	    if instr(1,trim(ThisForm.tGoodsName.Value),"'")>0 or instr(1,trim(ThisForm.tGoodsName.Value),chr(34))>0 then
        	Window.Alert "��Ʒ�����в��ܺ��� ' �� '' ��"
          	ThisForm.tGoodsName.focus()
          	Exit Sub
        end if

        If trim(ThisForm.tSpecs.Value) = Empty Then
            Window.Alert "����ܿա���"
            Exit Sub
        End If

  	    if instr(1,trim(ThisForm.tSpecs.Value),"'")>0 or instr(1,trim(ThisForm.tSpecs.Value),chr(34))>0 then
        	Window.Alert "������в��ܺ��� ' �� '' ��"
          	ThisForm.tSpecs.focus()
          	Exit Sub
        end if

        If trim(ThisForm.tNumber.Value) = Empty Then
            Window.Alert "��������Ϊ�ա���"
            Exit Sub
        End If
 
  	    if instr(1,trim(ThisForm.tNumber.Value),"'")>0 or instr(1,trim(ThisForm.tNumber.Value),chr(34))>0 then
        	Window.Alert "���������в��ܺ��� ' �� '' ��"
          	ThisForm.tNumber.focus()
          	Exit Sub
        end if
        
		if trim(ThisForm.tAllPack.value) = Empty then
            Window.Alert "�ܼ�������Ϊ�ա���"
            Exit Sub
        End If

  	    if instr(1,trim(ThisForm.tAllPack.Value),"'")>0 or instr(1,trim(ThisForm.tAllPack.Value),chr(34))>0 then
        	Window.Alert "���ܼ������в��ܺ��� ' �� '' ��"
          	ThisForm.tAllPack.focus()
          	Exit Sub
        end if

  	    if instr(1,trim(ThisForm.tPlace1.Value),"'")>0 or instr(1,trim(ThisForm.tPlace1.Value),chr(34))>0 then
        	Window.Alert "����λ1���в��ܺ��� ' �� '' ��"
          	ThisForm.tPlace1.focus()
          	Exit Sub
        end if
  	    if instr(1,trim(ThisForm.tPlace2.Value),"'")>0 or instr(1,trim(ThisForm.tPlace2.Value),chr(34))>0 then
        	Window.Alert "����λ2���в��ܺ��� ' �� '' ��"
          	ThisForm.tPlace2.focus()
          	Exit Sub
        end if
  	    if instr(1,trim(ThisForm.tPlace3.Value),"'")>0 or instr(1,trim(ThisForm.tPlace3.Value),chr(34))>0 then
        	Window.Alert "����λ3���в��ܺ��� ' �� '' ��"
          	ThisForm.tPlace3.focus()
          	Exit Sub
        end if
  	    if instr(1,trim(ThisForm.tremark.Value),"'")>0 or instr(1,trim(ThisForm.tremark.Value),chr(34))>0 then
        	Window.Alert "����ע���в��ܺ��� ' �� '' ��"
          	ThisForm.tremark.focus()
          	Exit Sub
        end if
        
  	    if instr(1,trim(ThisForm.tdeliver.Value),"'")>0 or instr(1,trim(ThisForm.tdeliver.Value),chr(34))>0 then
        	Window.Alert "���ͻ���λ���в��ܺ��� ' �� '' ��"
          	ThisForm.tdeliver.focus()
          	Exit Sub
        end if
  	    if instr(1,trim(ThisForm.tconsignee.Value),"'")>0 or instr(1,trim(ThisForm.tconsignee.Value),chr(34))>0 then
        	Window.Alert "���ջ��ˡ��в��ܺ��� ' �� '' ��"
          	ThisForm.tconsignee.focus()
          	Exit Sub
        end if
  	    if instr(1,trim(ThisForm.tdriver.Value),"'")>0 or instr(1,trim(ThisForm.tdriver.Value),chr(34))>0 then
        	Window.Alert "��˾�����в��ܺ��� ' �� '' ��"
          	ThisForm.tdriver.focus()
          	Exit Sub
        end if
  	    if instr(1,trim(ThisForm.tNumberPlate.Value),"'")>0 or instr(1,trim(ThisForm.tNumberPlate.Value),chr(34))>0 then
        	Window.Alert "�����š��в��ܺ��� ' �� '' ��"
          	ThisForm.tNumberPlate.focus()
          	Exit Sub
        end if
		
        If trim(ThisForm.tPackWeight.Value) = Empty Then
            Window.Alert "���ز���Ϊ�ա���"
            Exit Sub
        End If

        If trim(ThisForm.tWeight.Value) = Empty Then
            Window.Alert "���ز���Ϊ�ա���"
            Exit Sub
        End If

        If trim(ThisForm.tGross.Value) = Empty Then
            Window.Alert "ë�ز���Ϊ�ա���"
            Exit Sub
        End If

        If trim(ThisForm.tNet.Value) = Empty Then
            Window.Alert "���ز���Ϊ�ա���"
            Exit Sub
        End If

    	if not IsNumbers(ThisForm.tAllPack.value) then
    		Window.Alert "�ܰ�������Ϊ����!"
    		Exit Sub
    	End If

    	if not IsNumbers(ThisForm.tGross.value) then
    		Window.Alert "ë�ر���Ϊ����!"
    		Exit Sub
    	End If

    	if not IsNumbers(ThisForm.tNet.value) then
    		Window.Alert "���ر���Ϊ����!"
    		Exit Sub
    	End If

		if cdbl(ThisForm.tGross.value) < cdbl(ThisForm.tNet.value) then
			Window.Alert "ë��Ӧ���ڻ���ھ���"
			Exit Sub
		End If
		
		'OnPackChange    '�����ܰ���

		'CountWeight     '����������
		
		ThisForm.tAllPack.disabled=false
		
		ThisForm.Submit

    End Sub
    

   
    Sub OnR1V1()		'�ܰ���Disabled

    	Document.InStoreForm.tAllPack.disabled=true
    	Document.InStoreForm.tPackNumber.disabled = false
    	

    End Sub
    
	Sub OnR1V2()

    	Document.InStoreForm.tPackNumber.disabled = true
    	Document.InStoreForm.tAllPack.disabled = False
		
	End Sub    

    Sub OnPackChange()    '�����ܰ���
    	set Form1=Document.InStoreForm
    	If not IsNumbers(Form1.tPackNumber.value) then
    		Window.Alert "��������Ϊ����!"
    		Exit Sub
    	End IF
    	if (InStoreForm.tPackNumber.value<> Empty) and (InStoreForm.tNumber.value <>Empty) then
    		InStoreForm.tAllPack.value = InStoreForm.tNumber.value * InStoreForm.tPackNumber.value
    	end if
    	    
    End Sub
    
    Sub CountWeight()      '����������
    	set Form1=Document.InStoreForm
    	if not IsNumbers(Form1.tPackWeight.value) then
    		Window.Alert "���ر���Ϊ����!"
    		Exit Sub
    	end if
		if trim(Form1.tPackWeight.value) = Empty then
			' Window.Alert "���ز���Ϊ��!"
			Exit Sub
		end if
		if trim(Form1.tAllPack.value)=Empty then
			Window.Alert "�ܰ�������Ϊ��!"
			Exit Sub
		end if
	   	Form1.tWeight.value = gs1(Form1.tPackWeight.value * Form1.tAllPack.value /1000,4)
    End Sub
    

    Sub CheckNumber()   '�������Ƿ�Ϊ����
    	if not IsNumbers(InStoreForm.tNumber.value) then
    		Window.Alert "��������Ϊ����!"
    		Exit Sub
    	End If
    	if (InStoreForm.tPackNumber.value<> Empty) and (InStoreForm.tNumber.value <>Empty) then
    		InStoreForm.tAllPack.value = InStoreForm.tNumber.value * InStoreForm.tPackNumber.value
    	end if
    End Sub
    

    Sub CheckAllPack()   '����ܰ����Ƿ�Ϊ����
    	if not IsNumbers(InStoreForm.tAllPack.value) then
    		Window.Alert "�ܰ�������Ϊ����!"
    		Exit Sub
    	End If
    	if instoreform.tPackNumber.value = empty then
    		instoreform.tPackNumber.value=gs1( instoreform.tAllPack.value / instoreform.tNumber.value , 2 )
    	end if
    	
    	if instoreform.tpacknumber.value<> empty then
    		instoreform.tnumber.value=gs1(instoreform.tallpack.value / instoreform.tpacknumber.value,4)
	    end if
    	
    End Sub
    
    
    Sub CheckWeight()   '��������Ƿ�Ϊ����
    	if not IsNumbers(InStoreForm.tWeight.value) then
    		Window.Alert "���ر���Ϊ����!"
    		Exit Sub
    	End If
    	if (InStoreForm.tWeight.value<>Empty) and (InStoreForm.tAllPack.value<>Empty) then
    		InStoreForm.tPackWeight.value = gs1( InStoreForm.tWeight.value * 1000 / InStoreForm.tAllPack.value,4)
    		InStoreForm.tNet.value = InStoreForm.tPackWeight.value
    	end if
    	
    End Sub



    '����Ƿ�Ϊ����
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
		Choice = msgbox ("ɾ�����޷��ָ���ȷ��ɾ���ü�¼��?",vbYESNO,"ɾ����¼")
		if Choice = vbNO then
			Exit Sub
		end if
		window.location.href="InStore_Add.asp?Operate=Edit&ButtonName=Delete&RecNo=" & "<%=Request("RecNo")%>"
	
	End Sub


  -->
  </SCRIPT>


<HTML>
<HEAD>
<TITLE>���ֹ���</TITLE>
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
			if rs("operater") = Session("UserName") then     'Ҫ�޸ĵļ�¼�Ƿ�Ϊ��ǰ����Ա¼���
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
        
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><b><font size="3" face="����">����������                                           
        <input type="text" name="tBillNo" size="12" value="<% =rs("BillNo") %>"><font color="#FF0000">*</font>���ա���                                              
        <select size="1" name="DYear">
        <% for i=tempYear-5 to tempYear+5 %>  
        		<option <% if i=tempYear then %> selected <% end if %> ><%=i%></option>
        <% next %>
                	
        </select>��<select size="1" name="DMonth">  
        <% for i=1 to 12   
        		if i<10 then
        			tempvalue="0" + trim(CStr(i))
        		else
        			tempvalue=trim(CStr(i))
        		end if
        		
        %>
        		<option value="<% = tempvalue %>" <% if i=tempMonth then %> selected <% end if %> ><% =i %></option>
        <% next %>
        
        </select>��<select size="1" name="DDay">  
        <% for i=1 to 31%>  
               <option <% if i=tempDay then %> selected <% end if %> ><% =i %></option>
        <% next %>
               </select>��</font></b></p>  

        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><b><font size="3" face="����">
        ���������� <select size="1" name="D1" tabindex="1" >                                                                                                                                                                    
		
		<%	  
			rs.close
			sql = "select * from ClientInfo  " 
			rs.open sql,conn,1,1
			if rs.eof then
				response.write "�ͻ����ݴ���"
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

        
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><b><font size="3" face="����">������ͬ��                                           
        <input type="text" name="tContractID" size="20" maxlength="40" tabindex="2" value='<% =rs("ContractID") %>'></font>
        
        </p>
        
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">����Ʒ����                                           
        <input type="text" name="tGoodsName" size="24" maxlength="40" tabindex="3" value=<% = rs("GoodsName") %>><font color="#FF0000">*</font>����&nbsp;&nbsp;&nbsp;&nbsp;                                           
        �桡�� <input type="text" name="tSpecs" size="19" tabindex="4" value=<% =rs("Specs") %> maxlength="40"><font color="#FF0000">*</font></font></p>                                           
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">��������װ                                           
        <input type="text" name="tPack" size="18" maxlength="40" tabindex="5" value=<% = rs("Pack") %>>������</font></p>
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">����������                                           
        <input type="text" name="tNumber" size="9" onfocusout="CheckNumber()" maxlength="40" tabindex="6" value=<% = rs("AllNumber") %>><font color="#FF0000">*</font> 
        <input type="radio" value="V1" checked name="R1" onclick="OnR1V1()" tabindex="7">ÿ������<input type="text" name="tpackNumber" maxlength="40" size="10" onfocusout="OnPackChange()" tabindex="8" <% if rs("AllNumber")>0 then %> value = <% = gs(cdbl(rs("PackNumber") / rs("AllNumber")),2) %> <% end if %> ><font color="#FF0000">*</font>                                           
        <input type="radio" name="R1" value="V2" onClick="OnR1V2()" tabindex="9" maxlength="40">�ܼ���<input type="text" name="tAllPack" disabled onfocusout="CheckAllPack()" size="10" tabindex="10" value=<% = rs("PackNumber") %> ><font color="#FF0000">*</font></font></p>                                                                                                          
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">��ÿ������                                           
        <input type="text" name="tPackWeight" size="14" onfocusout="CountWeight()" tabindex="11" value=<% = rs("PackWeight") %> >����<font color="#FF0000">*</font>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                                           
        �ܡ���                                               
        <input type="text" name="tWeight" size="16" onfocusout="CheckWeight()" tabindex="12" value=<% = gs(cdbl(rs("Weight")),4) %> >��<font color="#FF0000">*</font></font></p>
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">����ë���� 
        <!--webbot bot="Validation" s-data-type="Number"
        s-number-separators=",."
        s-validation-constraint="Greater than or equal to"
        s-validation-value="0" --><input type="text" name="tGross" size="14" tabindex="13" value=<% = rs("Gross") %> >����<font color="#FF0000">*</font>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        ��������&nbsp;<!--webbot bot="Validation" s-data-type="Number"
        s-number-separators=",."
        s-validation-constraint="Greater than or equal to"
        s-validation-value="0" --><input type="text" name="tNet" size="16" tabindex="14" value=<% = rs("Net") %> >����<font color="#FF0000">*</font></font></p>                                     
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">���ֶ�λ��                                           
        <input type="text" name="tPlace1" tabindex="15" size="14" value=<% = rs("Place1") %> >��&nbsp;&nbsp;&nbsp;&nbsp;                                           
        <input type="hidden" name="tRecNo"  size="14" value=<% = rs("RecNo") %> > 
        <input type="text" name="tPlace2" maxlength="40" size="12" tabindex="16" value=<% = rs("Place2") %> >��&nbsp;&nbsp;                                           
        &nbsp; <input type="text" name="tPlace3" maxlength="40" size="12" tabindex="17" value=<% = rs("Place3") %> >λ</font></p>                                              
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">��������ע 
        <input type="text" name="tRemark"  maxlength="40" size="60" tabindex="18" value=<% = rs("Remark") %> ></font></p>    
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">
        ���ͻ���λ <input type="text" name="tDeliver" maxlength="40" size="20" value="<% =rs("Deliver") %>" >����&nbsp; 
        ���ջ��� <input type="text" name="tConsignee" size="20" value="<% =rs("Consignee")%>" maxlength="40"></font></p>                                               
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">����˾����                                           
        <input type="text" name="tDriver" maxlength="40" size="20" value="<% =rs("Driver") %>" >����&nbsp;&nbsp;&nbsp;                                           
        ������ <input type="text" name="tNumberPlate" size="20" value="<% =rs("NumberPlate")%>" maxlength="40"></font></p>                                               
        <hr color="#AFE0FA">
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2">&nbsp                                                       
      <input type="checkbox" name="C3" value="ON"></b><font face="����" size="2"><b><font color="#FF0000">�������ⵥ </font></b><br>  &nbsp;&nbsp;&nbsp;  ��ÿ��������ͬ�Ļ���һ��Ҫѡ��򹴣�����ֻҪ������(��ʹ��ʲô���ݶ�û����)�������Ѿ�����ĳ��ⵥ��������� ƽ������*����,ÿ������һ���ĺʹ���û��������벻Ҫ�򹴡���</font><b>
	 	 <hr color="#AFE0FA">
	 <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 

	 <input type="button" value="�� ��" name="btSave" <%if OperaterMatch=0 then %> disabled = true <%end if%> OnClick="ValidateData()" style="font-family: ����; font-size: 12pt; font-weight: bold" tabindex="20" maxlength="40">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                                                           
        <input type="button" value="ɾ ��" name="btDelete" style="font-family: ����; font-size: 12pt; font-weight: bold" OnClick="OnDeleteClick()" <%if OperaterMatch=0 then %> disabled=true <%end if%>>��&nbsp;                                                            
        &nbsp;&nbsp;&nbsp;<input type="reset" value="�� ��" name="btCancel"  style="font-weight: bold; font-family: ����; font-size: 12pt" tabindex="21" OnClick="vbscript:history.back()">&nbsp;&nbsp;&nbsp;&nbsp;</p>         
  		<font color="#FF0000">

        <% 
		   if OperaterMatch=0 then
		   		response.write("<br>&nbsp;&nbsp;&nbsp;&nbsp;�ý��ּ�¼���û� " & rs("operater") & " ��������û���޸Ĵ����ݵ�Ȩ�ޡ�")
		  %>           
				<script languge=vbscript>
					' window.alert "�ý��ּ�¼�������û���������û���޸Ĵ����ݵ�Ȩ�ޡ�"
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