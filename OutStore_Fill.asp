<!-- #include file="data\config.asp" -->
 
 <SCRIPT LANGUAGE ="VBSCRIPT">
  <!-- 
    Sub ValidateData()
        set ThisForm = Document.OutStoreForm 


        If trim(ThisForm.tOutNumber.Value) = Empty Then
             Window.Alert "���ּ������ܿա���"
             Exit Sub
        End If
      
        If trim(ThisForm.tOutPack.Value) = Empty Then
            Window.Alert "���ְ������ܿա���"
            Exit Sub
        End If
        
        if clng(ThisForm.tOutPack.value) > clng(ThisForm.tAllPack.value) then
            Window.Alert "���ְ������ܴ��ڽ��������"
            Exit Sub
        End If
        
        If trim(ThisForm.tOutWeight.Value) = Empty Then
            Window.Alert "�����������ܿա���"
            Exit Sub
        End If
        
        if clng(ThisForm.tOutWeight.value) > clng(ThisForm.tWeight.value) then
            Window.Alert "�����������ܴ��ڽ���������"
            Exit Sub
        End If

    	if not IsNumbers(ThisForm.tOutNumber.value) then
    		Window.Alert "���ּ�������Ϊ����!"
    		Exit Sub
    	End If

    	if not IsNumbers(ThisForm.tOutPack.value) then
    		Window.Alert "���ְ�������Ϊ����!"
    		Exit Sub
    	End If

    	if not IsNumbers(ThisForm.tOutWeight.value) then
    		Window.Alert "������������Ϊ����!"
    		Exit Sub
    	End If
    	
    	if not( cint(ThisForm.tYear.value) >1990 and CInt(ThisForm.tYear.value)<2900 ) then 
    		Window.Alert "�Ʒ���ݲ���ȷ��"
    		Exit Sub
    	end if
    	if not( cint(ThisForm.tMonth.value) <=12 and CInt(ThisForm.tMonth.value)>=1 ) then 
    		Window.Alert "�Ʒ��·ݲ���ȷ��"
    		Exit Sub
    	end if
    	if not( cint(ThisForm.tDay.value) <=31 and CInt(ThisForm.tDay.value)>= 1 ) then 
    		Window.Alert "�Ʒ����ڲ���ȷ��"
    		Exit Sub
    	end if
    	
    	if not( cint(ThisForm.tSYear.value) >1990 and CInt(ThisForm.tSYear.value)<2900 ) then 
    		Window.Alert "������ݲ���ȷ��"
    		Exit Sub
    	end if
    	if not( cint(ThisForm.tSMonth.value) <=12 and CInt(ThisForm.tSMonth.value)>=1 ) then 
    		Window.Alert "�����·ݲ���ȷ��"
    		Exit Sub
    	end if
    	if not( cint(ThisForm.tSDay.value) <=31 and CInt(ThisForm.tSDay.value)>= 1 ) then 
    		Window.Alert "�������ڲ���ȷ��"
    		Exit Sub
    	end if

    	if cdbl(thisform.tOutWeight.value) > cdbl(thisForm.tWeight.value) then
    		Window.Alert "�������ز��ܴ��ڽ������أ�"
    		Exit Sub
    	end if

		
		'OnOutNumberChange    '�����ܰ���

		'CountWeight     '����������
		
		ThisForm.tAllPack.disabled=false
		
		ThisForm.Submit

    End Sub
    

    Sub OnOutNumberChange()    '�����ܰ���
    	set Form1=Document.OutStoreForm
    	If not IsNumbers(Form1.tOutNumber.value) then
    		Window.Alert "���ּ�������Ϊ����!"
    		Exit Sub
    	End IF
    	
    	if trim(Form1.tOutNumber.value) = empty then
    		window.alert "���ּ�������Ϊ��!"
    		exit sub
    	end if
    	
    	if cdbl(Form1.tOutNumber.value) > cdbl(Form1.tNumber.value) then
    		Window.Alert "���ּ������ܴ��ڽ��������"
    		Exit Sub
    	end if

   	
    	if (Form1.t2.value<> Empty) and (Form1.tOutNumber.value <>Empty) then
    		
    		Form1.tOutPack.value = Form1.tOutNumber.value * Form1.t2.value

    		 CountWeight    '�������������
    		
    	end if
    	    
    End Sub
    
    Sub CountWeight()      '����������
    	set Form2=Document.OutStoreForm
    	if (not IsNumbers(Form2.tPackWeight.value))  or (trim(form2.tPackWeight.value)=Empty) then
    		Window.Alert "û�а��������ݣ��޷����������������������!"
    		Exit Sub
    	end if
		if (not IsNumbers(Form2.tOutPack.value)) or ( trim(Form2.tOutPack.value)=Empty) then
			Window.Alert "�ܰ���δ����������֣��޷������������!"
			Exit Sub
		end if
	   	
	   	if not form2.chkWeight.checked then     
			Form2.tOutWeight.value = gs1(Form2.tPackWeight.value * Form2.tOutPack.value /1000,4)     'ת��Ϊ��
		end if
		
		form2.tOutNumber.value=gs1(form2.tOutPack.value / form2.t2.value,4)

    End Sub
    

    Sub CheckNumber()   '�������Ƿ�Ϊ����
    	if not IsNumbers(OutStoreForm.tNumber.value) then
    		Window.Alert "��������Ϊ����!"
    		Exit Sub
    	End If
    	if (OutStoreForm.tPackNumber.value<> Empty) and (OutStoreForm.tNumber.value <>Empty) then
    		OutStoreForm.tAllPack.value = OutStoreForm.tNumber.value * OutStoreForm.tPackNumber.value
    	end if
    End Sub
    

    Sub CheckAllPack()   '����ܰ����Ƿ�Ϊ����
    	if not IsNumbers(OutStoreForm.tAllPack.value) then
    		Window.Alert "�ܰ�������Ϊ����!"
    		Exit Sub
    	End If
    End Sub
    
    
    Sub CheckWeight()   '��������Ƿ�Ϊ����
    	if not IsNumbers(OutStoreForm.tWeight.value) then
    		Window.Alert "���ر���Ϊ����!"
    		Exit Sub
    	End If
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

		<%

		    Set rs =server.createobject("adodb.recordset")

			'���ɳ��ⵥ��No
			dim BillNo
			if CInt(month(date())) <10 then
				tempmonth="0" & month(date())
			else
				tempmonth=month(date())
			end if
	
			BillDate=year(date()) & "-" & tempmonth & "-01"   'ÿ�£���
	
			sql="select * from BillNo where NoType='����֪ͨ��' and Date1=#" & BillDate & "#"
			rs.open sql,conn,1,1
			if rs.recordcount<1 then
				BillNo=1
			else
			 	BillNo=rs("No1") + 1 
			end if
		 	rs.close

			for i=1 to  3-len(BillNo)    '��ʽ�����ֵ���
				BillNo= "0" + cstr(BillNo)
			next

			BillNo="C" & mid(BillDate,3,2) & mid(BillDate,6,2) & BillNo     '�������磺C0410XXX�ĵ���,��һλΪ���֣�Ȼ����λ�ֱ�Ϊ����,����λΪ��ˮ��
			


			'ȡ�ͻ�����
			sql = "select * from ClientInfo where ID = " & Request("ClientID")
			rs.open sql,conn,1,1
			if rs.eof then
				response.write "���ݴ���"
				Response.End
			end if
			tempClientName = rs("ClientName")
			tempID = rs("ID")
			rs.close


		    sql = "select * from InStore where Recno = "  &  request("RecNo") 
			rs.open sql ,conn,1,1
			if rs.eof then
				if session("varClientID")="All" then
					Response.Redirect "OutStore_Browse_All_Frame.asp"
				else
					Response.Redirect "OutStore_Browse_Frame.asp"
				end if
				response.end
			end if
		%>



<TABLE border=0 cellPadding=0 cellSpacing=0 height="366">
  <TR>
    <TD width="579" valign="top" height="366"> 
      <form name="OutStoreForm" method="POST" action="OutStore_Count.asp?Operate=Add&ButtonName=Save">
        <b>
        <input type="Hidden" name="tGross" size="14" tabindex="13" value=<% = rs("Gross") %>><input type="Hidden" name="tNet" size="19" tabindex="14" value=<% = rs("Net") %>><input type="hidden" name="tPack" size="18" tabindex="5" value=<% = rs("Pack") %>>
        </b>
        <hr color="#AFE0FA">
        
        <b>
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font face="����" size="3">�����ֵ���                                                  
        <input type="text" name="tBillNo" size="20" value="<% = BillNo %>"><font color="#FF0000">*</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                        
        �ᵥ��                                                          
        <input type="text" name="tPickUpID" size="20"></font></p>                 
        
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2">&nbsp;&nbsp;&nbsp;                        
        <font face="����" size="3">�������� <input type="text" name="tOutNumber" onfocusout="onOutNumberChange()" size="12"><font color="#FF0000">*</font>&nbsp;                        
        ÿ������                                                         
        <input type="text" name="T2" size="8" readonly <% if rs("AllNumber")>0 then %> value = <% = gs(cdbl(rs("PackNumber") / rs("AllNumber")),2) %> <% end if %> ><font color="#FF0000">*</font>&nbsp;                        
        �ܼ���                                                         
        <input type="text" name="tOutPack" onfocusout ="CountWeight()" size="10"><font color="#FF0000">*</font></font></p>          
        
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font face="����" size="3">�����ء���                          
        <input type="text" name="tOutWeight" size="17">��<font color="#FF0000">*</font><input type="checkbox" name="chkWeight" value="ON"></font><font face="����" size="2">����</font><font face="����" size="3">�� 
        ��ע </font><input type="text" name="tOutRemark" size="21"></p>                 
        
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font face="����" size="3">���Ʒ�����                              
        <input type="text" name="tYear" readonly size="6" value=<% =year( Now()) %>>��                             
        <input type="text" name="tMonth" size="3" value=<% = month( Now()) %>>��                             
        <input type="text" name="tDay" size="3" value=<% = day(Now()) %>>�� <font color="#FF0000">*</font></font></p>                                 
        
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font face="����" size="3">�������λ                         
        <input type="text" name="tCrop" size="60"></font></p>          
        
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font face="����" size="3">���������  
        <input type="text" name="tNumberPlate" size="20">�����������˵�   
        <input type="text" name="tStartAddress" size="20" value=<% = rs("Place1") %>></font></b></p>          
        
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><b><font face="����" size="3">����������    
        <input type="text" name="tSYear" size="5" align="right" value="<% =year(date()) %>">��<input type="text" name="tSMonth" size="3" value="<% =month(date()) %>">��<input type="text" name="tSDay" size="3" value="<% =day(date()) %>">�ա����������� 
        <input type="text" name="tConsigner" size="21"></font></b><b></p>          
        
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
       <input type="checkbox" name="C2" value="ON"></b><font face="����" size="3">��ʾë��</font><b>��&nbsp; ��<input type="checkbox" name="C1" value="ON"></b><font face="����" size="3">��ӡ����֪ͨ��</font><b>��&nbsp;����<input type="button" value="�� ��" name="btSave" OnClick="ValidateData()" style="font-family: ����; font-size: 12pt; font-weight: bold" tabindex="20">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                                                 
        <input type="reset" value="�� ��" name="btCancel" style="font-weight: bold; font-family: ����; font-size: 12pt" tabindex="21" OnClick="vbscript:history.back()">&nbsp;&nbsp;</p>         
        
        

        <hr color="#AFE0FA">
        

        </b>
        
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><b><font size="3" face="����">������� <select size="1" name="D1" tabindex="1" >                                                                                                                                                                             
		
	       		 <OPTION value = "<% = tempID %>" > <% = tempClientName %>  </Option>
        
        </select>

        
        ��</font>
        
        </b>
        
        </p>

        
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><b><font size="3" face="����">����ͬ��                                  
        <input type="text" name="tContractID" readonly size="17" tabindex="2" value=<% =rs("ContractID") %> >��Ʒ��       
        <input type="text" name="tGoodsName" size="11" readonly tabindex="3" value=<% = rs("GoodsName") %> >����� <input type="text" name="tSpecs" size="13" readonly tabindex="4" value=<% =rs("Specs") %> ></font>  
        
        </p>

        
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">��������                                                                                                                                                                                                                
        <input type="text" name="tNumber" readonly size="9"  tabindex="6" value=<% = rs("RemainNumber") %>>&nbsp;                                
        ÿ������                                                                                                                                             
        <input type="text" name="tpackNumber" size="9" readonly tabindex="8" <% if rs("AllNumber")>0 then %> value = <% = gs(cdbl(rs("PackNumber") / rs("AllNumber")),2) %> <% end if %> >&nbsp;&nbsp;�������                                                                                                                                              
        <input type="text" name="tAllPack" readonly size="13" tabindex="10" value=<% = rs("RemainPack") %>></font></p>                                                                
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">ÿ                                
        �� ��                                                                                                                     
        <input type="text" name="tPackWeight" size="14" readonly tabindex="11" value=<% = rs("PackWeight") %>>�����&nbsp;��&nbsp;                      
        ��������                                                                                                                    
        <input type="text" name="tWeight" size="18" readonly onfocusout="CheckWeight()" tabindex="12" value=<% = rs("RemainWeight") %>>��</font></p>
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">�ֶ�λ��&nbsp; 
        <!--webbot bot="Validation" s-data-type="Number"
        s-number-separators=",."
        s-validation-constraint="Greater than or equal to"
        s-validation-value="0" --><input type="text" name="tPlace1" size="14" readonly value=<% = rs("Place1") %> tabindex="15">�֡�����<input type="text" name="tPlace2" size="12" readonly tabindex="16" value=<% = rs("Place2") %>>�� 
        ��&nbsp;&nbsp; <input type="text" name="tPlace3" size="12" readonly tabindex="17" value=<% = rs("Place3") %>>λ</font></p>                                                          
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">������ע      
        <input type="text" name="tRemark" size="60" readonly tabindex="18" value=<% = rs("Remark") %>><input type="hidden" name="tRecNo" size="20" value="<% =rs("RecNo") %>"></font></p>    
        </b>
      </form>
      </td>
  </tr>
</table>
        <% rs.close 
           set rs=nothing
           conn.close
           set conn=nothing 
        %>

</html>