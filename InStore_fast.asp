<!-- #include file="data\config.asp" -->
 
 <SCRIPT LANGUAGE ="VBSCRIPT">
  <!-- 
    Sub ValidateData()
        set ThisForm = Document.InStoreForm 
      
        If trim(ThisForm.tGoodsName.Value) = Empty Then
            Window.Alert "Ʒ�����ܿգ�"
            Exit Sub
        End If
     
  	    if instr(1,trim(ThisForm.tGoodsName.Value),"'")>0 or instr(1,trim(ThisForm.tGoodsName.Value),chr(34))>0 then
        	Window.Alert "��Ʒ�����в��ܺ��� ' �� '' ��"
          	ThisForm.tGoodsName.focus()
          	Exit Sub
        end if
        
        
        If trim(ThisForm.tSpecs.Value) = Empty Then
            Window.Alert "����ܿգ�"
            Exit Sub
        End If

  	    if instr(1,trim(ThisForm.tSpecs.Value),"'")>0 or instr(1,trim(ThisForm.tSpecs.Value),chr(34))>0 then
        	Window.Alert "������в��ܺ��� ' �� '' ��"
          	ThisForm.tSpecs.focus()
          	Exit Sub
        end if


        If trim(ThisForm.tNumber.Value) = Empty Then
            Window.Alert "��������Ϊ�գ�"
            Exit Sub
        End If

  	    if instr(1,trim(ThisForm.tNumber.Value),"'")>0 or instr(1,trim(ThisForm.tNumber.Value),chr(34))>0 then
        	Window.Alert "���������в��ܺ��� ' �� '' ��"
          	ThisForm.tNumber.focus()
          	Exit Sub
        end if
        
		if trim(ThisForm.tAllPack.value) = Empty then
            Window.Alert "�ܼ�������Ϊ�գ�"
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
            Window.Alert "ÿ���ز���Ϊ�գ�"
            Exit Sub
        End If

        If trim(ThisForm.tWeight.Value) = Empty Then
            Window.Alert "���ز���Ϊ�գ�"
            Exit Sub
        End If

        If trim(ThisForm.tGross.Value) = Empty Then
            Window.Alert "ë�ز���Ϊ�գ�"
            Exit Sub
        End If

        If trim(ThisForm.tNet.Value) = Empty Then
            Window.Alert "���ز���Ϊ�գ�"
            Exit Sub
        End If

    	if not IsNumbers(ThisForm.tAllPack.value) then
    		Window.Alert "�ܼ�������Ϊ����!"
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
		
		'OnPackChange    '�����ܼ���

		'CountWeight     '����������
		
		ThisForm.tAllPack.disabled=false
		
		ThisForm.Submit

    End Sub
    

   
    Sub OnR1V1()		'�ܼ���Disabled

    	Document.InStoreForm.tAllPack.disabled=true
    	Document.InStoreForm.tPackNumber.disabled = false
    	

    End Sub
    
	Sub OnR1V2()

    	Document.InStoreForm.tPackNumber.disabled = true
    	Document.InStoreForm.tAllPack.disabled = False
		
	End Sub    

    Sub OnPackChange()    '�����ܼ���
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
			' Window.Alert "ÿ����������Ϊ��!"
			Exit Sub
		end if
		if trim(Form1.tAllPack.value)=Empty then
			Window.Alert "�ܼ�������Ϊ��!"
			Exit Sub
		end if
	   	if form1.tweight.value=empty then
		   	Form1.tWeight.value = gs1(Form1.tPackWeight.value * Form1.tAllPack.value /1000,4)	
		   	Form1.tNet.value=gs1(Form1.tWeight.value * 1000 / form1.tAllPack.value, 4)
	    end if
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
    

    Sub CheckAllPack()   '����ܼ����Ƿ�Ϊ����
    	if not IsNumbers(InStoreForm.tAllPack.value) then
    		Window.Alert "�ܼ�������Ϊ����!"
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
			
			InStoreForm.tGross.value = InStoreForm.tPackWeight.value
			
    		
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

<body background="image/blackground.gif">


<TABLE border=0 cellPadding=0 cellSpacing=0 height=114 width=836>
  <TR>
    <TD background=image/top.gif colSpan=2 rowSpan=2 width=241 height="114">
      ��
	</TD>
    <TD bgColor=#00b2eb height=3 width=591>
    <IMG height=1 src="../smis/image/pixel.gif" width=1></TD>
  </TR>
  <TR>
    <TD bgColor=#ffffff vAlign=bottom width="591" height="111">
      <DIV align=right>
      <p align="center">
      <img border="0" src="image/buttons.gif" width="460" height="40"></p>
      <TABLE align=right border=0 cellPadding=0 cellSpacing=0 height=32 width=307 vspace="0" hspace="0">
        <TR>
          <TD background=image/button1.gif width=162 align="left">
            <DIV align=center>
              <p align="center"><blink><b><font size="4">�� �� �� ��</font></b></blink></DIV>  
		  </TD>
          <TD vAlign=bottom width="141" align="left">
		  </TD>
		</TR>
	  </TABLE>
	  </DIV>
	</TD>
  </TR>
</TABLE>
<TABLE border=0 cellPadding=0 cellSpacing=0 width=852 height="1">
  <TR>
    <TD height=1 width=250></TD>
    <TD rowSpan=3 width=5 height="1">&nbsp;</TD>
    <TD height=1 width=599></TD>
  </TR>
  <TR>
    <TD vAlign=top width=250 height="1" rowspan="2">
      <TABLE align=right border=0 cellPadding=4 cellSpacing=4 height=270 width=140>
        <TR bgColor=#29a8cd>
          <TD width="122">
            <DIV align=center><b><a href="index.asp?op=reset"><font size="2" color="#FFFFFF">���µ�¼</font></a></b></DIV>
		  </TD>
		</TR>
        <TR bgColor=#2cb1bc>
          <TD width="122">
            <DIV align=center><b><font color="#ffffff" size="2">���ֹ���</font></b></DIV>
		  </TD>
		</TR>
        <TR bgColor=#35a8a5>
          <TD width="122">
            <DIV align=center><b><a href="OutStore.asp"><font color="#ffffff" size="2">���ֹ���</font></a></b></DIV>
		  </TD>
		</TR>
        <TR bgColor=#35a88e>
          <TD width="122">
            <DIV align=center><b><a href="Client_List.htm"><font color="#ffffff" size="2">�ͻ���Ϣ</font></a></b></DIV>
		  </TD>
		</TR>
        <TR bgColor=#37a67c>
          <TD width="122">
            <DIV align=center><b><font color="#FFFFFF"><a href="User.htm"><font color="#ffffff" size="2">�û�����</font></a></font></b></DIV>
		  </TD>
		</TR>
        <TR bgColor=#339966>
          <TD width="122">
            <DIV align=center><b><a href="jiliang_List.htm"><font color="#ffffff" size="2">ѡ 
              ��</font></a></b></DIV>
		  </TD>
		</TR>
        <TR bgColor=#339966>
          <TD width="122">
            <DIV align=center><b><A href="report.htm"><font color="#ffffff" size="2">������Ʒ�</font></A></b></DIV>
		  </TD>
		</TR>
        <TR bgColor=#339966>
         <TD width="122">
            <DIV align=center><b><A href="cremain.htm"><font color="#ffffff" size="2">������</font></A></b></DIV>
		  </TD>
		</TR>

        <TR bgColor=#339966>
          <TD width="122">
            <DIV align=center><b><A href="recover.htm"><font color="#ffffff" size="2">���ݻָ�</font></A></b></DIV>
		  </TD>
		</TR>

        <TR bgColor=#339966>
          <TD width="122">
            <DIV align=center><A href="support.htm"><font color="#ffffff" size="2"><b>��&nbsp;                                                                                                                                                                                    
              ��</b></font></A></DIV>                                                                                                                             
		  </TD>
		</TR>
	  </TABLE>
      <P align=right>��</P>
	</TD>
    <TD width="589" valign="top" height="1"> 
      <table border="0" width="100%" height="31">
        <tr>
          <td width="20%" height="27" background="image/button3.gif">
            <p align="right"><font face="����" size="3"><b>
           <a href="InStore.asp">��д����</a></b></font></td>                                                                      
          <td width="25%" height="27" background="image/button3.gif">
            <p align="right"><b><font face="����" size="4">
			 <img border="0" src="image/stargold.gif" width="14" height="18">��������&nbsp;</font></b></td>      

 <td width="20%" height="27" background="image/button3.gif">
            <p align="right"><b><font face="����" size="4"><a href="InStore_Browse.asp">�ִ����</a></font></b></td>  
			
          <td width="38%" height="27" background="image/button3.gif">
            <p align="right"><b><font face="����" size="4"><a href="Print_InStore.asp">��ӡ���֪ͨ��</a></font></b></td>                                                                   
        </tr>
      </table>
      </td>
  </tr>
  <TR>
    <TD width="589" valign="top" height="57"> 
      <form name="InStoreForm" method="POST" action="InStore_Add.asp?Operate=Add&Is_fast=yes">
        <hr color="#A3D0E9">
 
<%


bill_last=request("BillNo")                          'ȡ�ص���
Client_ID_last=request("ClientID")                '   ȡ�ؿͻ�ID��



    Set rs =server.createobject("adodb.recordset")

	'������ⵥ��No
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
	 	BillNo=rs("No1") +1 
	end if
 	rs.close

	for i=1 to  3-len(BillNo)    '��ʽ�����ֵ���
		BillNo= "0" + cstr(BillNo)
	next

	BillNo="J" & mid(BillDate,3,2) & mid(BillDate,6,2) & BillNo     '�������磺J0410XXX�ĵ���,��һλ�ǽ��ֵ�ƴ����Ȼ����λ�ֱ�Ϊ����,����λΪ��ˮ��
			

			
if bill_last<>"" then            '���ڴ�INSTORE_add.asp��ȡ�ص��ţ�����лش����ž�ʹ�ã����û�о�����
BillNo=bill_last
end if
			
			
			
			
			%>
       
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><b><font size="3" face="����">����������                                                                                          
        <input type="text" name="tBIllNo" size="20" value="<%=BillNo%>"><font color="#FF0000">*</font>&nbsp;&nbsp;&nbsp;                     
        �ա��� <select size="1" name="DYear">                                                                                    
        <%for i=year(date())-5 to year(date()) + 5 %> 
        		<option <% if year(date())=i then %> selected <% end if %> ><% = i %></option>
        <% next %>
        
        </select>��<select size="1" name="Dmonth">    
		<% for i=1 to 12 
		        if i<10 then
		        	monthvalue="0" + CStr(i)
		        else
		        	monthvalue=CStr(i)
		        end if
		 %>
          <option value="<% =monthvalue %>" <% if i=month(date()) then %> selected <% end if %> ><% = i %> </option>
        <% next %>
        
        </select>��<select size="1" name="DDay"> 
        <% for i=1 to 31  
        		if i<10 then
        			dayvalue = "0" + CStr(i)
        		else
        			dayvalue=CStr(i)
        		end if
           
        %>
        	<option value="<% =dayvalue %>" <% if  i= day(date()) then %> selected <% end if %> ><% =i %></option>
        <% Next %>
        </select>��</font></b> 
        
        </p>
        
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><b><font size="3" face="����">���������� <select size="1" name="D1" tabindex="1" >                                                                                                                                                 
		
		<% 
		 if Client_ID_last<>"" then
		 			rs.open "select * from ClientInfo where ID="&Client_ID_last& " order by ClientName asc",conn,1,1

		 else
		 
		
			rs.open "select * from ClientInfo order by ClientName asc",conn,1,1
         end if

           do while not rs.EOF %> 
	       		 <OPTION value = "<% = rs("ID") %>" > <% = rs("ClientName") %>  </Option>
         		<% rs.MoveNext 
	       loop 
    	   
    	   rs.close 

         %>
        
        
        </select><font color="#FF0000">*</font></font></b> 
        
        </p>
        
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><b><font size="3" face="����">������ͬ��                       
        <input type="text" name="tContractID" size="20" tabindex="2" maxlength="40"></font>
        
        </p>
        
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">����Ʒ����                                                                                                               
        <input type="text" name="tGoodsName" size="18" tabindex="3" maxlength="40"><font color="#FF0000">*</font>����&nbsp;&nbsp;&nbsp;                     
        �桡�� <input type="text" name="tSpecs" size="22" tabindex="4"><font color="#FF0000">*</font></font></p>                                                                                                            
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">��������װ                                                                                     
        <input type="text" name="tPack" size="18" tabindex="5" value="��װ" maxlength="40">������</font></p>
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">����������                                                                                        
        <input type="text" name="tNumber" size="9" onfocusout="CheckNumber()" tabindex="6" ><font color="#FF0000">*</font> 
        <input type="radio" value="V1" checked name="R1" onclick="OnR1V1()" tabindex="7" >ÿ������<input type="text" name="tpackNumber" size="10" onfocusout="OnPackChange()" tabindex="8"><font color="#FF0000">*</font>                     
        <input type="radio" name="R1" value="V2" onClick="OnR1V2()" tabindex="9" >�ܼ���<input type="text" name="tAllPack" disabled onfocusout="CheckAllPack()" size="12" tabindex="10" ><font color="#FF0000">*</font></font></p>                                                                                   
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">��ÿ������                                                                                                               
        <input type="text" name="tPackWeight" size="14" onfocusout="CountWeight()" tabindex="11" maxlength="40">����<font color="#FF0000">*</font>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                     
        �ܡ���                                                                                                              
        <input type="text" name="tWeight" size="19" onfocusout="CheckWeight()" tabindex="12" maxlength="40">��<font color="#FF0000">*</font></font></p>
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">����ë����&nbsp;<!--webbot
        bot="Validation" s-data-type="Number" s-number-separators=",."
        s-validation-constraint="Greater than or equal to"
        s-validation-value="0" --><input type="text" name="tGross" size="14" tabindex="13" maxlength="40">����<font color="#FF0000">*</font>&nbsp;       
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; �������� <!--webbot bot="Validation"
        s-data-type="Number" s-number-separators=",."
        s-validation-constraint="Greater than or equal to"
        s-validation-value="0" --><input type="text" name="tNet" size="19" tabindex="14" maxlength="40">����<font color="#FF0000">*</font></font></p>                                     
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">���ֶ�λ��                                                                                                                                                         
        <input type="text" name="tPlace1" size="14" value="���˲�" tabindex="15" maxlength="40">��&nbsp;&nbsp;&nbsp;                                                   
        &nbsp;&nbsp; <input type="text" name="tPlace2" size="12" tabindex="16" maxlength="40">��&nbsp;&nbsp;  
        &nbsp; <input type="text" name="tPlace3" size="12" tabindex="17" maxlength="40">λ</font></p>                                                     
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">��������ע   
        <input type="text" name="tRemark" size="62" tabindex="18" maxlength="40"></font></p>    
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">���ͻ���λ   
        <input type="text" name="tDeliver" size="20" maxlength="40">������&nbsp; 
        ���ջ��� <input type="text" name="tConsignee" size="20" maxlength="40"></font></p>      
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">����˾����    
        <input type="text" name="tDriver" size="20" maxlength="40">������&nbsp; 
        �������� <input type="text" name="tNumberPlate" size="20" maxlength="40"><input type="hidden" name="tNo" size="20" maxlength="40"></font></p>      
        <hr color="#A3D0E9">
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2">&nbsp;&nbsp;&nbsp;����������&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                                                 
        &nbsp;&nbsp;&nbsp;������<input type="checkbox" name="chkPrintInStoreTable" value="ON" maxlength="40" checked=""ON"><font face="����" size="3">��ӡ���֪ͨ��</font>��������                                                                                                    
        <input type="button" value="��������" name="btOK" OnClick="ValidateData()" style="font-family: ����; font-size: 12pt; font-weight: bold" tabindex="20" maxlength="40">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                                                                                                    
        <input type="reset" value="ȡ ��" name="btCancel" maxlength="40" style="font-weight: bold; font-family: ����; font-size: 12pt" tabindex="21"></p>         
        </b>
		<p>˵�����������ֻ��ڵڶ�����������ʱ�ظ��ϴ�����ĵ��źͻ���</p>
      </form>
      <p style="line-height: 100%; margin-top: 2">��
      </td>
  </tr>
</table>
</html>