<!-- #include file="data\config.asp" -->
 
<Script Language=VBScript>
<!--
Sub OnDeleteClick()
	dim choice	
	Choice = msgbox ("ȷ��ɾ���ü�¼��?",vbYESNO,"ɾ����¼")
	if Choice = vbNO then
		Exit Sub
	end if
	
	window.location.href="OutStore_Delete.asp?RecNo=" & "<%=Request("RecNo")%>" 
	
End Sub

sub ValidaData()
    set ThisForm = Document.OutStoreDelete 
    
    'if cdbl(ThisForm.tGross.value) < cdbl(ThisForm.t11.value) then
	'    Window.Alert "ë��Ӧ���ڻ���ھ���"
	'	Exit Sub
	'End If

	ThisForm.Submit
end sub

-->
</Script>
 
 
<HTML>
<HEAD>
<TITLE>���ֹ���</TITLE>
<META content="Microsoft FrontPage 5.0" name=GENERATOR>
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


<TABLE border=0 cellPadding=0 cellSpacing=0 height="1">
  <TR>
    <TD width="579" valign="top" height="247"> 
      
   <%      Set rs = Server.CreateObject("ADODB.Recordset")
		   sql = "select * from OutStore where RecNo=" & request("RecNo")
		   rs.open sql, conn, 1, 1	
		   if rs.eof then
		   		if session("varClientID")="All" then
		   			response.Redirect "OutStore_Edit_All_Frame.asp"
		   		else
		   			response.Redirect "OutStore_Edit_Frame.asp"
		   		end if
		   		response.end
		   end if

			if rs("operater") = Session("UserName") then     'Ҫ�޸ĵļ�¼�Ƿ�Ϊ��ǰ����Ա¼���
				OperaterMatch=1
			else
				OperaterMatch=0
			end if
		   

		   
    %>
      <form name="OutStoreDelete" method="POST" action="OutStore_Save.asp?Recno=<%=rs("RecNo")%>"  >
        <b>

        </b>
        <hr color="#AFE0FA">
        
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><b><font size="3" face="����">�����ֵ���  
        <input type="text" name="T4" size="21" value=<% =rs("BillNo")%> >�������ᵥ��   
        <input type="text" name="T1" size="21" value= <% =rs("PickUpID") %> ></font></b>
        
        </p>

        
        <b>
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">����Ʒ����        
        <input type="text" name="tGoodsName" size="21" readonly tabindex="3" value=<% = rs("GoodsName") %> >�������桡�� <input type="text" name="tSpecs" size="21" readonly tabindex="4" value=<% =rs("Specs") %> ></font></p>      
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">��&nbsp;            
        ������                                                                                                                                                                                   
        <input type="text" name="tNumber" readonly size="9"  tabindex="6" value=<% = rs("Number1") %>>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                         
        &nbsp;&nbsp;������                                                                                                                 
        <input type="text" name="tAllPack" readonly size="21" readonly tabindex="10" value=<% = rs("Pack1") %>></font></p>                                                                
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
        <font size="3">ë&nbsp;&nbsp;&nbsp; �� <input type="text" name="tGross" size="19" value=<%=rs("Gross") %> >����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;   
        ������ <input type="text" name="T11" size="21" value=<%=rs("net") %> >����</font></p>                                                                   
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">�����ܡ���                                                               
        <input type="text" name="tWeight" size="19"  readonly tabindex="12" value=<% = gs(Cdbl(rs("Weight")),4) %>>��</font></p>
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">���ֶ�λ��                
        <input type="text" name="tPlace1" size="14" readonly value=<% = rs("Place1") %> tabindex="15">�֡�����<input type="text" name="tPlace2" size="12" readonly tabindex="16" value=<% = rs("Place2") %>>��                         
        &nbsp;&nbsp;&nbsp; <input type="text" name="tPlace3" size="12" readonly tabindex="17" value=<% = rs("Place3") %>>λ</font></p>                       
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">�����ա��� 
        <input type="text" name="T3" size="20"  value=<% =rs("date1") %> ></font></p>           
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">��������ע    
        <input type="text" name="tRemark" size="60" readonly tabindex="18" value=<% = rs("Remark") %>></font></p>    
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">���ջ���λ 
        <input type="text" name="T5" size="60" value=<% =rs("Crop") %> ></font></p>    
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">�������˵� 
        <input type="text" name="T6" size="24" value=<% =rs("StartAddress") %> >������������� <input type="text" name="T7" size="16" value=<% =rs("NumberPlate") %> ></font></p>    
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font size="3" face="����">������ʱ��  
        <input type="text" name="T8" size="24" value=<% =rs("date2") %> >�������������� <input type="text" name="T9" size="16" value=<% =rs("Consigner") %> ></font></p>     
        <hr color="#AFE0FA">
        <p style="line-height: 100%; margin-top: 2; margin-bottom: 2"><font face="����" size="3">��</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        ����������&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��    
        <input type="button" value="�� ��" name="btSave" style="font-size: 12pt; font-weight: bold" <%if OperaterMatch=0 then %> disabled=true <% end if %> OnClick="ValidaData()">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
        <input type="button" value="ɾ ��" name="btDelete" OnClick="OnDeleteClick()" style="font-family: ����; font-size: 12pt; font-weight: bold" tabindex="20" <%if OperaterMatch=0 then %> disabled=true  <% end if %> >��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
        <input type="reset" value="�� ��" name="btCancel" style="font-weight: bold; font-family: ����; font-size: 12pt" tabindex="21" OnClick="vbscript:history.back()">&nbsp;&nbsp;&nbsp;&nbsp;</p>                        
        <font color=#ff0000>
        <% 
		   if OperaterMatch=0 then
		   		response.write("<br>&nbsp;&nbsp;&nbsp;&nbsp;�ý��ּ�¼���û� " & rs("operater") & " ��������û���޸Ĵ����ݵ�Ȩ�ޡ�")
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