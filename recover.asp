<%
if session("userright")<>"�ֿ�����" then
   response.write "��û��ʹ�õ�Ȩ�ޣ�"
   response.end
end if
%>

<form method="POST" action="recover1.asp">

&nbsp; 

<select name="filelist" size=12>
 <%
   set fs=server.createobject("scripting.filesystemobject")
   
   for each fs1 in fs.getfolder(server.mappath("back")).files
       response.write "<option>" & fs1.name & "</option>"       
   next
 
 %>   
</select>

<input type="submit" name="tj" value="�ָ�"> 
</form>