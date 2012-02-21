<%
if session("userright")<>"仓库主管" then
   response.write "你没有使用的权限！"
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

<input type="submit" name="tj" value="恢复"> 
</form>