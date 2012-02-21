<%
   
    set fs=server.createobject("scripting.filesystemobject")
    'set fs1=server.createobject("scripting.filesystemobject")


    if fs.folderexists(server.mappath("back"))=false then fs.createfolder(server.mappath("back"))
    
    'if fs.GetFolder(server.mappath("back")).Files.Count=20 then 
      
     
        
    'else
    
    'end if 
 
    for each fs1 in fs.getfolder(server.mappath("back")).files
        'set a=fs1
        response.write fs1.path         
    next
%>