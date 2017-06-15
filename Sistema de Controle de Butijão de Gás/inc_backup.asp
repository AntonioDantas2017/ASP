<%
Dim objConn03,objrs03
	call abre_conexao03
	set objrs03 = objConn03.execute("SELECT DATA FROM CONFIG")
	%>
	<script>
//		alert("<%=(now() - 0.007) %>");
//		alert("<%=objrs03("data")%>");
	</script>
	<%
	if  cdate(objrs03("data")) <= (now() - 0.007) then

		dim fs,fo,x
      	'##### Enter your database details here #####
      	SourcePath = "c:\\inetpub\\wwwroot\\geral\\adm\\dados\\"
      	SourceFilename = "bd.mdb"
      	DestinationPath = "c:\\Backup\\"
      	BackupFile = "bd"
      	'############################################
      	DestinationFilename =  BackupFile & "_" & Year(now()) & "-" & right("0"&Month(Now()),2) & "-" & right("0"&Day(Now()),2)  & " " & replace(time,":",",") & ".mdb"
      	Set fs=Server.CreateObject("Scripting.FileSystemObject")
      	'Create new backup file

      	if fs.FileExists(DestinationPath & DestinationFileName) = False then
	       	fs.CopyFile SourcePath & SourceFileName, DestinationPath & DestinationFileName, False
      	Else
'          	Response.Write "Database already backed up today<br /><br />"
      	End if
      	set fo=fs.GetFolder(DestinationPath)
      	'Count Files
      	for each x in fo.files
        	Files = Files + 1
      	next
      	'Delete files > 10
      	for each x in fo.files
        	If Files > 10 Then
            	fs.DeleteFile(DestinationPath & x.Name)
              	Files = Files - 1
          	End if
      	next
      	set fo=nothing
      	set fs=nothing
'      	Response.Ends
		set objrs03 = objConn03.execute("UPDATE CONFIG SET DATA='"&now()&"'")
	end if	
	call fecha_conexao03	
%>
