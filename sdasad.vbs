option explicit 
 

 
    Dim oSession        ' AS NotesSession 
    Dim strServer 
    Dim strUserName 
    Dim strMailDbName 
    Dim oCurrentMailDb  ' as NOTESDATABASE 
    Dim oMailDoc        ' as NOTESDOCUMENT 
    Dim ortItem         ' as NOTESRICHTEXTITEM 
    Dim ortAttacment    ' as NOTESRICHTEXTITEM 
    Dim oEmbedObject    ' as ???? 
    dim cstrAttachment 
    Dim blAttachment 
 
    cstrAttachment =  "C:\softwaretest\test.txt"
 
    blAttachment = True 
 
    ' Start a session to notes 
    wscript.echo "## Connecting to Lotus Notes session..." 
    Set oSession = CreateObject("Notes.NotesSession") 
 
    wscript.echo("NotesVersion     : " & oSession.NotesVersion) 
    wscript.echo("NotesBuildVersion: " & oSession.NotesBuildVersion) 
    wscript.echo("UserName         : " & oSession.UserName) 
    wscript.echo("EffectiveUserName: " & oSession.EffectiveUserName) 
 
    wscript.echo "## GetEnvironmentString..." 
    strServer = oSession.GetEnvironmentString("MailServer",True) 
    wscript.echo("Server           :" & strServer) 
 
    ' eg. CN=Michael V Green/OU=CORPAU/OU=WBCAU/O=WBG 
    strUserName = oSession.UserName 
 
    strMailDbName = Left(strUserName, 1) & Right(strUserName, (Len(strUserName) - InStr(1, strUserName, "")))&".nsf" 
    wscript.echo("MailDbName        :" & strMailDbName) 
 
    wscript.echo "## Getting current Notes database..." 
    ' open the mail database in Notes 
 
    set oCurrentMailDb = oSession.CurrentDatabase 
 
    wscript.echo("fileName:" & oCurrentMailDb.fileName) 
    wscript.echo("filePath:" & oCurrentMailDb.filePath) 
    wscript.echo("server:" & oCurrentMailDb.server) 
    wscript.echo("Title:" & oCurrentMailDb.Title) 
 
    If oCurrentMailDb.IsOpen = True Then 
        ' Already open for mail 
        wscript.echo "## Lotus Notes mail database is already open !" 
    Else 
        wscript.echo "## Opening Lotus Notes mail database..." 
        oCurrentMailDb.OPENMAIL 
    End If 
  
    ' Create a document in the back end 
    Set oMailDoc = oCurrentMailDb.CREATEDOCUMENT 
 
    ' Set the form name to memo 
    OMailDoc.form = "Memo"  
 
    with oMailDoc 
        .SendTo = " glezz.alonso@gmail.com"
        .BlindCopyTo = "cglezz.alonso@gmail.com" 
        .CopyTo = "rohs1@schneider-electric.com" 
        .Subject = "Email de Prueba este es el asunto"  
    end with 
 
    set ortItem = oMaildoc.CREATERICHTEXTITEM("Body") 
    with ortItem 
	    .AppendText("texto Contenido en el email")
    End With 
 
    ' Create additional Rich Text item and attach it 
    If blAttachment Then 
        Set ortAttacment = oMailDoc.CREATERICHTEXTITEM("Attachment") 
 
        ' Function EMBEDOBJECT(ByVal TYPE As Short, ByVal CLASS As String, ByVal SOURCE As String, Optional ByVal OBJECTNAME As Object = Nothing) As Object 
        ' Member of lotus.NOTESRICHTEXTITEM 
        Set oEmbedObject = ortAttacment.EMBEDOBJECT(1454, "", cstrAttachment, "Attachment") 
 
    End If 
     
 
    wscript.echo "## Sending email..." 
    with oMailDoc 
        .PostedDate = Now() 
        .SAVEMESSAGEONSEND = "True" 
 
        .send(false) 
    end with 
    wscript.echo "## Sent !" 
 
 
    ' close objects 
    set oMailDoc       = nothing 
    set oCurrentMailDb = nothing 
    set oSession       = nothing