package com.schneider.tsm.maillotus;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;

/**
 *
 * @author SESA385697 Carlos Gonzalez Alonso @ Schneider-Electric, MDIC, TSM
 * Team.
 */

public class MailLotus {
    
    String emailTo;
    String emailCC;
    String emailBCC;
    String emailSubject;
    String emailBody;
    String emailAttachment;
    String scriptName;
    
    
    public void setMailData(String emailTo,String emailCC, String emailBCC,String emailSubject,String emailBody,String emailAttachment,String scriptName)
    {
        this.emailTo=emailTo;
        this.emailCC=emailCC;
        this.emailBCC=emailBCC;
        this.emailSubject=emailSubject;
        this.emailBody=emailBody;
        this.emailAttachment=emailAttachment;
        this.scriptName=scriptName;
        
    }
    
    public void sendMail(String ScriptToSent) throws InterruptedException
    {
        try
        {
            Process p = Runtime.getRuntime().exec("cmd.exe  /C "+ "C:\\scripting\\" + ScriptToSent );
            System.out.println("cmd.exe "+ "C:\\scripting\\" + ScriptToSent);
            System.out.println("Mail Sended");
            Thread.sleep(1000);
        }
        catch (IOException ioe)
        {
            System.out.println (ioe);
        }
    }
    
    public void createScript()
    {
        
        
        File theDir = new File("C:\\scripting");
        
        // if the directory does not exist, create it
        if (!theDir.exists()) {
            System.out.println("creating directory: C:\\scripting");
            boolean result = false;
            
            try{
                theDir.mkdir();
                result = true;
            }
            catch(SecurityException se){
                //handle it
            }
            if(result) {
                System.out.println("DIR created");
            }
        }
        String scriptVBS = "option explicit \n" +
                " \n" +
                "\n" +
                " \n" +
                "    Dim oSession        ' AS NotesSession \n" +
                "    Dim strServer \n" +
                "    Dim strUserName \n" +
                "    Dim strMailDbName \n" +
                "    Dim oCurrentMailDb  ' as NOTESDATABASE \n" +
                "    Dim oMailDoc        ' as NOTESDOCUMENT \n" +
                "    Dim ortItem         ' as NOTESRICHTEXTITEM \n" +
                "    Dim ortAttacment    ' as NOTESRICHTEXTITEM \n" +
                "    Dim oEmbedObject    ' as ???? \n" +
                "    dim cstrAttachment \n" +
                "    Dim blAttachment \n" +
                " \n" +
                "    cstrAttachment =  \""+emailAttachment+"\"\n" +
                " \n" +
                "    blAttachment = True \n" +
                " \n" +
                "    ' Start a session to notes \n" +             
                "    Set oSession = CreateObject(\"Notes.NotesSession\") \n" +
                " \n" +
                " \n" +
                "    strServer = oSession.GetEnvironmentString(\"MailServer\",True) \n" +
                " \n" +
                "    ' eg. CN=Michael V Green/OU=CORPAU/OU=WBCAU/O=WBG \n" +
                "    strUserName = oSession.UserName \n" +
                " \n" +
                "    strMailDbName = Left(strUserName, 1) & Right(strUserName, (Len(strUserName) - InStr(1, strUserName, \"\")))&\".nsf\" \n" +
                " \n" +
                "    ' open the mail database in Notes \n" +
                " \n" +
                "    set oCurrentMailDb = oSession.CurrentDatabase \n" +
                " \n" +
                " \n" +
                "    If oCurrentMailDb.IsOpen = True Then \n" +
                "        ' Already open for mail \n" +
                "    Else \n" +
                "        oCurrentMailDb.OPENMAIL \n" +
                "    End If \n" +
                "  \n" +
                "    ' Create a document in the back end \n" +
                "    Set oMailDoc = oCurrentMailDb.CREATEDOCUMENT \n" +
                " \n" +
                "    ' Set the form name to memo \n" +
                "    OMailDoc.form = \"Memo\"  \n" +
                " \n" +
                "    with oMailDoc \n" +
                "        .SendTo = \" "+ emailTo +"\"\n" +
                "        .BlindCopyTo = \""+ emailBCC + "\" \n" +
                "        .CopyTo = \""  + emailCC + "\" \n" +
                "        .Subject = \""  + emailSubject +  "\"  \n" +
                "    end with \n" +
                " \n" +
                "    set ortItem = oMaildoc.CREATERICHTEXTITEM(\"Body\") \n" +
                "    with ortItem \n" +
                "	    .AppendText(\"" + emailBody +"\")\n" +
                "    End With \n" +
                " \n" +
                "    ' Create additional Rich Text item and attach it \n" +
                "    If blAttachment Then \n" +
                "        Set ortAttacment = oMailDoc.CREATERICHTEXTITEM(\"Attachment\") \n" +
                " \n" +
                "        ' Function EMBEDOBJECT(ByVal TYPE As Short, ByVal CLASS As String, ByVal SOURCE As String, Optional ByVal OBJECTNAME As Object = Nothing) As Object \n" +
                "        ' Member of lotus.NOTESRICHTEXTITEM \n" +
                "        Set oEmbedObject = ortAttacment.EMBEDOBJECT(1454, \"\", cstrAttachment, \"Attachment\") \n" +
                " \n" +
                "    End If \n" +
                "     \n" +
                " \n" +
              // "    wscript.echo \"## Sending email...\" \n" +
                "    with oMailDoc \n" +
                "        .PostedDate = Now() \n" +
                "        .SAVEMESSAGEONSEND = \"True\" \n" +
                " \n" +
                "        .send(false) \n" +
                "    end with \n" +
            //   "    wscript.echo \"## Sent !\" \n" +
                " \n" +
                " \n" +
                "    ' close objects \n" +
                "    set oMailDoc       = nothing \n" +
                "    set oCurrentMailDb = nothing \n" +
                "    set oSession       = nothing";
        try
        {
            //Crear un objeto File se encarga de crear o abrir acceso a un archivo que se especifica en su constructor
            File archivo=new File("C:\\scripting\\"+scriptName+".vbs");
            //Crear objeto FileWriter que sera el que nos ayude a writeScript sobre archivo
            FileWriter writeScript=new FileWriter(archivo,true);
            writeScript.write(scriptVBS);//Escribimos en el archivo con el metodo write
            writeScript.close(); //Cerramos la conexion
            System.out.println("Successful file creation:" + scriptName);
        }
        catch(Exception e)
        {
            System.out.println("Error al escribir");
        }
    }
    
}
