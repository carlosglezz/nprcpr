package com.schneider.tsm.library;

import com.schneider.tsm.maillotus.MailLotus;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

/**
 *
 * @author SESA385697 Carlos Gonzalez Alonso @ Schneider-Electric, MDIC, TSM
 * Team.
 */

public class ReadLibraryAndSend {
    
    String emailTo;
    String emailCC;
    String emailBCC;
    String emailSubject;
    String emailBody;
    String emailAttachment;
    String scriptName;
    
    public void CheckFilesAndSend()                                              //lista los Archivos xls en la Carpeta de Output y los guarda en un string array
    {
        File dir = new File("C:\\softwaretest\\FileOutput");                     //Directorio a listar
        String[] ficheros = dir.list();                                          //Arreglo que guarda la lista de archivos
        if (ficheros == null)                                                    //Si la lista esta vacia entonces no existen archivos
            System.out.println("No hay ficheros en el directorio especificado"); //Imprime falta de Archivos
        else {                                                                   //Si existen entonces
            for (int x=0;x<ficheros.length;x++)                                  //Envia el nombre de Cada archivo listado en el arreglo uno por uno
            {
                ifexist(ficheros[x]);                                            //Se envia la lista de Archivos a la funcion ifexist para su busqueda en                                                             
            }                                                                    // el xls de relaciones entre TSE y Managers
        }
        
        File dirTSEMan = new File("C:\\softwaretest\\FileOutput\\TSE_Manager");
        
        
    }
    
    public void ifexist(String idnumber)
    {
        
        try {
            String cellvalue="";                                                 //String se usara para asignar el valor de las celdas
            FileInputStream file = new FileInputStream(new File("C:\\softwaretest\\LibraryTest.xls")); // xls de relaciones entre TSE y Managers
            HSSFWorkbook workbook = new HSSFWorkbook(file);                      //declara el libro de trabajo
            HSSFSheet sheet = workbook.getSheetAt(0);                            //declara la hoja de trabajo
            Cell cell = null;                                                    //declara la celda
            int sheetsize = sheet.getPhysicalNumberOfRows() ;                    //obtiene el numero desde 1 hasta N de Rows
            for (int i = 1; i < sheetsize; i++) {                                //Recorre cada row
                cell = sheet.getRow(i).getCell(0);                               //se posiciona en la celda deseada (TSE SESA)
                cellvalue=cell.getStringCellValue();                             //se asigna el valor del texto de la celda a cellvalue
                scriptName=cellvalue;
                if (idnumber.endsWith(cellvalue + ".xls")) {                     //verifica si el Valor de la celda(SESA) es igual al nombre del archivo
                 
                    emailAttachment="C:\\softwaretest\\FileOutput\\"+idnumber;
                    MailLotus createscript = new MailLotus();
                    cell = sheet.getRow(i).getCell(1);
                    emailTo=cell.getStringCellValue();
                    cell = sheet.getRow(i).getCell(3);
                    emailCC=cell.getStringCellValue();
                    cell = sheet.getRow(i).getCell(5);
                    emailBCC=cell.getStringCellValue();
                    emailSubject="TSE "+ scriptName +" Notification: Timeout CPR  & NPR Status";
                    emailBody="You have one or more CPR/NPR timeout items. Please find below attachment the listed items";
                    createscript.setMailData(emailTo, emailCC, emailBCC, emailSubject, emailBody, emailAttachment, scriptName);
                    createscript.createScript();
                    System.out.println(emailTo +",  "+ emailCC +",  " +emailBCC +",  "+ emailSubject +",  "+ emailBody +",  "+ emailAttachment +",  "+ scriptName);
                    
//  System.out.println("ok for:" + cellvalue);                      
                }
            }
            file.close();
            FileOutputStream outFile =new FileOutputStream(new File("C:\\softwaretest\\LibraryTest.xls"));
            workbook.write(outFile);
            outFile.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
}
