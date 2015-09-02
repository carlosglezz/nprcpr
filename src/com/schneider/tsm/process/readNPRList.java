package com.schneider.tsm.process;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import javax.swing.JOptionPane;
import javax.swing.SwingWorker;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING;


public class readNPRList {
    
    public void funcion(String directory)
    {
        final SwingWorker worker;
        worker = new SwingWorker()                                              // Soluciona los problemas en el Stack te trabajo gestionando con threads
        {                                                                       // Evita que la interface Swing se bloquee al ejecutar el proceso
            protected Object doInBackground() throws Exception
            {
                try {
                    AddXPRData createFiles=new AddXPRData();
                    FileInputStream file = new FileInputStream(new File(directory));
                    HSSFWorkbook workbook = new HSSFWorkbook(file);
                    HSSFSheet sheet = workbook.getSheetAt(0);
                    Cell cell = null;
                    int nums=0;
                    int aas=0;
                    String requestStatus="",requestID="",requestorID="",dayFromSubmit="",requestType="",submitDate="";
                    int sheetsize = sheet.getPhysicalNumberOfRows() ;           //Obtiene el Numero de rows en el sheet
                    
                    for (int i=0 ;i<sheetsize;i++)                              // Recorre el Sheet del xls Completo
                    {
                        cell = sheet.getRow(i).getCell(30);
                        if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK) // Verifica si la celda esta vacia
                        {
                            cell= sheet.getRow(i).getCell(37);
                            String leafClass=cell.getStringCellValue();
                            cell = sheet.getRow(i).getCell(35);                 // Se posiciona en la celda de "Request Status"
                            String cellContent = cell.getStringCellValue();     // Obtiene el Valor de la celda actual "Reques Status"
                            requestStatus=cell.getStringCellValue();            // Obtiene el Valor de la celda actual "Reques Status" para el envia a addCPR/NPR
                            
                            if((cellContent.equals("SUBMITTED") || cellContent.equals("ACKNOWLEDGED")
                            || cellContent.equals("ACCEPTED") || cellContent.equals("PART_NUMBER_ASSIGNED"))&& !leafClass.equals("Prototype Only"))// Verifica si se encuentra en los "Request Status" Correctos
                            {
                                
                                System.out.print(i+1 + "  " + cellContent);     //Imprime el Request Status y el numero de Row
                                
                                cell = sheet.getRow(i).getCell(31);
                                aas=(int) cell.getNumericCellValue();
                                System.out.print("// Request ID: "+aas);
                                requestID = Integer.toString(aas);
                                
                                cell = sheet.getRow(i).getCell(0);
                                try {
                                    cellContent=cell.getStringCellValue();
                                    System.out.print("// Requestor: "+cellContent);
                                    requestorID=cellContent;
                                }
                                catch (Exception a){}
                                try {
                                   aas=(int) cell.getNumericCellValue();
                                   System.out.print("// Requestor: "+ aas);
                                   requestorID=Integer.toString(aas);
                                }
                                catch (Exception a){}
                                
                                cell = sheet.getRow(i).getCell(12);
                                requestType=cell.getStringCellValue();
                                System.out.print("// "+requestType +" ");
                                cell = sheet.getRow(i).getCell(30);             // Se posiciona en la celda de "Submit Date"
                                Date olddate =cell.getDateCellValue() ;         //Obtiene el valor de la celda actual "Submit Date" y lo almacena en una variable tipo Date
                                Date today = new Date();                        // Se crea variable tipo Date que contendra la fecha Actual
                                long diff = today.getTime() - olddate.getTime();//diferencia de tiempo entre submit y olddate actual en milisegundos
                                System.out.print("//  " + (diff / (1000 * 60 * 60 * 24)) + " days old." + "// " ); //diferencia de tiempo entre submit y olddate actual en dias
                                dayFromSubmit="" + (diff / (1000 * 60 * 60 * 24));
                                nums++;                                         // Aumenta el contador del numero de NPR/CPR
                                if(Integer.parseInt(dayFromSubmit) > 5 ) {      //Si la fecha desde el "Submit" es mayor a 5 dias
                                createFiles.setRowData(requestStatus,requestID,requestorID,dayFromSubmit,requestType,olddate);//Establece los datos del row actual que seran agregados al excel
                                createFiles.addRow2File();                      //determina a que xls sera agregada la info y la inserta
                                }
                                System.out.println(" "+leafClass);
                            }
                        }
                    }
                    System.out.println("Total NPR in Status:  " + nums);
                    System.out.println("Total NPR:  " +sheetsize);
                    file.close();                                               //cierra el Archivo Excel
                    
                    FileOutputStream outFile =new FileOutputStream(new File(directory)); //lo guarda como "file.xls" archivo de salida
                    workbook.write(outFile);
                    outFile.close();
                    
                }
                catch (FileNotFoundException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                }
                
                return null;
            }
        };
        worker.execute();
    }
}


