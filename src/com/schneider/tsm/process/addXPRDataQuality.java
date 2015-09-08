/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.schneider.tsm.process;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import javax.swing.SwingWorker;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

/**
 *
 * @author SESA385697 Carlos Gonzalez Alonso @ Schneider-Electric, MDIC, TSM
 * Team.
 */
class addXPRDataQuality {

    String requestStatus="",requestID="",requestorID="",dayFromSubmit="",requestType="";
    Date submitDate;
    String requestorQuality="";
    String requestorManager="";
    
    
    void setRowData(String requestStatus, String requestID, String requestorID, String dayFromSubmit,String requestType,Date submitDate) {
        this.dayFromSubmit=dayFromSubmit;
        this.requestID=requestID;
        this.requestStatus=requestStatus;
        this.requestorID=requestorID;
        this.requestType=requestType;
        this.submitDate=submitDate;
        requestorManager=getManager();
        requestorQuality=getQuality();
        
    }
    public boolean ifFileExist()
    {
        String path="C:\\softwaretest\\FileOutput\\Quality\\NPR_CPR_Report_"+ requestorID +".xls";
        File xlsFile= new File(path);
        return xlsFile.exists();
    }
    
    void addRow2File() throws InterruptedException
    {
        if(ifFileExist())
        {
            add2ExistingReport();
        }
        else
        {
            createNewFileReport();
        }
    }

    private void add2ExistingReport() throws InterruptedException {
           final SwingWorker worker = new SwingWorker() {
            @Override
            protected Object doInBackground() throws Exception {  
                try {
                    FileInputStream file = new FileInputStream(new File("C:\\softwaretest\\FileOutput\\Quality\\NPR_CPR_Report_"+ requestorQuality +".xls"));
                    HSSFWorkbook workbook = new HSSFWorkbook(file);
                    HSSFSheet sheet = workbook.getSheetAt(0);
                    int sheetsize = sheet.getPhysicalNumberOfRows() ;
                    Cell cell = null;
                    int numm=1;
                    
                    for (int i = 7; i < sheetsize; i++) {
                        cell = sheet.getRow(i).getCell(1);
                        if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK)
                        {
                            cell.setCellValue(""+numm);
                            cell = sheet.getRow(i).getCell(2);
                            cell.setCellValue(requestID);
                            cell = sheet.getRow(i).getCell(3);
                            cell.setCellValue(requestorID);
                            cell = sheet.getRow(i).getCell(4);
                            cell.setCellValue(dayFromSubmit);
                            cell = sheet.getRow(i).getCell(5);
                            cell.setCellValue(requestStatus);
                            cell = sheet.getRow(i).getCell(6);
                            cell.setCellValue(submitDate);
                            cell = sheet.getRow(i).getCell(7);
                            cell.setCellValue(requestType);
                            i=sheetsize;
                        }
                        numm++;
                    }
                    
                    file.close();
                    FileOutputStream outFile =new FileOutputStream(new File("C:\\softwaretest\\FileOutput\\Quality\\NPR_CPR_Report_"+ requestorQuality +".xls"));
                    workbook.write(outFile);
                    outFile.close();
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                }
                
                return null;
            }
        };
        worker.execute();
        Thread.sleep(160);

         }

    private void createNewFileReport() throws InterruptedException {
           
        final SwingWorker worker = new SwingWorker() {
            @Override
            protected Object doInBackground() throws Exception {
                
                try {
                    FileInputStream file = new FileInputStream(new File("C:\\softwaretest\\template.xls"));
                    HSSFWorkbook workbook = new HSSFWorkbook(file);
                    HSSFSheet sheet = workbook.getSheetAt(0);
                    Cell cell = null;           
                    cell = sheet.getRow(7).getCell(1);
                    cell.setCellValue("1");
                    cell = sheet.getRow(7).getCell(2);
                    cell.setCellValue(requestID);
                    cell = sheet.getRow(7).getCell(3);
                    cell.setCellValue(requestorID);
                    cell = sheet.getRow(7).getCell(4);
                    cell.setCellValue(dayFromSubmit);
                    cell = sheet.getRow(7).getCell(5);
                    cell.setCellValue(requestStatus);
                    cell = sheet.getRow(7).getCell(6);
                    cell.setCellValue(submitDate);
                    cell = sheet.getRow(7).getCell(7);
                    cell.setCellValue(requestType);
                    file.close();
                    FileOutputStream outFile =new FileOutputStream(new File("C:\\softwaretest\\FileOutput\\Quality\\NPR_CPR_Report_"+ requestorQuality +".xls"));
                    workbook.write(outFile);
                    outFile.close();
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                }
                
                return null;
            }
        };
        worker.execute();
        Thread.sleep(160);
    }


    private String getManager() {
        String manager=null;
        
        try {
            FileInputStream file = new FileInputStream(new File("C:\\softwaretest\\LibraryTest.xls"));
            HSSFWorkbook workbook = new HSSFWorkbook(file);
            HSSFSheet sheet = workbook.getSheetAt(0);
            Cell cell = null;
            int sheetsize = sheet.getPhysicalNumberOfRows();
            for (int i=1;i<sheetsize;i++){
                cell = sheet.getRow(i).getCell(0);
                if (cell.getStringCellValue().equals(requestorID))
                {
                    cell = sheet.getRow(i).getCell(2);
                    manager= cell.getStringCellValue();
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
        
        return manager;
        
    }

    private String getQuality() {
             String Quality="No_Exist";
        
        try {
            FileInputStream file = new FileInputStream(new File("C:\\softwaretest\\LibraryTest.xls"));
            HSSFWorkbook workbook = new HSSFWorkbook(file);
            HSSFSheet sheet = workbook.getSheetAt(0);
            Cell cell = null;
            int sheetsize = sheet.getPhysicalNumberOfRows();
            for (int i=1;i<sheetsize;i++){
                cell = sheet.getRow(i).getCell(0);
                if (cell.getStringCellValue().equals(requestorID))
                {
                    cell = sheet.getRow(i).getCell(4);
                    Quality= cell.getStringCellValue();
                }
            }
            file.close();
            FileOutputStream outFile =new FileOutputStream(new File("C:\\softwaretest\\LibraryTest.xls"));
            workbook.write(outFile);
            outFile.close();
        } catch (FileNotFoundException e) {
        } catch (IOException e) {
        }
        
        return Quality;
    }
}
