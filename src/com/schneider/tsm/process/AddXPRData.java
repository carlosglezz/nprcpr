package com.schneider.tsm.process;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import javax.swing.SwingWorker;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import static org.apache.poi.ss.usermodel.CellStyle.ALIGN_CENTER;

/**
 *
 * @author SESA385697 Carlos Gonzalez Alonso @ Schneider-Electric, MDIC, TSM
 * Team.
 */

public class AddXPRData {
    
    String requestStatus="",requestID="",requestorID="",dayFromSubmit="",requestType="";
    Date submitDate;
    
    
    void setRowData(String requestStatus, String requestID, String requestorID, String dayFromSubmit,String requestType,Date submitDate) {
        this.dayFromSubmit=dayFromSubmit;
        this.requestID=requestID;
        this.requestStatus=requestStatus;
        this.requestorID=requestorID;
        this.requestType=requestType;
        this.submitDate=submitDate;
        
    }
    public boolean ifFileExist()
    {
        String path="C:\\softwaretest\\FileOutput\\NPR_CPR_Report_"+ requestorID +".xls";
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
    
    
    public void add2ExistingReport() throws InterruptedException
    {
        
        final SwingWorker worker = new SwingWorker() {
            @Override
            protected Object doInBackground() throws Exception {
                
                try {
                    addXPRDataTSEAlert dataTSEAlert =new addXPRDataTSEAlert();
                    addXPRDataManager dataManager =new addXPRDataManager();
                    addXPRDataQuality dataQuality =new addXPRDataQuality();
                    FileInputStream file = new FileInputStream(new File("C:\\softwaretest\\FileOutput\\NPR_CPR_Report_"+ requestorID +".xls"));
                    HSSFWorkbook workbook = new HSSFWorkbook(file);
                    HSSFSheet sheet = workbook.getSheetAt(0);
                    int sheetsize = sheet.getPhysicalNumberOfRows() ;
                    Cell cell = null;
                    int numm=1;
                    
                    for (int i = 7; i < sheetsize; i++) {
                        cell = sheet.getRow(i).getCell(1);
                        if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK)
                        {
                            HSSFCellStyle cellStyle = workbook.createCellStyle();
                            cellStyle=workbook.createCellStyle();
                            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                            cellStyle.setAlignment(ALIGN_CENTER);
                            if(Integer.parseInt(dayFromSubmit) < 5)
                                cellStyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
                            
                            if(Integer.parseInt(dayFromSubmit) >= 5 && Integer.parseInt(dayFromSubmit) <= 19)
                                cellStyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
                            
                            if((Integer.parseInt(dayFromSubmit) > 19 && requestType.equals("NPR")) ||(Integer.parseInt(dayFromSubmit) > 12 && requestType.equals("CPR")) )
                                cellStyle.setFillForegroundColor(HSSFColor.GOLD.index);
                            
                            if(Integer.parseInt(dayFromSubmit) > 40  )
                                cellStyle.setFillForegroundColor(HSSFColor.ORANGE.index);
                            
                            
                            if(Integer.parseInt(dayFromSubmit) > 90  )
                                cellStyle.setFillForegroundColor(HSSFColor.RED.index);
                            
                            //  cell.setCellStyle(cellStyle);
                            cell.setCellValue(""+numm);
                            cell = sheet.getRow(i).getCell(2);
                            //  cell.setCellStyle(cellStyle);
                            cell.setCellValue(requestID);
                            cell = sheet.getRow(i).getCell(3);
                            //  cell.setCellStyle(cellStyle);
                            cell.setCellValue(requestorID);
                            cell = sheet.getRow(i).getCell(4);
                            cell.setCellStyle(cellStyle);
                            cell.setCellValue(dayFromSubmit);
                            cell = sheet.getRow(i).getCell(5);
                            //   cell.setCellStyle(cellStyle);
                            cell.setCellValue(requestStatus);
                            cell = sheet.getRow(i).getCell(6);
                            //    cell.setCellStyle(cellStyle);
                            cell.setCellValue(submitDate);
                            cell = sheet.getRow(i).getCell(7);
                            //    cell.setCellStyle(cellStyle);
                            cell.setCellValue(requestType);
                            
                            i=sheetsize;
                            
                        }
                        numm++;
                    }
                    
                    file.close();
                    FileOutputStream outFile =new FileOutputStream(new File("C:\\softwaretest\\FileOutput\\NPR_CPR_Report_"+ requestorID +".xls"));
                    workbook.write(outFile);
                    outFile.close();
                    Thread.sleep(160);
                    if((Integer.parseInt(dayFromSubmit) > 19 && requestType.equals("NPR") && Integer.parseInt(dayFromSubmit)<= 40) ||(Integer.parseInt(dayFromSubmit) > 12 && requestType.equals("CPR") && Integer.parseInt(dayFromSubmit)<= 40 ) )//to TSE cc Manager
                    {
                        dataTSEAlert.setRowData(requestStatus,requestID,requestorID,dayFromSubmit,requestType,submitDate);//Establece los datos del row actual que seran agregados al excel
                        dataTSEAlert.addRow2File();
                    }
                    
                    if(Integer.parseInt(dayFromSubmit) > 40  &&  Integer.parseInt(dayFromSubmit) <= 90)//to Manager cc TSE
                    {
                        dataManager.setRowData(requestStatus,requestID,requestorID,dayFromSubmit,requestType,submitDate);//Establece los datos del row actual que seran agregados al excel
                        dataManager.addRow2File();
                    }
                    
                    if(Integer.parseInt(dayFromSubmit) > 90  )//to leslie cc Manager
                    {
                        dataQuality.setRowData(requestStatus,requestID,requestorID,dayFromSubmit,requestType,submitDate);//Establece los datos del row actual que seran agregados al excel
                        dataQuality.addRow2File();
                    }
                    
                    
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                }
                
                return null;
            }
        };
        worker.execute();
        Thread.sleep(260);
    }
    
    private void createNewFileReport() throws InterruptedException
    {
        final SwingWorker worker = new SwingWorker() {
            @Override
            protected Object doInBackground() throws Exception {
                
                try {
                    addXPRDataTSEAlert dataTSEAlert =new addXPRDataTSEAlert();
                    addXPRDataManager dataManager =new addXPRDataManager();
                    addXPRDataQuality dataQuality =new addXPRDataQuality();
                    FileInputStream file = new FileInputStream(new File("C:\\softwaretest\\template.xls"));
                    HSSFWorkbook workbook = new HSSFWorkbook(file);
                    HSSFSheet sheet = workbook.getSheetAt(0);
                    Cell cell = null;
                    
                    HSSFCellStyle cellStyle = workbook.createCellStyle();
                    cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                    cellStyle.setAlignment(ALIGN_CENTER);
                    
                    if(Integer.parseInt(dayFromSubmit) < 5)
                        cellStyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
                    
                    if(Integer.parseInt(dayFromSubmit) >= 5 && Integer.parseInt(dayFromSubmit) <= 19)
                        cellStyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
                    
                    if((Integer.parseInt(dayFromSubmit) > 19 && requestType.equals("NPR")) ||(Integer.parseInt(dayFromSubmit) > 12 && requestType.equals("CPR")) )
                        cellStyle.setFillForegroundColor(HSSFColor.GOLD.index);
                    
                    if(Integer.parseInt(dayFromSubmit) > 40  )
                        cellStyle.setFillForegroundColor(HSSFColor.ORANGE.index);
                    
                    
                    if(Integer.parseInt(dayFromSubmit) > 90  )
                        cellStyle.setFillForegroundColor(HSSFColor.RED.index);
                    
                    cell = sheet.getRow(7).getCell(1);
                    //  cell.setCellStyle(cellStyle);
                    cell.setCellValue("1");
                    cell = sheet.getRow(7).getCell(2);
                    //  cell.setCellStyle(cellStyle);
                    cell.setCellValue(requestID);
                    cell = sheet.getRow(7).getCell(3);
                    // cell.setCellStyle(cellStyle);
                    cell.setCellValue(requestorID);
                    cell = sheet.getRow(7).getCell(4);
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue(dayFromSubmit);
                    cell = sheet.getRow(7).getCell(5);
                    //   cell.setCellStyle(cellStyle);
                    cell.setCellValue(requestStatus);
                    cell = sheet.getRow(7).getCell(6);
                    //   cell.setCellStyle(cellStyle);
                    cell.setCellValue(submitDate);
                    cell = sheet.getRow(7).getCell(7);
                    //   cell.setCellStyle(cellStyle);
                    cell.setCellValue(requestType);
                    
                    
                    file.close();
                    FileOutputStream outFile =new FileOutputStream(new File("C:\\softwaretest\\FileOutput\\NPR_CPR_Report_"+ requestorID +".xls"));
                    workbook.write(outFile);
                    outFile.close();
                     Thread.sleep(160);
                    if((Integer.parseInt(dayFromSubmit) > 19 && requestType.equals("NPR") && Integer.parseInt(dayFromSubmit)<= 40) ||(Integer.parseInt(dayFromSubmit) > 12 && requestType.equals("CPR") && Integer.parseInt(dayFromSubmit)<= 40 ) )//to TSE cc Manager
                    {
                        dataTSEAlert.setRowData(requestStatus,requestID,requestorID,dayFromSubmit,requestType,submitDate);//Establece los datos del row actual que seran agregados al excel
                        dataTSEAlert.addRow2File();
                    }
                    
                    if(Integer.parseInt(dayFromSubmit) > 40 &&  Integer.parseInt(dayFromSubmit) < 90)//to Manager cc TSE
                    {
                        dataManager.setRowData(requestStatus,requestID,requestorID,dayFromSubmit,requestType,submitDate);//Establece los datos del row actual que seran agregados al excel
                        dataManager.addRow2File();
                    }
                    
                    if(Integer.parseInt(dayFromSubmit) > 90  )//to leslie cc Manager
                    {
                        dataQuality.setRowData(requestStatus,requestID,requestorID,dayFromSubmit,requestType,submitDate);//Establece los datos del row actual que seran agregados al excel
                        dataQuality.addRow2File();
                    }
                    
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                }
                
                return null;
            }
        };
        worker.execute();
        Thread.sleep(260);
    }
}
