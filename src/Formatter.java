
import org.apache.poi.ss.usermodel.*;
import java.io.*;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author OTR2
 */
public class Formatter {
    private Workbook wb;
    private Sheet worksheet;
    
    public Formatter(File file) {
        FileInputStream ins = null;
        try {
            ins = new FileInputStream(file);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(Formatter.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        try {
            wb = WorkbookFactory.create(ins);
        } catch (IOException | InvalidFormatException | EncryptedDocumentException ex) {
            Logger.getLogger(Formatter.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        worksheet = wb.getSheetAt(0);
    }
    
    public Workbook getWorkbook() {
        return wb;
    }
    
    public Sheet getWorksheet() {
        return worksheet;
    }
    
    public void format() {
        int i = 0;
        int j = 1;

        while(worksheet.getRow(0).getCell(i) != null && !worksheet.getRow(0).getCell(i).toString().equals("")){
            i++;
        }
        
        if(worksheet.getRow(0).getCell(i) == null) {
            worksheet.getRow(0).createCell(i).setCellValue("Documents");
            worksheet.getRow(0).createCell(++i).setCellValue("Plaintiff");
            worksheet.getRow(0).createCell(++i).setCellValue("Defendant");
            
        }
        
        int numCol = i;
        i = 0;
        
        while(worksheet.getRow(j).getCell(0) != null) {
            i = 0;
            while(worksheet.getRow(j).getCell(i) != null) {  
                if(worksheet.getRow(0).getCell(i).toString().equalsIgnoreCase("Home Addy")
                        || worksheet.getRow(0).getCell(i).toString().equalsIgnoreCase("Home City")
                        || worksheet.getRow(0).getCell(i).toString().equalsIgnoreCase("Style")
                        || worksheet.getRow(0).getCell(i).toString().equalsIgnoreCase("Def")) {
                    
                    String[] temp = worksheet.getRow(j).getCell(i).toString().split(" ");
                    
                    String outString = "";
                    for(String it : temp) {
                        if(!it.equals(""))
                            outString += it.substring(0, 1).toUpperCase() + it.substring(1).toLowerCase();
                        outString += " ";
                    }
                    worksheet.getRow(j).getCell(i).setCellValue(outString);
                }
                
                //Check if the court is either superior or justice court
                if(worksheet.getRow(j).getCell(i).toString().toLowerCase().contains("superior court"))
                    worksheet.getRow(j).createCell(numCol-2).setCellValue("Summon, Complaint and Certificate of Compulsory Arbitration");
                if(worksheet.getRow(j).getCell(i).toString().toLowerCase().contains("justice court"))
                    worksheet.getRow(j).createCell(numCol-2).setCellValue("Summon, Notice to the Defendants, and Complaint");

                //change the courts to their respective codes
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Superior Court - Central"))
                    worksheet.getRow(j).getCell(i).setCellValue(500);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Agua Fria Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(501);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Pinal County Superior Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(2);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Santa Cruz County Superior Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(640);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Pima County Superior Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(4);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("San Marcos Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(520);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("North Valley Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(519);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Moon Valley Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(517);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Arcadia Biltmore Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(502);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("San Tan Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(521);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Hassayampa Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(510);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Flagstaff Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(551);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Pima County Consolidated Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(621);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Manistee Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(514);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Justice Court - Precinct 7"))
                    worksheet.getRow(j).getCell(i).setCellValue(638);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Justice Court - Precinct 8"))
                    worksheet.getRow(j).getCell(i).setCellValue(637);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Prescott Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(653);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Dreamy Draw Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(506);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("McDowell Mountain Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(516);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("East Mesa Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(507);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Highland Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(511);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Maryvale Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(515);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Justice Court - Precinct 1"))
                    worksheet.getRow(j).getCell(i).setCellValue(631);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Lake Havasu Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(603);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Yuma Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(661);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("West Mesa Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(525);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Payson Regional Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(562);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Sierra Vista Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue("SVJC");
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Arrowhead Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(503);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("South Mountain Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(522);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Douglas Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(542);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Encanto Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(508);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Bowie Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(546);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("White Tank Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(509);
                if(worksheet.getRow(j).getCell(i).toString().equalsIgnoreCase("Kyrene Justice Court"))
                    worksheet.getRow(j).getCell(i).setCellValue(513);
                
              
                if(worksheet.getRow(0).getCell(i).toString().equalsIgnoreCase("STYLE")) {
                    String[] temp = worksheet.getRow(j).getCell(i).toString().split("Vs. ");
                    worksheet.getRow(j).createCell(numCol-1).setCellValue(temp[0]);
                    worksheet.getRow(j).createCell(numCol).setCellValue(temp[1]);
                }
                
                
                i++;
            }
            j++;
        }
    }
    //end format
    
    public void save (File file) {
        try {
            wb.write(new FileOutputStream(file));
        } catch (FileNotFoundException ex) {
            Logger.getLogger(Formatter.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(Formatter.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}
