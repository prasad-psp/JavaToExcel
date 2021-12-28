

package javatoexcel;

import java.awt.Component;
import java.io.BufferedOutputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.JTable;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ExportToExcel {
    
    private FileOutputStream excelFOU = null;
    private BufferedOutputStream excelBOU = null;
    private XSSFWorkbook excelJTableExporter = null;
    
    
    private final short START_ROW = 3;
    private final short START_COLUMN = 1;
    
   
    
    private CellStyle setHeadingStyle(XSSFWorkbook excelJTableExporter) {
        try {
            CellStyle style = null;
                                
            XSSFFont font = excelJTableExporter.createFont();
            font.setFontHeightInPoints((short)12);
            font.setFontName("Calibri");
            font.setColor(IndexedColors.BLACK.getIndex());
            font.setBold(true);
            font.setItalic(false);

            style = excelJTableExporter.createCellStyle();
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
            style.setFont(font);
            
            return style;
        }
        catch(Exception e) {
            return null;
        }
    }    
    
    private void createHeading(XSSFSheet excelSheet,XSSFWorkbook excelJTableExporter,JTable table) {
        
        try {
            // Row
            XSSFRow hexcelRow = excelSheet.createRow(START_ROW);
            
            int columnCount = table.getColumnCount();
            
            // column
            for(int i = 0; i < columnCount; i++) {
                String columnName = table.getColumnName(i);
                
                if(!columnName.isEmpty()) {
                    
                    try {
                        XSSFCell excelCell = hexcelRow.createCell(i + START_COLUMN);
                        excelCell.setCellValue(columnName);
                        excelCell.setCellStyle(setHeadingStyle(excelJTableExporter));
                    }
                    catch(Exception e) {
                        JOptionPane.showMessageDialog(null,"Create heading ERROR "+e.getMessage());
                    }
                }
            } 
        }
        catch(Exception e) {
            JOptionPane.showMessageDialog(null,"Create heading ERROR "+e.getMessage());
        }
    }
    
    private void createExcelUsingTable(XSSFSheet excelSheet,JTable table) {
        try {
            int rowCount = table.getRowCount();
            int columnCount = table.getColumnCount();
            
            for(int i = 0; i < rowCount; i++) {
                
                // For row
                XSSFRow excelRow = excelSheet.createRow(i+START_ROW+1);

                // For column
                for(int j = 0; j < columnCount; j++) {
                    
                    XSSFCell excelCell = excelRow.createCell(j+START_COLUMN);
                    excelCell.setCellValue(table.getValueAt(i, j).toString());
                }
            }
        }
        catch(Exception e) {
            JOptionPane.showMessageDialog(null,"Create excel using table ERROR "+e.getMessage());
        }
    }
    
    
    public void export(Component parent,JTable table) {
                
        // Open file chooser
        JFileChooser excelFileChooser = new JFileChooser();
        excelFileChooser.setDialogTitle("Save As");      

        FileNameExtensionFilter fnef = new FileNameExtensionFilter("EXCEL FILES","xls","xlsx","xlsm");
        excelFileChooser.setFileFilter(fnef);
       
        int excelChooser = excelFileChooser.showSaveDialog(parent);
        
        
        if(excelChooser == JFileChooser.APPROVE_OPTION) {
            
            try {
                excelJTableExporter = new XSSFWorkbook();
                XSSFSheet excelSheet = excelJTableExporter.createSheet("Java to excel");

                // For heading
                createHeading(excelSheet,excelJTableExporter,table);
                
                // For data
                createExcelUsingTable(excelSheet,table);
                
                // For write data into excel
                excelFOU = new FileOutputStream(excelFileChooser.getSelectedFile() + ".xlsx");
                excelBOU = new BufferedOutputStream(excelFOU);
                try {
                    excelJTableExporter.write(excelBOU);
                    excelBOU.close();
                    excelFOU.close();
                } 
                catch (IOException ex) {
                    Logger.getLogger("JavaToExcel").log(Level.SEVERE, null, ex);
                }

                JOptionPane.showMessageDialog(null, "Export Successfully");      
                    
            }
            catch (FileNotFoundException ex) {
                Logger.getLogger("JavaToExcel").log(Level.SEVERE, null, ex);
            }
        }
    }    
    
    
}
