

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
    
    private final ArrayList<String> headingList = new ArrayList<>();
    
    private final short START_ROW = 3;
    private final short START_COLUMN = 1;
    
        
    
    
    
    public void addHeading(String heading) {
        
        if(heading != null && !heading.isEmpty()) {
            headingList.add(heading);
        }
    }
    
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
    
    private void createHeading(XSSFSheet excelSheet,XSSFWorkbook excelJTableExporter) {
        
        try {
            // Row
            XSSFRow hexcelRow = excelSheet.createRow(START_ROW);
            
            int headingSize = headingList.size();

            // Columns
            for(int k = 0; k < headingSize; k++) {
                
                try {
                    XSSFCell hexcelCell = hexcelRow.createCell(k + START_COLUMN);
                    hexcelCell.setCellValue(headingList.get(k));
                    hexcelCell.setCellStyle(setHeadingStyle(excelJTableExporter));
                }
                catch(Exception e) {
                    JOptionPane.showMessageDialog(null,"Create heading ERROR "+e.getMessage());
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
                    
                    XSSFCell excelCell = excelRow.createCell(j);
                    excelCell.setCellValue(table.getValueAt(i, j).toString());
                }
            }
        }
        catch(Exception e) {
            JOptionPane.showMessageDialog(null,"Create excel using table ERROR "+e.getMessage());
        }
    }
    
    
    public void export(Component parent) {
                
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
                createHeading(excelSheet,excelJTableExporter);
                
                // For data
//                createExcelUsingTable(excelSheet,JTable);
                
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
