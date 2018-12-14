package vstu.pauexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * @author Tsitsilin Denis
 */
public class ReportExcel {
public void writeIntoExcel(){
        try {
            HSSFWorkbook book = (HSSFWorkbook) WorkbookFactory.create(getFile());
            HSSFSheet sheet = book.getSheet("List");
                
                Iterator<Row> ri = sheet.rowIterator();

                while(ri.hasNext()) {
                    HSSFRow rowa = (HSSFRow) ri.next();
                    PasswordGenerator passwordGenerator = new PasswordGenerator.PasswordGeneratorBuilder()
                                                                           .useDigits(true)
                                                                           .useLower(true)
                                                                           .useUpper(true)
                                                                           .build();
                    String password = passwordGenerator.generate(8);
                    rowa.getCell(1).setCellValue(password);
                }
                FileOutputStream fileOut = new FileOutputStream("C:\\workbook.xls");
                book.write(fileOut);
                fileOut.close();
        } catch (IOException iOException) {
            JOptionPane.showMessageDialog(null, iOException);
        }
    }
    
    public File getFile(){
        try {
            JFileChooser fileopen = new JFileChooser();
            fileopen.showDialog(null, "ОК");            
            File inputFile = fileopen.getSelectedFile();
            return inputFile;            
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, ex); 
            return null; 
        }
    }
}
