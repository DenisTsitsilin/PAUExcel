package vstu.pauexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

/**
 * @author Tsitsilin Denis
 */
public class ReportExcel {
public void writeIntoExcel(ArrayList <String> list){
        try {
            HSSFWorkbook book = new HSSFWorkbook(getFileInputStreamWithFile());
            /**
             * Поменять на правильное название листа
             */
            HSSFSheet sheet = book.getSheet("List");
            int rowCount = 1;
            for (String value : list) {
                // Нумерация начинается с нуля
                Row row = sheet.getRow(rowCount);
                
                PasswordGenerator passwordGenerator = new PasswordGenerator.PasswordGeneratorBuilder()
                                                                           .useDigits(true)
                                                                           .useLower(true)
                                                                           .useUpper(true)
                                                                           .build();
                String password = passwordGenerator.generate(8);
        
                row.getCell(0).setCellValue(password);
                
            }
            book.close();
        } catch (IOException iOException) {
            JOptionPane.showMessageDialog(null, iOException);
        }
    }
    
    public FileInputStream getFileInputStreamWithFile(){
        try {
            JFileChooser fileopen = new JFileChooser();
            fileopen.showDialog(null, "Выберете шаблон EXCELL для выгрузки данных");            
            File inputFile = fileopen.getSelectedFile();
            FileInputStream fis = new FileInputStream(inputFile);
            return fis;            
        } catch (IOException iOException) {
            JOptionPane.showMessageDialog(null, iOException); 
            return null; 
        }
    }
}
