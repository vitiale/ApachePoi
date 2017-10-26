/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package apachepoi;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 *
 * @author Alba Proyecto
 */
public class ApachePoi {

    /**
     * @param args the command line arguments
     */
    
    public void readWriteExcelFile(File excelFile, File excelNewFile){
        InputStream excelStream=null;
        OutputStream excelNewOutputStream=null;
        try {
            HSSFWorkbook hssfWorkbook =new HSSFWorkbook(excelStream);
            HSSFWorkbook hssfWorkbookNew =new HSSFWorkbook();
            
        } catch (Exception e) {
        }
    }
    
    public static void main(String[] args) {
        // TODO code application logic here
    }
    
}
