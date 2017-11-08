/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package apachepoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Date;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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
    
    //Crear libro en blanco
    public static void libro_blanco() throws FileNotFoundException, IOException{
        File file=new File("blanco.xlsx");
        FileOutputStream out=new FileOutputStream(file);
        XSSFWorkbook libro_blanco=new XSSFWorkbook();
        XSSFSheet hoja=libro_blanco.createSheet("Primera hoja");
        libro_blanco.write(out);
        System.out.println("Se ha creado el libro excel en blanco correctamente");
        out.close();
    }
    
    //Abrir un libro ya creado
    public static void abrir_libro() throws FileNotFoundException, IOException{
        File file=new File("PaisesIdiomasMonedas.xlsx");
        FileInputStream file_input=new FileInputStream(file);
        XSSFWorkbook libro_abierto=new XSSFWorkbook(file_input);
        if(file.isFile() && file.exists()){
            System.out.println("Fichero abierto correctamente");
        }else{
            System.out.println("No se ha podido abriri correctamente el fichero");
        }
        file_input.close();
    }
    
    //Crear un libro y añadirle datos en una hoja de calculo
    public static void escribir_hoja_calc() throws IOException{
        File file = new File("Hoja_calc.xlsx");
        FileOutputStream out=new FileOutputStream(file);
        XSSFWorkbook libro=new XSSFWorkbook();
        //creamos hoja blanco
        XSSFSheet hoja=libro.createSheet("Informacion");
        //creamos un objeto row(fila)
        XSSFRow row;
        //escribir los datos
        Map<String, Object[]> map=new TreeMap<String, Object[]>();
        map.put("1", new Object[]{"NOMBRE","EDAD","SEXO"});
        map.put("2", new Object[]{"Pedro","29","M"});
        map.put("3", new Object[]{"Oscar","26","M"});
        map.put("4", new Object[]{"María","23","F"});
        Set<String> llaves=map.keySet();
        int row_id=0;
        for(String llave:llaves){
            //esto es lo que va a escribir
            System.out.println(llave+"  "+map.get(llave)[0]+"           "+map.get(llave)[1]+"                       "+map.get(llave)[2]);
            row=hoja.createRow(row_id++);
            int cell_id=0;
            for(Object elemento: map.get(llave)){
                Cell cell=row.createCell(cell_id++);
                cell.setCellValue((String)elemento);
            }            
        }
        libro.write(out);
        out.close();        
    }
    
    //leer de una hoja de calculo
    public static void leer_hoja_calc() throws FileNotFoundException, IOException{
        FileInputStream in= new FileInputStream(new File("Hoja_calc.xlsx"));
        XSSFWorkbook libro=new XSSFWorkbook(in);
        XSSFSheet hoja=libro.getSheetAt(0);
        XSSFRow row;
        Iterator<Row> iterar_filas=hoja.iterator();
        while(iterar_filas.hasNext()){
            row=(XSSFRow)iterar_filas.next();
            Iterator<Cell>iterar_cell=row.cellIterator();
            Cell cell   ;
            while(iterar_cell.hasNext()){
                cell=iterar_cell.next();
                System.out.println(""+cell);
            }
        }
        in.close();
    }
    
    //diferentes tipos de celdas en una hoja
    public static void diferentes_tipos_celdas()/* throws FileNotFoundException, IOException*/ {
        FileOutputStream out;
        try {
            out = new FileOutputStream(new File("diferentes_tipos de celdas.xlsx"));

            XSSFWorkbook libro = new XSSFWorkbook();

            XSSFSheet hoja = libro.createSheet("cel types");

            XSSFRow row = hoja.createRow((short) 2);
            row.createCell(0).setCellValue("se le pone valor");
            row.createCell(1).setCellValue("cel value");

            row = hoja.createRow((short) 3);
            row.createCell(0).setCellValue("cel en blanco");
            row.createCell(1);

            row = hoja.createRow((short) 4);
            row.createCell(0).setCellValue("se le pone boolean");
            row.createCell(1).setCellValue(true);

            row = hoja.createRow((short) 5);
            row.createCell(0).setCellValue("celda de error");
            row.createCell(1).setCellValue(XSSFCell.CELL_TYPE_ERROR);
            row.createCell(2, CellType.ERROR); //esta es la forma que se debe utilizar (+actual) para establecer un tipo determinado en la celda

            row = hoja.createRow((short) 6);
            row.createCell(0).setCellValue("poener valor de fecha");
            row.createCell(1).setCellValue(new Date());

            row = hoja.createRow((short) 7);
            row.createCell(0).setCellValue("poner de tipo numérico");
            row.createCell(1, CellType.NUMERIC);

            row = hoja.createRow((short) 0);
            row.createCell(0).setCellValue("poner un string");
            row.createCell(1).setCellValue("esto es un string");

            //fechas con formato
            row = hoja.createRow((short) 8);
            CellStyle style = libro.createCellStyle();
            CreationHelper ch = libro.getCreationHelper();
            style.setDataFormat(ch.createDataFormat().getFormat("d/m/yy h:mm"));
            row.createCell(0).setCellValue("formato de fecha d/m/yy h:mm");
            Cell cell = row.createCell(1);
            cell.setCellValue(new Date());
            cell.setCellStyle(style);

            libro.write(out);
            out.close();

        } catch (FileNotFoundException ex) {
            Logger.getLogger(ApachePoi.class.getName()).log(Level.SEVERE, null, ex);
            System.out.println(ex.getMessage());
            JOptionPane.showMessageDialog(null, ex.getMessage());
        } catch (IOException ex) {
            Logger.getLogger(ApachePoi.class.getName()).log(Level.SEVERE, null, ex);
            System.out.println(ex);
        }
    }
    
    public static void main(String[] args) throws IOException {    
        // TODO code application logic here
        libro_blanco();
        abrir_libro();
        escribir_hoja_calc();
        System.out.println("");
        leer_hoja_calc();
        diferentes_tipos_celdas();
    }
    
}
