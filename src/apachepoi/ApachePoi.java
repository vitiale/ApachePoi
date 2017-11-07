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
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
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
        FileOutputStream file=new FileOutputStream(new File("blanco.xlsx"));
        XSSFWorkbook libro_blanco=new XSSFWorkbook();
        libro_blanco.write(file);
        System.out.println("Se ha creado el libro excel en blanco correctamente");
        file.close();
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
        map.put("1", new Object[]{"NOMBRE","EDAD"});
        map.put("2", new Object[]{"Pedro","29"});
        map.put("3", new Object[]{"Oscar","26"});
        map.put("4", new Object[]{"María","23"});
        Set<String> llaves=map.keySet();
        int row_id=0;
        for(String llave:llaves){
            //esto es lo que va a escribir
            System.out.println(llave+"  "+map.get(llave)[0]+"  "+map.get(llave)[1]);
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
    
    public static void main(String[] args) throws IOException {    
        // TODO code application logic here
        libro_blanco();
        abrir_libro();
        escribir_hoja_calc();
        System.out.println("");
        leer_hoja_calc();
    }
    
}
