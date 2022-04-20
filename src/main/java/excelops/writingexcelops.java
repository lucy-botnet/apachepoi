package excelops;


import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

///workbook->sheet->rows->cells
public class writingexcelops  {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();  ///creating a workbook
        XSSFSheet sheet = workbook.createSheet("EMP INFO");   ///creating a sheet in the workbook

        //creating data in some 2D array
        Object empdata[][] = {{"Emp ID", "Name", "Job"},
                {101, "utkarsh", "HR"},
                {102, "bot", "Engineer"},
                {103, "saruabh", "developer"},
                {104, "sunder", "Accountant"}
        };
//        //using FOR LOOP
//        int rows = empdata.length;   ///getting the number of rows
//        int cols = empdata[0].length;    //getting the number of cells in a row
//
////        System.out.println(rows);     ///printing the number of rows
////        System.out.println(cols);      ////printing the number of cells
//
//        for (int r = 0; r < rows; r++) ////for rows
//        {
//            XSSFRow row = sheet.createRow(r);
//            for (int c = 0; c < cols; c++)   ///responsible for writing in the cells
//            {
//                XSSFCell col = row.createCell(c);
//                Object value = empdata[r][c];
//
//                if (value instanceof String)                           ////checking the type of data and then
//                    col.setCellValue((String)value);                   ////writing the data accordingly in the cell
//                if (value instanceof Integer)
//                    col.setCellValue((Integer)value);
//                if (value instanceof Boolean)
//                    col.setCellValue((Boolean) value);
//            }
//        }

        ///using FOR each loop
        int rowcount=0;
        for (Object emp[]:empdata) {
            XSSFRow row = sheet.createRow(rowcount++);
            int oclcount=0;
            for (Object value: emp){
                XSSFCell col=row.createCell(oclcount++);
                if (value instanceof String)                           ////checking the type of data and then
                    col.setCellValue((String)value);                   ////writing the data accordingly in the cell
                if (value instanceof Integer)
                    col.setCellValue((Integer)value);
                if (value instanceof Boolean)
                    col.setCellValue((Boolean) value);
            }
        }
        String filepath =".\\collection\\employee.xlsx";   //creating the file path
        FileOutputStream outStream = new FileOutputStream(filepath);    ///fileoutputstream cuz we want to write data
        workbook.write(outStream);   /////writing the data in the workbook

        outStream.close();

        System.out.println(" Employee.xls file written successfully");
    }
}
