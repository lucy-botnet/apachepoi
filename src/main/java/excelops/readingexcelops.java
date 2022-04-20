package excelops;

import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import static com.sun.org.apache.bcel.internal.classfile.ElementValue.STRING;
import static java.sql.Types.BOOLEAN;
import static java.sql.Types.NUMERIC;

public class readingexcelops{
    public static void main(String[] args) throws IOException {
        String excelfilepath=".\\collection\\samplebook.xlsx";    //specifying the path of the excel sheet

        FileInputStream inputStream = new FileInputStream(excelfilepath);  //creating a stream to read the data

        //now need to get the workbook, sheet, rows and the cells from the excel
        XSSFWorkbook workbook = new  XSSFWorkbook(inputStream);

        //extracting the sheet from the workbook and referring the sheet with the object sheet
        XSSFSheet sheet =workbook.getSheet("Sheet 1");    //can get sheet through this or

         // XSSFSheet sheet = workbook.getSheetAt(0);
        // this is another method through which we can get the sheet and the index  will always start from 0

        //getting rows and cells USING FOR LOOP
        int rows =sheet.getLastRowNum(); //this will return the number of rows
        int cols =sheet.getRow(1).getLastCellNum(); //this will tell the number of columns or cells in one row

//        for (int r=0;r<rows;r++)  //for getting the row
//        {
//           XSSFRow row = sheet.getRow(r);
//            for (int c=0;c<cols;c++)  ///for getting to the cell
//            {
//                XSSFCell cell = row.getCell(c);
//                switch(cell.getCellType())       //get the type of value the cell holds using switch case
//                {case STRING:
//                    System.out.println(cell.getStringCellValue()); break;
//                case NUMERIC:
//                    System.out.println(cell.getNumericCellValue()); break;
//                case BOOLEAN:
//                    System.out.println(cell.getBooleanCellValue()); break;
//                }
//            }
//            System.out.println();
//        }
      //////// using ITERATOR METHOD
        Iterator iterator =sheet.iterator();
        while (iterator.hasNext())
        {
            XSSFRow row = (XSSFRow) iterator.next();

            Iterator celliterator =row.iterator();
            while(celliterator.hasNext())
            {
                XSSFCell cell= (XSSFCell) celliterator.next();
                switch(cell.getCellTypeEnum())       //get the type of value the cell holds using switch case
                {case STRING:
                    System.out.println(cell.getStringCellValue()); break;
                case NUMERIC:
                    System.out.println(cell.getNumericCellValue()); break;
                case BOOLEAN:
                    System.out.println(cell.getBooleanCellValue()); break;
                }
            }
            System.out.println();
        }
    }
}
