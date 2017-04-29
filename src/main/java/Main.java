import org.apache.poi.hssf.usermodel.HSSFCell;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.Map;

public class Main {

    public static Map<String, String> fromFile = new LinkedHashMap<String, String>();

    public static void main (String [] args) throws FileNotFoundException, IOException{
        readFromExcel("/home/cosysoft/parser/src/main/resources/Haulmont .xlsx");
        writeFromExcel("/home/cosysoft/parser/src/main/resources/Province.xlsx");
    }

    public static void readFromExcel(String file) throws IOException {



        String nameRu = null;
        String nameEN = null;

        XSSFWorkbook myExcelBook = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet myExcelSheet = myExcelBook.getSheet("province");



        for (int i=1;i<2013;i++ ) {

            XSSFRow row = myExcelSheet.getRow(i);
            if (row.getCell(0).getCellType() == HSSFCell.CELL_TYPE_STRING) {
                nameRu = row.getCell(0).getStringCellValue();
                //System.out.println(nameRu);
            }


            if (row.getCell(0).getCellType() == HSSFCell.CELL_TYPE_STRING) {
                nameEN = row.getCell(1).getStringCellValue();
                //System.out.println(nameEN);
            }

            fromFile.put(nameRu, nameEN);
        }



        myExcelBook.close();

    }

    @SuppressWarnings("deprecation")
    private static void writeFromExcel(String file) throws FileNotFoundException, IOException {

        String name = null;

        XSSFWorkbook myExcelBook = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet myExcelSheet = myExcelBook.getSheet("1");

        System.out.println("setup1" );

        for (int i = 3;i < 3105;i++) {

            XSSFRow row = myExcelSheet.getRow(3);

            System.out.println("setup2");
            if (row.getCell(1).getCellType() == HSSFCell.CELL_TYPE_STRING) {
                name = row.getCell(1).getStringCellValue();
            }

            if (fromFile.containsValue(name)) {
                Cell nameEn = row.createCell(2);
                nameEn.setCellValue(fromFile.get(name));
                System.out.println("setup2");
            }

        }

        myExcelBook.write(new FileOutputStream(file.concat("new")));
        System.out.println("setup3");
        myExcelBook.close();
    }
}
