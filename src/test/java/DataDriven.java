import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class DataDriven {

    static String filePath = "C:\\Users\\strdo\\Documents\\demoData.xlsx";
    static String sheetName = "test data";
    static String testCaseCellName = "Test Cases";
    static String testCaseName = "Purchase";

    public static void main(String[] args) throws IOException {
        ArrayList<String> testCaseData = getExcelData(testCaseName);
        for (int i = 0; i < testCaseData.size(); i++) {
            System.out.println(testCaseData.get(i));
        }
    }
    public static ArrayList<String> getExcelData(String testCaseName) throws IOException {
        ArrayList<String> excelData = new ArrayList<>();
        FileInputStream fis = new FileInputStream(filePath);
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fis);
        int numSheets = xssfWorkbook.getNumberOfSheets();
        for (int i = 0; i < numSheets; i++) {
            if (xssfWorkbook.getSheetName(i).equalsIgnoreCase(sheetName))
            {
                XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(i);

                //find 'Test Cases' column
                Iterator<Row> rows = xssfSheet.iterator();
                Row topRow = rows.next();
                Iterator<Cell> cells = topRow.cellIterator();

                int testCasesColumnIndex = 0;
                while (cells.hasNext()) {
                    Cell cell = cells.next();
                    if (cell.getStringCellValue().equalsIgnoreCase(testCaseCellName)) {
                        testCasesColumnIndex = cell.getColumnIndex();
                        break;
                    }
                }

                //in the 'Test Cases' column, find row with the required test case and get that row's data
                while (rows.hasNext()) {
                    Row row = rows.next();
                    if (row.getCell(testCasesColumnIndex).getStringCellValue().equalsIgnoreCase(testCaseName)) {
                        Iterator<Cell> cellsTestRow = row.cellIterator();
                        cellsTestRow.next();
                        while (cellsTestRow.hasNext()) {

                            Cell cell = cellsTestRow.next();
                            if (cell.getCellType() == CellType.STRING) {
                                excelData.add(cell.getStringCellValue());
                            }
                            else if (cell.getCellType() == CellType.NUMERIC) {
                                excelData.add(String.valueOf(cell.getNumericCellValue()));
                            }
                        }
                    }
                }
            }
        }
        return excelData;
    }
}
