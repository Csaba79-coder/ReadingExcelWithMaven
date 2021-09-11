package utils;

import model.Book;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;

public class ExcelReader {

    private final String EXCEL_FILEPATH = "src/main/resources/myBooks.xlsx";

    public void readBookFromExcel() {
        try {
            List<Book> books = Collections.unmodifiableList(readDataFromExcelFile(EXCEL_FILEPATH));
            printExcel(books);
        } catch (IOException e) {
            System.out.println("Reading excel file failed!");
            System.out.println(e);
        }

    }

    private void printExcel(List<Book> books) {
        for (Book s : books) {
            System.out.println(s);
        }
    }

    private List<Book> readDataFromExcelFile(String excelFilePath) throws IOException {

        // Creating an List object of Employee type
        // Note: User defined type

        List<Book> books = new ArrayList<>();

        FileInputStream inputStream = new FileInputStream(excelFilePath);
        // As used 'xlsx' file is used so XSSFWorkbook will be
        // used
        Workbook workbook = new XSSFWorkbook(inputStream);

        // Read the first sheet and if the contents are in
        // different sheets specifying the correct index
        Sheet firstSheet = workbook.getSheetAt(0);

        // Iterators to traverse over

        // Condition check using hasNext() method which holds
        // true till there is single element remaining in List

        for (Row nextRow : firstSheet) {
            // Get a row in sheet
            // This is for a Row's cells
            Iterator<Cell> cellIterator
                    = nextRow.cellIterator();
            // We are taking Employee as reference.
            Book book = new Book();
            // Iterate over the cells
            while (cellIterator.hasNext()) {
                Cell nextCell = cellIterator.next();

                // Switch case variable to
                // get the columnIndex
                int columnIndex = nextCell.getColumnIndex();

                // Depends upon the cell contents we need to
                // typecast

                // Switch-case
                switch (columnIndex) {

                    // Case 1
                    case 0 ->
                            // First column is alpha and hence
                            // it is typecasted to String
                            book.setAuthor(
                                    (String) getCellValue(nextCell));

                    // Break keyword to directly terminate
                    // if this case is hit
                    // Case 2
                    case 1 ->
                            // Second  column is alpha and hence
                            // it is typecasted to String
                            book.setTitle(
                                    (String) getCellValue(nextCell));

                    // Break keyword to directly terminate
                    // if this case is hit
                    // Case 3
                    case 2 ->
                            // Third  column is double value and
                            // hence it is typecasted to Double
                            book.setSeries(
                                    (String) getCellValue(nextCell));


                    // Note: If additional cells are present
                    // then
                    // they should be specified further down,
                    // and POJO class should accommodate those
                    // cell values

                    case 3 -> {

                        try {
                            // System.out.println(getCellValue(nextCell));
                            book.setTome(
                                    Double.valueOf(getCellValue(nextCell)));
                        } catch (Exception  e) {
                            System.out.println("Error in the excel field called: tome");
                            System.out.println(e.getMessage());
                        }
                    }


                    case 4 ->

                            book.setPublisher(
                                    (String) getCellValue(nextCell));

                    case 5 -> {

                        try {
                            // System.out.println(getCellValue(nextCell));
                            book.setYear(
                                    Double.valueOf(getCellValue(nextCell)));
                        } catch (Exception  e) {
                            System.out.println("Error in the excel field called: year");
                            System.out.println(e.getMessage());
                        }
                    }


                    default -> throw new IllegalStateException("Unexpected value: " + columnIndex);
                }
            }
            // Adding up to the list
            books.add(book);
        }

        // Closing the workbook and inputstream
        // as it free up the space in memory
        workbook.close();
        inputStream.close();

        // Return all the employees present in List
        // object of Employee type
        return books;
    }


    // Java Program to get the cell value
    // of the corresponding cells

    // Method
    // To get the cell value

    @Deprecated
    // deprecated method!!! soon this one will not operate only if I use the same value

    /*
    in pom.xml: current version 5.0.0 but it is not including enums anymore!
    <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-ooxml</artifactId>
            <version>3.17</version>
        </dependency>
    */
    private String getCellValue(Cell cell) {

        // Now either do-while or switch can be used
        // to display menu/user's choice
        // Switch case is used here for illustration
        // Switch case to get the users choice
        return switch (cell.getCellType()) {



            // Case 1
            // If cell contents are string
            case Cell.CELL_TYPE_STRING -> cell.getStringCellValue();

            // Case 2
            // If cell contents are Boolean
            case Cell.CELL_TYPE_BOOLEAN -> String.valueOf(cell.getBooleanCellValue());

            // Case 3
            // If cell contents are Numeric which includes
            // int, float , double etc
            case Cell.CELL_TYPE_NUMERIC -> String.valueOf(cell.getNumericCellValue());
            default ->

                    // Case 4
                    // Default case
                    // If cell contents are neither
                    // string nor Boolean nor Numeric,
                    // simply nothing is returned
                    null;
        };
    }
}
