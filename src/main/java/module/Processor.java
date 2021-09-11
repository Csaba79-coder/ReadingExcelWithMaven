package module;

import model.Book;
import utils.ExcelReader;

public class Processor {

    public void run() {
        // Books.getAuthor(); // when static's written behind getter
        // new Books().getAuthor(); // when static's not written behind getter

        new ExcelReader().readBookFromExcel();
    }
}
