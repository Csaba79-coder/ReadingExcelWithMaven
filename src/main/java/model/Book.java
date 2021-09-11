package model;

import lombok.Getter;
import lombok.Setter;

public class Book {

    // @Getter static
    @Getter
    @Setter
    public String author;

    @Getter
    @Setter
    private String title;

    @Getter
    @Setter
    private String series;

    @Setter
    private Double tome;

    public Integer getTome() {
        if (this.tome!= null) {
            return (int)Math.round(tome);
        }
        return 0;
    }

    @Getter
    @Setter
    private String publisher;

    @Setter
    private Double year;

    private Integer getYear() {
        if (this.year!= null) {
            return (int)Math.round(year);
        }
        return 0;
    }

    public Book() {
    }

    @Override
    public String toString() {
        return this.getAuthor() + ": " + this.getTitle() + " published " + this.getPublisher() + ": at " + this.getYear();
    }
}
