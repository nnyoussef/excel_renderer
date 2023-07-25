package nnyo.excel.renderer.excel_element;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.StringJoiner;

public class Table {

    List<Row> header = Collections.emptyList();
    List<Row> body = Collections.emptyList();
    List<Row> footer = Collections.emptyList();
    private final List<Integer> nbOfIndexCols = new ArrayList<>(8);

    public List<Integer> getNbOfIndexCols() {
        return nbOfIndexCols;
    }

    public List<Row> getHeader() {
        return header;
    }

    public void setHeader(List<Row> header) {
        this.header = header;
    }

    public List<Row> getBody() {
        return body;
    }

    public void setBody(List<Row> body) {
        this.body = body;
    }

    public List<Row> getFooter() {
        return footer;
    }

    public void setFooter(List<Row> footer) {
        this.footer = footer;
    }

    @Override
    public String toString() {
        return new StringJoiner(", ", Table.class.getSimpleName() + "[", "]")
                .add("header=" + header)
                .add("body=" + body)
                .add("footer=" + footer)
                .toString();
    }

}
