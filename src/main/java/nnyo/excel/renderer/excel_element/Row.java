package nnyo.excel.renderer.excel_element;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;
import java.util.StringJoiner;

public class Row implements Serializable {

    private List<Cell> cells = new ArrayList<>();

    public List<Cell> getCells() {
        return cells;
    }

    public Row setCells(List<Cell> cells) {
        this.cells = cells;
        return this;
    }

    @Override
    public String toString() {
        return new StringJoiner(", ", Row.class.getSimpleName() + "[", "]")
                .add("cells=" + cells)
                .toString();
    }
}
