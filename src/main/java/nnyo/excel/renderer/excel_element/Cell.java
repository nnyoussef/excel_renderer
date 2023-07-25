package nnyo.excel.renderer.excel_element;

import java.io.Serializable;
import java.util.LinkedHashMap;
import java.util.StringJoiner;
import java.util.function.Function;

public class Cell implements Serializable {

    private String cssClass;
    private Object data;
    private int colSpan = 1;
    private int rowSpan = 1;
    private boolean isDataKey;

    private boolean editable;

    private boolean selectable;

    private LinkedHashMap<String, String> selectionValueLabelMap;

    private String displayFunction = "";

    public String getDisplayFunction() {
        return displayFunction;
    }

    public Cell setDisplayFunction(String displayFunction) {
        this.displayFunction = displayFunction;
        return this;
    }

    public boolean isDataKey() {
        return isDataKey;
    }

    public boolean isEditable() {
        return editable;
    }

    public void setEditable(boolean editable) {
        this.editable = editable;
    }

    public boolean isSelectable() {
        return selectable;
    }

    public void setSelectable(boolean selectable) {
        this.selectable = selectable;
    }

    public LinkedHashMap<String, String> getSelectionValueLabelMap() {
        if (selectionValueLabelMap == null)
            selectionValueLabelMap = new LinkedHashMap<>();
        return selectionValueLabelMap;
    }

    public void setSelectionValueLabelMap(LinkedHashMap<String, String> selectionValueLabelMap) {
        this.selectionValueLabelMap = selectionValueLabelMap;
    }

    public Cell setDataKey(boolean dataKey) {
        isDataKey = dataKey;
        return this;
    }

    public String getCssClass() {
        return cssClass;
    }

    public int getColSpan() {
        return colSpan;
    }

    public Cell setColSpan(int colSpan) {
        this.colSpan = colSpan;
        return this;
    }

    public int getRowSpan() {
        return rowSpan;
    }

    public Cell setRowSpan(int rowSpan) {
        this.rowSpan = rowSpan;
        return this;
    }

    public Cell setCssClass(String cssClass) {
        this.cssClass = cssClass;
        return this;
    }

    public Object getData() {
        return data;
    }

    public Cell setData(Object data) {
        this.data = data;
        return this;
    }

    public Cell setData(Object data, Function<Object, Object> transformer) {
        if (transformer != null)
            this.data = transformer.apply(data);
        else
            this.data = data;
        return this;
    }

    @Override
    public String toString() {
        return new StringJoiner(", ", Cell.class.getSimpleName() + "[", "]")
                .add("cssClass='" + cssClass + "'")
                .add("data=" + data)
                .add("colSpan=" + colSpan)
                .add("rowSpan=" + rowSpan)
                .add("isDataKey=" + isDataKey)
                .add("editable=" + editable)
                .add("selectable=" + selectable)
                .add("selectionValueLabelMap=" + selectionValueLabelMap)
                .toString();
    }
}
