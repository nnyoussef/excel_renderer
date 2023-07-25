package nnyo.excel.renderer;

import nnyo.excel.renderer.excel_element.Cell;
import nnyo.excel.renderer.excel_element.Row;
import nnyo.excel.renderer.excel_element.Table;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Random;
import java.util.stream.IntStream;

import static java.nio.file.Files.newOutputStream;
import static java.nio.file.Path.of;
import static java.util.Arrays.asList;

public class ExportTablesSheetTest {

    @Test
    public void testSimpleTableFormat1() throws IOException {
        String css = "/* Cell Style for header*/\n" +
                "thead .bleuFonce {\n" +
                "  background: #000080;\n" +
                "  border: 1px solid #7F7F7F !important;\n" +
                "  color: #FFFFFF;\n" +
                "  padding-right: 12px;\n" +
                "  padding-left: 12px;\n" +
                "  text-align: center;\n" +
                "}\n" +
                "\n" +
                "thead .bleuClaire {\n" +
                "  background: #6EAAE6;\n" +
                "  border: 1px solid #7F7F7F;\n" +
                "  color: #FFFFFF;\n" +
                "  text-align: center;\n" +
                "}\n" +
                "\n" +
                "/* Cell style By Default*/\n" +
                "thead .default {\n" +
                "  border: none;\n" +
                "}\n" +
                "\n" +
                "tbody .default {\n" +
                "  border: 1px solid #7F7F7F;\n" +
                "  color: #000000;\n" +
                "}\n" +
                "\n" +
                "tfoot .default {\n" +
                "  background: #A9A9A9;\n" +
                "  color: #000000;\n" +
                "  font-weight: bold;\n" +
                "  border: 1px solid #7F7F7F;\n" +
                "}\n";
        String exportPath = "/home/nyoussef/Desktop/Workspace/excel-renderer/excel.xlsx";

        LinkedHashMap<String, List<Object>> data = new LinkedHashMap<>();

        data.put("Sheet 1 - Test", getData());
        data.put("Sheet 2 - Test", getData());
        data.put("Sheet 3 - Test", getData());
        data.put("Sheet 4 - Test", getData());


        ExcelFileGenerator.generate(data, css, newOutputStream(of(exportPath)));
    }

    private static List<Object> getData() {

        Table table = new Table();

        Cell year = new Cell().setData("Year").setColSpan(4).setCssClass("bleuFonce");
        Cell empty = new Cell().setData("I am an Empty Cell").setRowSpan(2).setColSpan(2).setCssClass("bleuFonce");

        Cell y1 = new Cell().setData("2020").setCssClass("bleuClaire");
        Cell y2 = new Cell().setData("2021").setCssClass("bleuClaire");
        Cell y3 = new Cell().setData("2022").setCssClass("bleuClaire");
        Cell y4 = new Cell().setData("2023").setCssClass("bleuClaire");

        List<Row> headerRow = new LinkedList<>();
        headerRow.add(new Row().setCells(asList(year, empty)));
        headerRow.add(new Row().setCells(asList(y1, y2, y3, y4)));

        List<Row> body = IntStream.range(0, 50)
                .mapToObj(i -> {
                    Random random = new Random();
                    Cell data1 = new Cell().setData(random.nextLong(10000));
                    Cell data2 = new Cell().setData(random.nextLong(10000));
                    Cell data3 = new Cell().setData(random.nextLong(10000));
                    Cell data4 = new Cell().setData(random.nextLong(10000));
                    Cell data5 = new Cell().setData(random.nextLong(10000)).setColSpan(2);
                    return new Row().setCells(asList(data1, data2, data3, data4, data5));
                }).toList();

        List<Row> footerRow = new LinkedList<>();
        Cell t1 = new Cell().setData("2020").setCssClass("");
        Cell t2 = new Cell().setData("2021").setCssClass("");
        Cell t3 = new Cell().setData("2022").setCssClass("");
        Cell t4 = new Cell().setData("2023").setCssClass("");
        Cell t5 = new Cell().setData("Total").setCssClass("").setColSpan(2);
        footerRow.add(new Row().setCells(asList(t1, t2, t3, t4, t5)));

        table.setHeader(headerRow);
        table.setBody(body);
        table.setFooter(footerRow);

        return asList(table, table, table, table, table);
    }
}
