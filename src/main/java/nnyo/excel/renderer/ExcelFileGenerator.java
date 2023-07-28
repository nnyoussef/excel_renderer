package nnyo.excel.renderer;


import nnyo.excel.renderer.dto.CoordinateDto;
import nnyo.excel.renderer.excel_element.Table;
import nnyo.excel.renderer.excel_element_handlers.TableExcelElementHandler;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;


public class ExcelFileGenerator {

    private static final Map<Class<?>, ExcelElementHandler> STRATEGIES_MAP = new HashMap<>();

    static {
        STRATEGIES_MAP.put(Table.class, new TableExcelElementHandler());
    }

    public static XSSFWorkbook generate(Map<String, List<Object>> rendereableObjectsPerSheet,
                                        String css,
                                        OutputStream outputStream) {
        ConcurrentHashMap<String, XSSFSheet> xssfSheetConcurrentHashMap = new ConcurrentHashMap<>();

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            final StyleContext styleContext = StyleContext.init(css, workbook);

            rendereableObjectsPerSheet.entrySet().stream()
                    .peek(stringListEntry -> xssfSheetConcurrentHashMap.put(stringListEntry.getKey(), workbook.createSheet(stringListEntry.getKey())))
                    .forEach(stringListEntry -> {
                        XSSFSheet xssfSheet = xssfSheetConcurrentHashMap.get(stringListEntry.getKey());
                        CoordinateDto coordinateDto = new CoordinateDto();
                        stringListEntry.getValue()
                                .forEach(o -> {
                                    if (o instanceof Table) {
                                        STRATEGIES_MAP.get(Table.class).handle(coordinateDto, o, xssfSheet, styleContext);
                                    }

                                });
                    });

            workbook.write(outputStream);
            return workbook;
        } catch (Exception e) {
            return null;
        }
    }
}
