package nnyo.excel.renderer.excel_element_handlers;


import nnyo.excel.renderer.ExcelElementHandler;
import nnyo.excel.renderer.StyleContext;
import nnyo.excel.renderer.dto.CoordinateDto;
import nnyo.excel.renderer.excel_element.Row;
import nnyo.excel.renderer.excel_element.Table;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.Collection;

import static java.util.Optional.ofNullable;
import static nnyo.excel.renderer.utils.CellDataUtils.setData;
import static nnyo.excel.renderer.utils.SpanUtils.createSpan;
import static nnyo.excel.renderer.utils.SpanUtils.findClosestUnmergedCellOfRow;
import static org.apache.poi.ss.util.RegionUtil.setBorderBottom;
import static org.apache.poi.ss.util.RegionUtil.setBorderRight;

public class TableExcelElementHandler implements ExcelElementHandler {

    @Override
    public void handle(CoordinateDto coordinateDto,
                       Object elementToHandle,
                       XSSFSheet worksheet,
                       StyleContext styleContext) {

        Table table = ((Table) elementToHandle);

        renderTableType1(worksheet, table.getHeader(), coordinateDto, "thead", styleContext);
        renderTableType1(worksheet, table.getBody(), coordinateDto, "tbody", styleContext);
        renderTableType1(worksheet, table.getFooter(), coordinateDto, "tfoot", styleContext);

        coordinateDto.setCellPosition(1);
        coordinateDto.incrementPosition(1, 0);
    }

    private void renderTableType1(XSSFSheet worksheet,
                                  Collection<Row> header,
                                  CoordinateDto coordinateDto,
                                  String tableSection,
                                  StyleContext styleContext) {
        header.forEach(row -> {
            row.getCells()
                    .forEach(cell -> {
                        int rowSpan = cell.getRowSpan();
                        int colSpan = cell.getColSpan();
                        String css = ofNullable(cell.getCssClass())
                                .map(c -> StringUtils.isEmpty(c) ? null : c)
                                .map(c -> String.format("%s .%s", tableSection, c))
                                .orElse(String.format("%s .default", tableSection));

                        int closestUnmergedCellIndex = findClosestUnmergedCellOfRow(worksheet, coordinateDto);
                        coordinateDto.setCellPosition(closestUnmergedCellIndex);

                        XSSFRow xssfRow = getRow(worksheet, coordinateDto);
                        XSSFCell xssfCell = xssfRow.createCell(coordinateDto.getCellPosition());
                        XSSFCellStyle xssfCellStyle = styleContext.createStyle(css);

                        CellRangeAddress cellRangeAddress = createSpan(worksheet, rowSpan, colSpan, coordinateDto);

                        coordinateDto.incrementPosition(0, colSpan);

                        setData(cell.getData(), xssfCell);
                        resolveBorderForMergedCells(cellRangeAddress, xssfCellStyle, worksheet);
                        xssfCell.setCellStyle(xssfCellStyle);

                    });
            coordinateDto.setCellPosition(1); // each row iteration the cell cursor should be back to 1
            coordinateDto.incrementPosition(1, 0); // tracing the position of the current row
        });
    }

    private XSSFRow getRow(XSSFSheet sheet,
                           CoordinateDto coordinateDto) {
        XSSFRow xssfRow = sheet.getRow(coordinateDto.getRowPosition());
        if (xssfRow == null)
            return sheet.createRow(coordinateDto.getRowPosition());
        else return sheet.getRow(coordinateDto.getRowPosition());
    }

    private void resolveBorderForMergedCells(CellRangeAddress cellRangeAddress,
                                             XSSFCellStyle xssfCellStyle,
                                             XSSFSheet sheet) {
        setBorderBottom(xssfCellStyle.getBorderBottom(), cellRangeAddress, sheet);
        setBorderRight(xssfCellStyle.getBorderRight(), cellRangeAddress, sheet);
    }

    private int convertByteArrayToInt(byte[] bytes) {
        int value = 0;
        for (byte b : bytes) {
            value = (value << 8) + (b & 0xFF);
        }
        return value;
    }

}
