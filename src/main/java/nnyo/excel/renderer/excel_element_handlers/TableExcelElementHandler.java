package nnyo.excel.renderer.excel_element_handlers;


import nnyo.excel.renderer.CellStyleProcessor;
import nnyo.excel.renderer.ExcelElementHandler;
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
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;

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
                       CellStyleProcessor cellStyleProcessor) {

        Table table = ((Table) elementToHandle);

        renderTableType1(worksheet, table.getHeader(), coordinateDto, "thead", cellStyleProcessor);
        renderTableType1(worksheet, table.getBody(), coordinateDto, "tbody", cellStyleProcessor);
        renderTableType1(worksheet, table.getFooter(), coordinateDto, "tfoot", cellStyleProcessor);

        coordinateDto.setCellPosition(1);
        coordinateDto.incrementPosition(1, 0);
    }

    private void renderTableType1(XSSFSheet worksheet,
                                  Collection<Row> header,
                                  CoordinateDto coordinateDto,
                                  String tableSection,
                                  CellStyleProcessor cellStyleProcessor) {

        final AtomicBoolean isThereAnyMergeInProgress = new AtomicBoolean(false);
        final AtomicInteger deepestRowSpanPosition = new AtomicInteger(coordinateDto.getRowPosition());

        header.forEach(row -> {
            row.getCells()
                    .forEach(cell -> {
                        int rowSpan = cell.getRowSpan();
                        int colSpan = cell.getColSpan();
                        String css = ofNullable(cell.getCssClass())
                                .map(c -> StringUtils.isEmpty(c) ? null : c)
                                .map(c -> String.format("%s .%s", tableSection, c))
                                .orElse(String.format("%s .default", tableSection));

                        if (isThereAnyMergeInProgress.get())
                            coordinateDto.setCellPosition(findClosestUnmergedCellOfRow(worksheet, coordinateDto));

                        XSSFRow xssfRow = getRow(worksheet, coordinateDto);
                        XSSFCell xssfCell = xssfRow.createCell(coordinateDto.getCellPosition());
                        XSSFCellStyle xssfCellStyle = cellStyleProcessor.createStyle(css);

                        CellRangeAddress cellAddressesAfterMerging = createSpan(worksheet, rowSpan, colSpan, coordinateDto);

                        coordinateDto.incrementPosition(0, colSpan);

                        setData(cell.getData(), xssfCell);
                        resolveBorderForMergedCells(cellAddressesAfterMerging, xssfCellStyle, worksheet);
                        xssfCell.setCellStyle(xssfCellStyle);

                        if (deepestRowSpanPosition.get() < cellAddressesAfterMerging.getLastRow()) {
                            deepestRowSpanPosition.set(cellAddressesAfterMerging.getLastRow());
                        }


                    });
            isThereAnyMergeInProgress.set(deepestRowSpanPosition.get() > coordinateDto.getRowPosition());
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
        if (cellRangeAddress.getNumberOfCells() == 1)
            return;

        setBorderBottom(xssfCellStyle.getBorderBottom(), cellRangeAddress, sheet);
        setBorderRight(xssfCellStyle.getBorderRight(), cellRangeAddress, sheet);

        XSSFRow lastRow = sheet.getRow(cellRangeAddress.getLastRow());
        for (int i = cellRangeAddress.getFirstColumn(); i < cellRangeAddress.getLastColumn() + 1; i++) {
            XSSFCell xssfCell = lastRow.getCell(i);
            xssfCell.setCellStyle(xssfCellStyle);
        }

        for (int i = cellRangeAddress.getFirstRow(); i < cellRangeAddress.getLastRow() + 1; i++) {
            XSSFCell xssfCell = sheet.getRow(i).getCell(cellRangeAddress.getLastColumn());
            xssfCell.setCellStyle(xssfCellStyle);
        }

    }
}
