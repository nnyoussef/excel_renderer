package nnyo.excel.renderer.utils;

import nnyo.excel.renderer.dto.CoordinateDto;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressBase;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.concurrent.atomic.AtomicInteger;

public class SpanUtils {

    private SpanUtils() {
    }

    public static int findClosestUnmergedCellOfRow(XSSFSheet xssfSheet,
                                                   CoordinateDto coordinateDto) {
        AtomicInteger currentCellIndex = new AtomicInteger(coordinateDto.getCellPosition());
        AtomicInteger currentRowIndex = new AtomicInteger(coordinateDto.getRowPosition());

        xssfSheet.getMergedRegions()
                .stream()
                .filter(cellAddresses -> {
                    int firstRow = cellAddresses.getFirstRow();
                    int lastRow = cellAddresses.getLastRow();

                    int firstColumn = cellAddresses.getFirstColumn();
                    int lastColumn = cellAddresses.getLastColumn();

                    int currentRow = currentRowIndex.get();
                    int currentColumn = currentCellIndex.get();

                    if (currentRow >= firstRow && currentRow <= lastRow)
                        return currentColumn >= firstColumn && currentColumn <= lastColumn;

                    return false;
                }).map(CellRangeAddressBase::getLastColumn)
                .forEach(integer -> currentCellIndex.set(integer + 1));
        return currentCellIndex.get();


    }

    public static CellRangeAddress createSpan(XSSFSheet sheet,
                                              int rowSpan,
                                              int colSpan,
                                              CoordinateDto coordinateDto) {
        int rowIndex = coordinateDto.getRowPosition();
        int colIndex = coordinateDto.getCellPosition();
        CellRangeAddress cellAddresses = new CellRangeAddress(rowIndex, rowIndex + rowSpan - 1, colIndex, colIndex + colSpan - 1);

        if (rowSpan * colSpan > 1)
            sheet.addMergedRegion(cellAddresses);
        return cellAddresses;
    }
}
