package glgc.jjgys.system.utils;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class RowCopy {

    public static void copyRowMultiLine(Workbook wb, String pSourceSheetName,
                                String pTargetSheetName, int pStartRow, int pEndRow, int pPosition) {
        XSSFRow sourceRow = null;
        XSSFRow targetRow = null;
        XSSFCell sourceCell = null;
        XSSFCell targetCell = null;
        XSSFSheet sourceSheet = null;
        XSSFSheet targetSheet = null;
        CellRangeAddress region = null;
        int cType;
        int i;
        short j;
        int targetRowFrom;
        int targetRowTo;

        if ((pStartRow == -1) || (pEndRow == -1)) {
            return;
        }
        sourceSheet = (XSSFSheet) wb.getSheet(pSourceSheetName);
        targetSheet = (XSSFSheet) wb.getSheet(pTargetSheetName);

        Header header = targetSheet.getHeader();
        header.setLeft(sourceSheet.getHeader().getLeft());
        header.setCenter(sourceSheet.getHeader().getCenter());
        header.setRight(sourceSheet.getHeader().getRight());

        Footer footer = targetSheet.getFooter();
        footer.setLeft(sourceSheet.getFooter().getLeft());
        footer.setCenter(sourceSheet.getFooter().getCenter());
        footer.setRight(sourceSheet.getFooter().getRight());

        // 拷贝合并的单元格
        for (i = 0; i < sourceSheet.getNumMergedRegions(); i++) {
            region = sourceSheet.getMergedRegion(i);
            if ((region.getFirstRow() >= pStartRow)
                    && (region.getLastRow()) <= pEndRow) {
                targetRowFrom = region.getFirstRow() - pStartRow + pPosition;
                targetRowTo = region.getLastRow() - pStartRow + pPosition;
                region.setFirstRow(targetRowFrom);
                region.setLastRow(targetRowTo);
                targetSheet.addMergedRegion(region);
            }
        }
        // 设置列宽
        for (i = pStartRow; i <= pEndRow; i++) {
            sourceRow = sourceSheet.getRow(i);
            if (sourceRow != null) {
                for (j = sourceRow.getLastCellNum(); j > sourceRow
                        .getFirstCellNum(); j--) {
                    targetSheet
                            .setColumnWidth(j, sourceSheet.getColumnWidth(j));
                    targetSheet.setColumnHidden(j, false);
                }
                break;
            }
        }
        // 拷贝行并填充数据
        for (; i <= pEndRow; i++) {
            sourceRow = sourceSheet.getRow(i);
            if (sourceRow == null) {
                continue;
            }
            targetRow = targetSheet.createRow(i - pStartRow + pPosition);
            targetRow.setHeight(sourceRow.getHeight());
            for (j = sourceRow.getFirstCellNum(); j < sourceRow.getPhysicalNumberOfCells(); j++) {
                sourceCell = sourceRow.getCell(j);
                if (sourceCell == null) {
                    continue;
                }
                targetCell = targetRow.createCell(j);
                targetCell.setCellStyle(sourceCell.getCellStyle());
                cType = sourceCell.getCellType();
                targetCell.setCellType(cType);
                switch (cType) {
                    case XSSFCell.CELL_TYPE_BLANK:
                        targetCell.setCellValue("");
                        break;
                    case XSSFCell.CELL_TYPE_BOOLEAN:
                        targetCell.setCellValue(sourceCell.getBooleanCellValue());
                        break;
                    case XSSFCell.CELL_TYPE_ERROR:
                        targetCell
                                .setCellErrorValue(sourceCell.getErrorCellValue());
                        break;
                    case XSSFCell.CELL_TYPE_FORMULA:
                        // parseFormula这个函数的用途在后面说明
                        targetCell.setCellFormula(parseFormula(sourceCell
                                .getCellFormula()));
                        break;
                    case XSSFCell.CELL_TYPE_NUMERIC:
                        targetCell.setCellValue(sourceCell.getNumericCellValue());
                        break;
                    case XSSFCell.CELL_TYPE_STRING:
                        targetCell
                                .setCellValue(sourceCell.getRichStringCellValue());
                        break;
                }

            }

        }

    }


    /**
     *
     * @param wb
     * @param pSourceSheetName
     * @param pTargetSheetName
     * @param pStartRow
     * @param pEndRow
     * @param pPosition
     */
    public static void copyRows(XSSFWorkbook wb, String pSourceSheetName, String pTargetSheetName, int pStartRow, int pEndRow, int pPosition) {
        XSSFRow sourceRow = null;
        XSSFRow targetRow = null;
        XSSFCell sourceCell = null;
        XSSFCell targetCell = null;
        XSSFSheet sourceSheet = null;
        XSSFSheet targetSheet = null;
        CellRangeAddress region = null;
        int cType;
        int i;
        short j;
        int targetRowFrom;
        int targetRowTo;

        if ((pStartRow == -1) || (pEndRow == -1)) {
            return;
        }
        sourceSheet = wb.getSheet(pSourceSheetName);
        targetSheet = wb.getSheet(pTargetSheetName);

        Header header = targetSheet.getHeader();
        header.setLeft(sourceSheet.getHeader().getLeft());
        header.setCenter(sourceSheet.getHeader().getCenter());
        header.setRight(sourceSheet.getHeader().getRight());

        Footer footer = targetSheet.getFooter();
        footer.setLeft(sourceSheet.getFooter().getLeft());
        footer.setCenter(sourceSheet.getFooter().getCenter());
        footer.setRight(sourceSheet.getFooter().getRight());

        // 拷贝合并的单元格
        for (i = 0; i < sourceSheet.getNumMergedRegions(); i++) {
            region = sourceSheet.getMergedRegion(i);
            if ((region.getFirstRow() >= pStartRow) && (region.getLastRow()) <= pEndRow) {
                targetRowFrom = region.getFirstRow() - pStartRow + pPosition;
                targetRowTo = region.getLastRow() - pStartRow + pPosition;
                region.setFirstRow(targetRowFrom);
                region.setLastRow(targetRowTo);
                targetSheet.addMergedRegion(region);
            }
        }
        // 设置列宽
        for (i = pStartRow; i <= pEndRow; i++) {
            sourceRow = sourceSheet.getRow(i);
            if (sourceRow != null) {
                for (j = sourceRow.getLastCellNum(); j > sourceRow
                        .getFirstCellNum(); j--) {
                    targetSheet
                            .setColumnWidth(j, sourceSheet.getColumnWidth(j));
                    targetSheet.setColumnHidden(j, false);
                }
                break;
            }
        }
        // 拷贝行并填充数据
        for (; i <= pEndRow; i++) {
            sourceRow = sourceSheet.getRow(i);
            if (sourceRow == null) {
                continue;
            }
            targetRow = targetSheet.createRow(i - pStartRow + pPosition);
            targetRow.setHeight(sourceRow.getHeight());
            for (j = sourceRow.getFirstCellNum(); j < sourceRow.getPhysicalNumberOfCells(); j++) {
                sourceCell = sourceRow.getCell(j);
                if (sourceCell == null) {
                    continue;
                }
                targetCell = targetRow.createCell(j);
                targetCell.setCellStyle(sourceCell.getCellStyle());
                cType = sourceCell.getCellType();
                targetCell.setCellType(cType);
                switch (cType) {
                    case XSSFCell.CELL_TYPE_BLANK:
                        targetCell.setCellValue("");
                        break;
                    case XSSFCell.CELL_TYPE_BOOLEAN:
                        targetCell.setCellValue(sourceCell.getBooleanCellValue());
                        break;
                    case XSSFCell.CELL_TYPE_ERROR:
                        targetCell
                                .setCellErrorValue(sourceCell.getErrorCellValue());
                        break;
                    case XSSFCell.CELL_TYPE_FORMULA:
                        // parseFormula这个函数的用途在后面说明
                        targetCell.setCellFormula(parseFormula(sourceCell
                                .getCellFormula()));
                        break;
                    case XSSFCell.CELL_TYPE_NUMERIC:
                        targetCell.setCellValue(sourceCell.getNumericCellValue());
                        break;
                    case XSSFCell.CELL_TYPE_STRING:
                        targetCell
                                .setCellValue(sourceCell.getRichStringCellValue());
                        break;
                }

            }

        }

    }
    private static String parseFormula(String pPOIFormula) {
        final String cstReplaceString = "ATTR(semiVolatile)"; //$NON-NLS-1$
        StringBuffer result = null;
        int index;

        result = new StringBuffer();
        index = pPOIFormula.indexOf(cstReplaceString);
        if (index >= 0) {
            result.append(pPOIFormula.substring(0, index));
            result.append(pPOIFormula.substring(index
                    + cstReplaceString.length()));
        } else {
            result.append(pPOIFormula);
        }

        return result.toString();
    }

}
