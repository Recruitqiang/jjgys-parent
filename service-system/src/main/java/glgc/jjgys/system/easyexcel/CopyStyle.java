package glgc.jjgys.system.easyexcel;

import cn.hutool.core.collection.CollUtil;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.main.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTConnector;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTShape;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import static org.apache.poi.ss.usermodel.CellType.NUMERIC;

public class CopyStyle {

    /**
     * 行复制(多行)
     *
     * @param workbook     文档对象
     * @param sheet        sheet对象
     * @param fromRowIndex
     * @param toRowIndex
     * @param copyRowNum   复制行数
     */
    public static void copyRow(Workbook workbook, Sheet sheet, int fromRowIndex, int toRowIndex, int copyRowNum, boolean insertFlag
            , Integer colNum) {
        for (int i = 0; i < copyRowNum; i++) {
            Row fromRow = sheet.getRow(fromRowIndex + i);
            Row toRow = sheet.getRow(toRowIndex + i);
            if (insertFlag) {
                //复制行超出原有sheet页的最大行数时，不需要移动直接插入
                if (toRowIndex + i <= sheet.getLastRowNum()) {
                    //先移动要插入的行号所在行及之后的行
                    sheet.shiftRows(toRowIndex + i, sheet.getLastRowNum(), 1, true, false);
                }
                //然后再插入行
                toRow = sheet.createRow(toRowIndex + i);
                //设置行高
                toRow.setHeight(fromRow.getHeight());
            }
            for (int colIndex = 0; colIndex < (colNum != null ? colNum : fromRow.getLastCellNum()); colIndex++) {
                Cell tmpCell = fromRow.getCell(colIndex);
                if (tmpCell == null) {
                    tmpCell = fromRow.createCell(colIndex);
                    if (colIndex != 0) {
                        copyCell(workbook, fromRow.createCell(colIndex - 1), tmpCell);
                    }
                }
                Cell newCell = toRow.createCell(colIndex);
                copyCell(workbook, tmpCell, newCell);
            }
        }
        //获取合并单元格
        List<CellRangeAddress> cellRangeAddressList = sheet.getMergedRegions();
        Map<Integer, List<CellRangeAddress>> rowCellRangeAddressMap = CollUtil.isNotEmpty(cellRangeAddressList) ? cellRangeAddressList.stream().collect(Collectors.groupingBy(x ->
                x.getFirstRow())) : new HashMap<>();
        //获取形状（线条）
        XSSFDrawing drawing = (XSSFDrawing) sheet.getDrawingPatriarch();
        List<XSSFShape> shapeList = new ArrayList<>();
        Map<Integer, List<XSSFShape>> rowShapeMap = new HashMap<>();
        if (drawing != null) {
            shapeList = drawing.getShapes();
            rowShapeMap = shapeList.stream().filter(x -> x.getAnchor() != null)
                    .collect(Collectors.groupingBy(x -> ((XSSFClientAnchor) x.getAnchor()).getRow1()));
        }
        List<XSSFShape> insertShapeList = new ArrayList<>();
        for (int i = 0; i < copyRowNum; i++) {
            Row toRow = sheet.getRow(toRowIndex + i);
            //复制合并单元格
            List<CellRangeAddress> rowCellRangeAddressList = rowCellRangeAddressMap.get(fromRowIndex + i);
            if (CollUtil.isNotEmpty(rowCellRangeAddressList)) {
                for (CellRangeAddress cellRangeAddress : rowCellRangeAddressList) {
                    CellRangeAddress newCellRangeAddress = new CellRangeAddress(toRow.getRowNum(), (toRow.getRowNum() +
                            (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())), cellRangeAddress
                            .getFirstColumn(), cellRangeAddress.getLastColumn());
                    sheet.addMergedRegionUnsafe(newCellRangeAddress);
                }
            }
            //复制形状（线条）
            List<XSSFShape> rowShapeList = rowShapeMap.get(fromRowIndex + i);
            if (CollUtil.isNotEmpty(rowShapeList)) {
                for (XSSFShape shape : rowShapeList) {
                    //复制描点
                    XSSFClientAnchor fromAnchor = (XSSFClientAnchor) shape.getAnchor();
                    XSSFClientAnchor toAnchor = new XSSFClientAnchor();
                    toAnchor.setDx1(fromAnchor.getDx1());
                    toAnchor.setDx2(fromAnchor.getDx2());
                    toAnchor.setDy1(fromAnchor.getDy1());
                    toAnchor.setDy2(fromAnchor.getDy2());
                    toAnchor.setRow1(toRow.getRowNum());
                    toAnchor.setRow2(toRow.getRowNum() + fromAnchor.getRow2() - fromAnchor.getRow1());
                    toAnchor.setCol1(fromAnchor.getCol1());
                    toAnchor.setCol2(fromAnchor.getCol2());
                    //复制形状
                    if (shape instanceof XSSFConnector) {
                        copyXSSFConnector((XSSFConnector) shape, drawing, toAnchor);
                    } else if (shape instanceof XSSFSimpleShape) {
                        copyXSSFSimpleShape((XSSFSimpleShape) shape, drawing, toAnchor);
                    }
                }
            }
        }
    }

    /**
     * 复制XSSFSimpleShape类
     *
     * @param fromShape
     * @param drawing
     * @param anchor
     * @return
     */
    public static XSSFSimpleShape copyXSSFSimpleShape(XSSFSimpleShape fromShape, XSSFDrawing drawing, XSSFClientAnchor anchor) {
        XSSFSimpleShape toShape = drawing.createSimpleShape(anchor);
        CTShape ctShape = fromShape.getCTShape();
        CTShapeProperties ctShapeProperties = ctShape.getSpPr();
        CTLineProperties lineProperties = ctShapeProperties.isSetLn() ? ctShapeProperties.getLn() : ctShapeProperties.addNewLn();
        CTPresetLineDashProperties dashStyle = lineProperties.isSetPrstDash() ? lineProperties.getPrstDash() : CTPresetLineDashProperties.Factory.newInstance();
        STPresetLineDashVal.Enum dashStyleEnum = dashStyle.isSetVal() ? dashStyle.getVal() : STPresetLineDashVal.Enum.forInt(1);
        CTSolidColorFillProperties fill = lineProperties.isSetSolidFill() ? lineProperties.getSolidFill() : lineProperties.addNewSolidFill();
        CTSRgbColor rgb = fill.isSetSrgbClr() ? fill.getSrgbClr() : CTSRgbColor.Factory.newInstance();
        // 设置形状类型
        toShape.setShapeType(fromShape.getShapeType());
        // 设置线宽
        toShape.setLineWidth(lineProperties.getW() * 1.0 / Units.EMU_PER_POINT);
        // 设置线的风格
        toShape.setLineStyle(dashStyleEnum.intValue() - 1);
        // 设置线的颜色
        byte[] rgbBytes = rgb.getVal();
        if (rgbBytes == null) {
            toShape.setLineStyleColor(0, 0, 0);
        } else {
            toShape.setLineStyleColor(rgbBytes[0], rgbBytes[1], rgbBytes[2]);
        }
        return toShape;
    }

    /**
     * 复制XSSFConnector类
     *
     * @param fromShape
     * @param drawing
     * @param anchor
     * @return
     */
    public static XSSFConnector copyXSSFConnector(XSSFConnector fromShape, XSSFDrawing drawing, XSSFClientAnchor anchor) {
        XSSFConnector toShape = drawing.createConnector(anchor);
        CTConnector ctConnector = fromShape.getCTConnector();
        CTShapeProperties ctShapeProperties = ctConnector.getSpPr();
        CTLineProperties lineProperties = ctShapeProperties.isSetLn() ? ctShapeProperties.getLn() : ctShapeProperties.addNewLn();
        CTPresetLineDashProperties dashStyle = lineProperties.isSetPrstDash() ? lineProperties.getPrstDash() : CTPresetLineDashProperties.Factory.newInstance();
        STPresetLineDashVal.Enum dashStyleEnum = dashStyle.isSetVal() ? dashStyle.getVal() : STPresetLineDashVal.Enum.forInt(1);
        CTSolidColorFillProperties fill = lineProperties.isSetSolidFill() ? lineProperties.getSolidFill() : lineProperties.addNewSolidFill();
        CTSRgbColor rgb = fill.isSetSrgbClr() ? fill.getSrgbClr() : CTSRgbColor.Factory.newInstance();
        // 设置形状类型
        toShape.setShapeType(fromShape.getShapeType());
        // 设置线宽
        toShape.setLineWidth(lineProperties.getW() * 1.0 / Units.EMU_PER_POINT);
        // 设置线的风格
        toShape.setLineStyle(dashStyleEnum.intValue() - 1);
        // 设置线的颜色
        byte[] rgbBytes = rgb.getVal();
        if (rgbBytes == null) {
            toShape.setLineStyleColor(0, 0, 0);
        } else {
            toShape.setLineStyleColor(rgbBytes[0], rgbBytes[1], rgbBytes[2]);
        }
        return toShape;
    }

    /**
     * 复制单元格
     *
     * @param srcCell
     * @param distCell
     */
    public static void copyCell(Workbook workbook, Cell srcCell, Cell distCell) {
        CellStyle newStyle = workbook.createCellStyle();
        copyCellStyle(srcCell.getCellStyle(), newStyle, workbook);
        //样式
        distCell.setCellStyle(newStyle);
        //设置内容
        CellType srcCellType = srcCell.getCellTypeEnum();
        distCell.setCellType(srcCellType);
        if (srcCellType == NUMERIC) {
            if (DateUtil.isCellDateFormatted(srcCell)) {
                distCell.setCellValue(srcCell.getDateCellValue());
            } else {
                distCell.setCellValue(srcCell.getNumericCellValue());
            }
        } else if (srcCellType == CellType.STRING) {
            distCell.setCellValue(srcCell.getRichStringCellValue());
        } else if (srcCellType == CellType.BOOLEAN) {
            distCell.setCellValue(srcCell.getBooleanCellValue());
        } else if (srcCellType == CellType.ERROR) {
            distCell.setCellErrorValue(srcCell.getErrorCellValue());
        } else if (srcCellType == CellType.FORMULA) {
            distCell.setCellFormula(srcCell.getCellFormula());
        } else {
        }
    }

    /**
     * 复制一个单元格样式到目的单元格样式
     *
     * @param fromStyle
     * @param toStyle
     */
    public static void copyCellStyle(CellStyle fromStyle, CellStyle toStyle, Workbook workbook) {
        //水平垂直对齐方式
        toStyle.setAlignment(fromStyle.getAlignmentEnum());
        toStyle.setVerticalAlignment(fromStyle.getVerticalAlignmentEnum());
        //边框和边框颜色
        toStyle.setBorderBottom(fromStyle.getBorderBottomEnum());
        toStyle.setBorderLeft(fromStyle.getBorderLeftEnum());
        toStyle.setBorderRight(fromStyle.getBorderRightEnum());
        toStyle.setBorderTop(fromStyle.getBorderTopEnum());
        toStyle.setTopBorderColor(fromStyle.getTopBorderColor());
        toStyle.setBottomBorderColor(fromStyle.getBottomBorderColor());
        toStyle.setRightBorderColor(fromStyle.getRightBorderColor());
        toStyle.setLeftBorderColor(fromStyle.getLeftBorderColor());
        //背景和前景
        if (fromStyle instanceof XSSFCellStyle) {
            XSSFCellStyle xssfToStyle = (XSSFCellStyle) toStyle;
            xssfToStyle.setFillBackgroundColor(((XSSFCellStyle) fromStyle).getFillBackgroundColorColor());
            xssfToStyle.setFillForegroundColor(((XSSFCellStyle) fromStyle).getFillForegroundColorColor());
        } else {
            toStyle.setFillBackgroundColor(fromStyle.getFillBackgroundColor());
            toStyle.setFillForegroundColor(fromStyle.getFillForegroundColor());
        }
        toStyle.setDataFormat(fromStyle.getDataFormat());
        toStyle.setFillPattern(fromStyle.getFillPatternEnum());
        if (fromStyle instanceof XSSFCellStyle) {
            toStyle.setFont(((XSSFCellStyle) fromStyle).getFont());
        } else if (fromStyle instanceof HSSFCellStyle) {
            toStyle.setFont(((HSSFCellStyle) fromStyle).getFont(workbook));
        }
        toStyle.setHidden(fromStyle.getHidden());
        //首行缩进
        toStyle.setIndention(fromStyle.getIndention());
        toStyle.setLocked(fromStyle.getLocked());
        //旋转
        toStyle.setRotation(fromStyle.getRotation());
        toStyle.setWrapText(fromStyle.getWrapText());
    }
}
