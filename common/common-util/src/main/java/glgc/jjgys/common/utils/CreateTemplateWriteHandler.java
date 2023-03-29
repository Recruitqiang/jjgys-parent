package glgc.jjgys.common.utils;

import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

public class CreateTemplateWriteHandler implements SheetWriteHandler {
    /**
     * 第一行内容
     */
    private String firstTitle;


    /**
     * 实体模板类的行高
     */
    private int height;


    /**
     * 实体类 最大的列坐标 从0开始算
     */
    private int  lastCellIndex;



    public CreateTemplateWriteHandler(String firstTitle, int height, int cellCounts) {
        this.firstTitle = firstTitle;
        this.height = height;
        this.lastCellIndex = cellCounts;

    }

    @Override
    public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {

    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        Workbook workbook = writeWorkbookHolder.getWorkbook();
        Sheet sheet = workbook.getSheetAt(0);
        Row row1 = sheet.createRow(0);
        row1.setHeight((short) height);
        //字体样式
        Font font = workbook.createFont();
        font.setColor((short)2);
        Cell cell = row1.createCell(0);

        //单元格样式
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setFont(font);
        cellStyle.setWrapText(true);
        cell.setCellStyle(cellStyle);

        //设置单元格内容
        cell.setCellValue(firstTitle);


        //合并单元格  --> 起始行， 终止行   ，起始列，终止列
        sheet.addMergedRegionUnsafe(new CellRangeAddress(0, 0, 0, lastCellIndex));

    }
}
