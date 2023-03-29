package glgc.jjgys.system.utils;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.merge.LoopMergeStrategy;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.WriteTable;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.alibaba.excel.write.style.row.SimpleRowHeightStyleStrategy;
import glgc.jjgys.system.exception.JjgysException;
import org.apache.poi.ss.usermodel.*;

import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class HdgqdEasyExcelUtils {

    /**
     * 下载 excel
     * @param response
     * @param list 主体表格数据
     * @throws Exception
     */
    public static void download(HttpServletResponse response, List<?> list) throws Exception {
        // 总共 28 列
        int columnTotal = 28;

        // 首行标题
        List<List<String>> titleData = new ArrayList<>();
        List<String> title = Arrays.asList("混凝土强度质量鉴定表（回弹法）");
        titleData.add(title);

        // 中间抬头部分
        List<List<String>> secondData = new ArrayList<>();
        // 合并前3列, 合并后5列
        List<String> second1 = Arrays.asList("项目名称：","", "合同段：");
        List<String> second2 = Arrays.asList("分部工程名称：","", "检测日期：");
        secondData.add(second1);
        secondData.add(second2);

        excelCreate(response, list, columnTotal, titleData, secondData);
    }

    /**
     * excel 表格生成
     *
     * @param response
     * @param list 主体表格数据
     * @param columnTotal 总列数
     * @param titleData 表头数据
     * @param secondData 第二部分抬头数据
     * @throws Exception
     */
    private static void excelCreate(HttpServletResponse response, List<?> list, int columnTotal, List<List<String>> titleData,
                                    List<List<String>> secondData) throws Exception {
        /*String fileName = URLEncoder.encode("路基涵洞砼强度" + DateTimeFormatter.BASIC_ISO_DATE.format(LocalDate.now()), "UTF-8")
                .replaceAll("\\+", "%20");
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");*/
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        try {
            ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream()).build();
            WriteSheet sheet = EasyExcel.writerSheet(0,"sheetName").build();
            // 按每20行数据进行分页
            int pageNum = (int) Math.ceil(list.size() / 20d);
            if (pageNum == 0) {
                pageNum = 1;
            }
            for (int i = 1; i <= pageNum; i++) {
                WriteTable titleTable = titleTableCreate(columnTotal);
                WriteTable secondTable = secondTableCreate(columnTotal);
                WriteTable headTable = headTableCreate(i + 1);

                excelWriter.write(titleData, sheet, titleTable);
                excelWriter.write(secondData, sheet, secondTable);

                int startIndex = (i - 1) * 5;
                int endIndex = Math.min(startIndex + 5, list.size());
                excelWriter.write(list.subList(startIndex, endIndex), sheet, headTable);
            }
            excelWriter.finish();
        }catch (Exception e){
            throw new JjgysException(20001,"表格生成出错");
        }
    }

    /**
     * 生成标题表格
     *
     * @param columnTotal
     * @return
     */
    private static WriteTable titleTableCreate(int columnTotal) {

        // 合并单元格
        LoopMergeStrategy loopMergeStrategy = new LoopMergeStrategy(1, columnTotal, 0);

        // 样式
        WriteCellStyle titleStyle = new WriteCellStyle();
        // 水平居中
        titleStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        // 垂直居中
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 设置白色背景
        titleStyle.setFillForegroundColor(IndexedColors.WHITE1.getIndex());

        // 设置字体格式
        WriteFont font = new WriteFont();
        font.setBold(Boolean.TRUE);
        font.setFontHeightInPoints((short)14);
        font.setFontName("宋体");
        titleStyle.setWriteFont(font);

        HorizontalCellStyleStrategy titleStyleStrategy = new HorizontalCellStyleStrategy(titleStyle, titleStyle);

        // 生成表格
        WriteTable titleTable = EasyExcel.writerTable(0)
                .registerWriteHandler(new SimpleRowHeightStyleStrategy((short)25, (short)25))
                .registerWriteHandler(loopMergeStrategy)
                .registerWriteHandler(titleStyleStrategy)
                .needHead(Boolean.FALSE)
                .build();
        return titleTable;
    }

    /**
     * 生成抬头表格
     *
     * @param columnTotal
     * @return
     */
    private static WriteTable secondTableCreate(int columnTotal) {

        // 合并单元格
        LoopMergeStrategy loopMergeStrategy1 = new LoopMergeStrategy(1, 3, 0);
        LoopMergeStrategy loopMergeStrategy2 = new LoopMergeStrategy(1, columnTotal - 3, 3);

        // 样式
        WriteCellStyle secondStyle = new WriteCellStyle();
        // 水平居中
        secondStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        // 垂直居中
        secondStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 设置白色背景
        secondStyle.setFillForegroundColor(IndexedColors.WHITE1.getIndex());

        // 设置字体格式
        WriteFont font = new WriteFont();
        font.setBold(Boolean.FALSE);
        font.setFontHeightInPoints((short)12);
        font.setFontName("宋体");
        secondStyle.setWriteFont(font);

        HorizontalCellStyleStrategy secondStyleStrategy = new HorizontalCellStyleStrategy(secondStyle, secondStyle);

        // 生成表格
        WriteTable secondTable = EasyExcel.writerTable(1)
                .registerWriteHandler(new SimpleRowHeightStyleStrategy((short)25, (short)25))
                .registerWriteHandler(loopMergeStrategy1)
                .registerWriteHandler(loopMergeStrategy2)
                .registerWriteHandler(secondStyleStrategy)
                .needHead(Boolean.FALSE)
                .build();
        return secondTable;
    }

    /**
     * 生成主体数据表格
     *
     * @param tableNum
     * @return
     */
    private static WriteTable headTableCreate(int tableNum) {

        List<List<String>> head = new ArrayList<>();
        List<String> headList = Arrays.asList("桩号/结构名称", "部位", "回弹仪测定值", "平均值", "回弹角度", "浇筑面","碳化深度","是否泵送","换算强度（MPa）","设计强度","结论");
        List<String> subList1 = Arrays.asList("1", "2", "3","4", "5", "6","7", "8", "9","10", "11", "12","13", "14", "15","16");
        List<String> subList2 = Arrays.asList("√","×");
        headList.forEach(title -> {
            if ("回弹仪测定值".equals(title)) {
                subList1.forEach(sub -> head.add(Arrays.asList(title, sub)));
                return;
            }
            if ("结论".equals(title)) {
                subList2.forEach(sub -> head.add(Arrays.asList(title, sub)));
                return;
            }
            head.add(Arrays.asList(title, title));
        });

        HorizontalCellStyleStrategy headStyleStrategy = headStyleStrategy();

        // 生成表格
        WriteTable titleTable = EasyExcel.writerTable(tableNum)
               /* .registerWriteHandler(new AbstractColumnWidthStyleStrategy() {
                    @Override
                    protected void setColumnWidth(WriteSheetHolder writeSheetHolder, List<CellData> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
                        Sheet sheet = writeSheetHolder.getSheet();
                        int columnIndex = cell.getColumnIndex();
                        if (columnIndex == 3 || columnIndex == 4 || columnIndex == 7) {
                            sheet.setColumnWidth(columnIndex, 7000);
                        } else {
                            sheet.setColumnWidth(columnIndex, 3600);
                        }
                    }
                })*/
                .registerWriteHandler(new SimpleRowHeightStyleStrategy((short)25, (short)25))
                .registerWriteHandler(headStyleStrategy)
                .head(head)
                .needHead(Boolean.TRUE)
                .build();
        return titleTable;
    }

    /**
     * 主体表格样式
     *
     * @return
     */
    private static HorizontalCellStyleStrategy headStyleStrategy() {

        // 主体表头样式
        WriteCellStyle headStyle = new WriteCellStyle();
        // 水平居中
        headStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        // 垂直居中
        headStyle.setVerticalAlignment(VerticalAlignment.CENTER);


        // 设置字体格式
        WriteFont font = new WriteFont();
        font.setBold(Boolean.FALSE);
        font.setFontHeightInPoints((short)12);
        font.setFontName("宋体");
        headStyle.setWriteFont(font);

        // 设置边框
        headStyle.setBorderLeft(BorderStyle.THIN);
        headStyle.setBorderRight(BorderStyle.THIN);
        headStyle.setBorderTop(BorderStyle.THIN);
        headStyle.setBorderBottom(BorderStyle.THIN);

        // 主题内容样式
        WriteCellStyle contentStyle = new WriteCellStyle();
        // 水平居中
        contentStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        // 垂直居中
        contentStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 自动换行
        contentStyle.setWrapped(true);

        // 设置字体格式
        font = new WriteFont();
        font.setBold(Boolean.FALSE);
        font.setFontHeightInPoints((short)10);
        font.setFontName("宋体");
        contentStyle.setWriteFont(font);

        // 设置边框
        contentStyle.setBorderLeft(BorderStyle.THIN);
        contentStyle.setBorderRight(BorderStyle.THIN);
        contentStyle.setBorderTop(BorderStyle.THIN);
        contentStyle.setBorderBottom(BorderStyle.THIN);


        HorizontalCellStyleStrategy headStyleStrategy = new HorizontalCellStyleStrategy(headStyle, contentStyle);
        return headStyleStrategy;
    }


}
