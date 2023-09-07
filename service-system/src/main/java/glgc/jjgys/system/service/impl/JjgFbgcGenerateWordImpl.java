package glgc.jjgys.system.service.impl;


import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import com.spire.doc.documents.HorizontalAlignment;
import com.spire.doc.documents.VerticalAlignment;
import com.spire.presentation.Cell;
import com.spire.xls.collections.WorksheetsCollection;
import glgc.jjgys.system.mapper.JjgFbgcGenerateWordMapper;
import glgc.jjgys.system.service.JjgFbgcGenerateWordService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import com.spire.doc.Document;
import com.spire.doc.Table;
import com.spire.doc.TableCell;
import com.spire.doc.fields.TextRange;
import com.spire.xls.CellRange;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;


import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;


@Slf4j
@Service
public class JjgFbgcGenerateWordImpl extends ServiceImpl<JjgFbgcGenerateWordMapper, Object> implements JjgFbgcGenerateWordService {

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @Override
    public void generateword(String proname) throws IOException {
        String excelFilePath = filespath+File.separator+proname+File.separator+"报告中表格.xlsx";
        File f = new File(filespath + File.separator + proname + File.separator + "报告.docx");
        File fdir = new File(filespath + File.separator + proname);
        if (!fdir.exists()) {
            //创建文件根目录
            fdir.mkdirs();
        }
        Document xw = null;
        File directory = new File("src/main/resources/static");
        String reportPath = directory.getCanonicalPath();
        String name = "报告.docx";
        String path = reportPath + File.separator + name;
        Files.copy(Paths.get(path), new FileOutputStream(f));
        FileInputStream out = new FileInputStream(f);
        xw = new Document(out);
        try {
            //读取Excel文件
            /*Workbook workbook = new Workbook();
            workbook.loadFromFile(excelFilePath);
            WorksheetsCollection sheetnum = workbook.getWorksheets();
            for (int j = 0; j < sheetnum.getCount(); j++) {
                String sheetname = sheetnum.get(j).getName();
                Worksheet sheet = workbook.getWorksheets().get(j);
                CellRange allocatedRange = sheet.getAllocatedRange();
                //删除没用数据的行
                CellRange dataRange = RemoveEmptyRows(sheet,allocatedRange);
                //复制到Word文档
                log.info("开始复制{}中的数据到word中",sheetname);
                //copyToWord(sheet.getAllocatedRange(), xw,f.getPath());
                copyToWord(dataRange, xw,f.getPath());
                log.info("{}中的数据复制完成",sheetname);
            }*/
        } catch (Exception e) {
            e.printStackTrace();
        }
        out.close();
        xw.close();
    }


    /*public CellRange RemoveEmptyRows(Worksheet sheet, CellRange allocatedRange) {
        CellRange dataRange = null;
        for (int i = allocatedRange.getLastRow(); i >= 1; i--) {
            boolean isRowEmpty = true;
            for (int col = 1; col <= allocatedRange.getLastColumn(); col++) {
                if (!sheet.getCellRange(i, col).getText().isEmpty()) {
                    isRowEmpty = false;
                    break;
                }
            }
            if (isRowEmpty) {
                sheet.deleteRow(i);
            } else {
                dataRange = sheet.getCellRange(i, 1, 1, allocatedRange.getLastColumn());
            }
        }
        return dataRange;
    }*/


    /**
     * @param cell
     * @param doc
     * @param path
     */
    public static void copyToWord(CellRange cell, Document doc, String path) {
        //添加表格
        Table table = doc.addSection().addTable(true);
        table.resetCells(cell.getRowCount(), cell.getColumnCount());
        //复制表格内容
        for (int r = 1; r <= cell.getRowCount(); r++) {
            for (int c = 1; c <= cell.getColumnCount(); c++) {
                CellRange xCell = cell.get(r, c);
                CellRange mergeArea = xCell.getMergeArea();
                //合并单元格
                if (mergeArea != null && mergeArea.getRow() == r && mergeArea.getColumn() == c) {
                    int rowIndex = mergeArea.getRow();
                    int columnIndex = mergeArea.getColumn();
                    int rowCount = mergeArea.getRowCount();
                    int columnCount = mergeArea.getColumnCount();

                    for (int m = 0; m < rowCount; m++) {
                        table.applyHorizontalMerge(rowIndex - 1 + m, columnIndex - 1, columnIndex + columnCount - 2);
                    }
                    table.applyVerticalMerge(columnIndex - 1, rowIndex - 1, rowIndex + rowCount - 2);
                }
                //复制内容
                TableCell wCell = table.getRows().get(r - 1).getCells().get(c - 1);
                if (!xCell.getDisplayedText().isEmpty()) {
                    TextRange textRange = wCell.addParagraph().appendText(xCell.getDisplayedText());
                    copyStyle(textRange, xCell, wCell);
                } else {
                    wCell.getCellFormat().setBackColor(xCell.getStyle().getColor());
                }
            }

        }
        doc.saveToFile(path,com.spire.doc.FileFormat.Docx);
    }



    /**
     *
     * @param wTextRange
     * @param xCell
     * @param wCell
     */
    private static void copyStyle(TextRange wTextRange, CellRange xCell, TableCell wCell) {
        //复制字体样式
        wTextRange.getCharacterFormat().setTextColor(xCell.getStyle().getFont().getColor());
        wTextRange.getCharacterFormat().setFontSize((float) xCell.getStyle().getFont().getSize());
        wTextRange.getCharacterFormat().setFontName(xCell.getStyle().getFont().getFontName());
        wTextRange.getCharacterFormat().setBold(xCell.getStyle().getFont().isBold());
        wTextRange.getCharacterFormat().setItalic(xCell.getStyle().getFont().isItalic());
        //复制背景色
        wCell.getCellFormat().setBackColor(xCell.getStyle().getColor());
        //复制排列方式
        switch (xCell.getHorizontalAlignment()) {
            case Left:
                wTextRange.getOwnerParagraph().getFormat().setHorizontalAlignment(HorizontalAlignment.Left);
                break;
            case Center:
                wTextRange.getOwnerParagraph().getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
                break;
            case Right:
                wTextRange.getOwnerParagraph().getFormat().setHorizontalAlignment(HorizontalAlignment.Right);
                break;
            default:
                break;
        }
        switch (xCell.getVerticalAlignment()) {
            case Bottom:
                wCell.getCellFormat().setVerticalAlignment(VerticalAlignment.Bottom);
                break;
            case Center:
                wCell.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                break;
            case Top:
                wCell.getCellFormat().setVerticalAlignment(VerticalAlignment.Top);
                break;
            default:
                break;
        }
    }

}
