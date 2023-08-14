package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcSdgcDmpzd;
import glgc.jjgys.model.project.JjgFbgcSdgcZtkd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.sdgc.JjgFbgcSdgcDmpzdVo;
import glgc.jjgys.model.projectvo.sdgc.JjgFbgcSdgcZtkdVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcSdgcDmpzdMapper;
import glgc.jjgys.system.service.JjgFbgcQlgcSbBhchdService;
import glgc.jjgys.system.service.JjgFbgcSdgcDmpzdService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-03-26
 */
@Service
public class JjgFbgcSdgcDmpzdServiceImpl extends ServiceImpl<JjgFbgcSdgcDmpzdMapper, JjgFbgcSdgcDmpzd> implements JjgFbgcSdgcDmpzdService {

    @Autowired
    private JjgFbgcSdgcDmpzdMapper jjgFbgcSdgcDmpzdMapper;


    @Value(value = "${jjgys.path.filepath}")
    private String filepath;


    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获取数据
        QueryWrapper<JjgFbgcSdgcDmpzd> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("sdmc","zh");
        List<JjgFbgcSdgcDmpzd> data = jjgFbgcSdgcDmpzdMapper.selectList(wrapper);

        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"40隧道大面平整度.xlsx");
        //健壮性判断如果没有数据返回"请导入数据"
        if (data == null || data.size()==0){
            return;
        }else {
            //存放鉴定表的目录
            File fdir = new File(filepath + File.separator + proname + File.separator + htd);
            if (!fdir.exists()) {
                //创建文件根目录
                fdir.mkdirs();
            }
            File directory = new File("service-system/src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String name = "隧道大面平整度.xlsx";
            String path = reportPath + File.separator + name;
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            createTable(gettableNum(data.size()),wb);
            if(DBtoExcel(data,wb)){
                calculatePZDSheet(wb);
                for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                    JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
                }
                FileOutputStream fileOut = new FileOutputStream(f);
                wb.write(fileOut);
                fileOut.flush();
                fileOut.close();
            }
            out.close();
            wb.close();
        }


    }

    /**
     *
     * @param wb
     */
    private void calculatePZDSheet(XSSFWorkbook wb) {
        String sheetname = "二衬大面平整度";
        XSSFSheet sheet = wb.getSheet(sheetname);
        XSSFRow row = null;
        boolean flag = false;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        XSSFRow rowrecord = null;
        ArrayList<String> bridgename = new ArrayList<String>();
        ArrayList<XSSFRow> start = new ArrayList<XSSFRow>();
        ArrayList<XSSFRow> end = new ArrayList<XSSFRow>();
        String name = "";
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            // 计算
            if (!"".equals(row.getCell(4).toString()) && flag) {
                row.getCell(5).setCellFormula("IF("
                        +row.getCell(4).getReference()+"-"
                        +row.getCell(3).getReference()+"<=0,\"√\",\"\")");//F6=IF(E6-D6<=0,"√","")
                row.getCell(6).setCellFormula("IF("
                        +row.getCell(4).getReference()+"-"
                        +row.getCell(3).getReference()+">0,\"×\",\"\")");//G6=IF(E6-D6>0,"×","")
            }
            // 到下一个表了
            if (!"".equals(row.getCell(0).toString()) && row.getCell(0).toString().contains("鉴定表") && flag) {
                flag = false;
            }
            // 下一座桥
            if (flag && !"".equals(row.getCell(0).toString()) && !name.equals(row.getCell(0).toString())) {
                if (start.size() > 0) {
                    rowrecord = start.get(0);
                    setTotalData(wb,sheetname, rowrecord, start, end, 4, 5, 6,7,8, 9,10);
                }
                start.clear();
                end.clear();
                name = row.getCell(0).toString();
                rowstart = sheet.getRow(i);
                rowrecord = rowstart;

                rowend = sheet.getRow(getCellEndRow(sheet, i, 0));
                start.add(rowstart);
                end.add(rowend);
            }
            // 同一座桥，但是在不同的单元格
            else if (flag && !"".equals(row.getCell(0).toString())
                    && name.equals(row.getCell(0).toString())
                    && rowstart != row) {
                rowstart = sheet.getRow(i);
                rowend = sheet.getRow(getCellEndRow(sheet, i, 0));
                start.add(rowstart);
                end.add(rowend);
            }
            // 可以计算啦
            if ("隧道名称".equals(row.getCell(0).toString())) {
                flag = true;
                i += 1;
                rowstart = sheet.getRow(i + 1);
                rowrecord = rowstart;
                rowend = sheet.getRow(getCellEndRow(sheet, i + 1, 0));
                if (name.equals(rowstart.getCell(0).toString())) {
                    start.add(rowstart);
                    end.add(rowend);
                } else {
                    if (start.size() > 0) {
                        rowrecord = start.get(0);
                        setTotalData(wb,sheetname, rowrecord, start, end, 4, 5, 6,7,8, 9,10);
                    }
                    start.clear();
                    end.clear();
                    name = rowstart.getCell(0).toString();
                    start.add(rowstart);
                    end.add(rowend);
                    bridgename.add(rowstart.getCell(0).toString());
                }
            }
        }
        if (start.size() > 0) {
            rowrecord = start.get(0);
            setTotalData(wb,sheetname, rowrecord, start, end, 4, 5, 6,7,8, 9,10);
        }
    }

    /**
     *
     * @param sheet
     * @param cellstartrow
     * @param cellstartcol
     * @return
     */
    public int getCellEndRow(XSSFSheet sheet, int cellstartrow, int cellstartcol) {
        int sheetmergerCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetmergerCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            if (cellstartrow == ca.getFirstRow()
                    && cellstartcol == ca.getFirstColumn()) {
                return ca.getLastRow();
            }
        }
        return cellstartrow;
    }

    /**
     *
     * @param wb
     * @param sheetname
     * @param rowrecord
     * @param start
     * @param end
     * @param c1
     * @param c2
     * @param c3
     * @param s1
     * @param s2
     * @param s3
     * @param s4
     */
    public void setTotalData(XSSFWorkbook wb,String sheetname, XSSFRow rowrecord, ArrayList<XSSFRow> start, ArrayList<XSSFRow> end, int c1, int c2, int c3, int s1, int s2, int s3, int s4) {
        XSSFSheet sheet = wb.getSheet(sheetname);
        XSSFCellStyle cellstyle = wb.createCellStyle();
        cellstyle.setAlignment(HorizontalAlignment.CENTER); // 居中
        String H = "";
        String I = "";
        String J = "";
        for (int index = 0; index < start.size(); index++) {
            H += "COUNT(" + start.get(index).getCell(c1).getReference() + ":"
                    + end.get(index).getCell(c1).getReference() + ")+";
            I += "COUNTIF(" + start.get(index).getCell(c2).getReference() + ":"
                    + end.get(index).getCell(c2).getReference() + ",\"√\")+";
            J += "COUNTIF(" + start.get(index).getCell(c3).getReference() + ":"
                    + end.get(index).getCell(c3).getReference() + ",\"×\")+";
        }
        setTotalTitle(sheet.getRow(rowrecord.getRowNum() - 1), cellstyle, s1, s2, s3, s4);
        rowrecord.createCell(s1).setCellFormula(H.substring(0, H.length() - 1));// H=COUNT(E216:E245)
        rowrecord.getCell(s1).setCellStyle(cellstyle);
        rowrecord.createCell(s2).setCellFormula(I.substring(0, I.length() - 1));// I=COUNTIF(F41:F60,"√")
        rowrecord.getCell(s2).setCellStyle(cellstyle);

        rowrecord.createCell(s3).setCellFormula(J.substring(0, J.length() - 1));// J=COUNTIF(G216:G245,"×")
        rowrecord.getCell(s3).setCellStyle(cellstyle);
        rowrecord.createCell(s4).setCellFormula(
                rowrecord.getCell(s2).getReference() + "/"
                        + rowrecord.getCell(s1).getReference() + "*100");// K==I41/H41*100
        rowrecord.getCell(s4).setCellStyle(cellstyle);

    }

    /**
     *
     * @param rowtitle
     * @param cellstyle
     * @param s1
     * @param s2
     * @param s3
     * @param s4
     */
    public void setTotalTitle(XSSFRow rowtitle, XSSFCellStyle cellstyle, int s1, int s2, int s3, int s4) {
        rowtitle.createCell(s1).setCellStyle(cellstyle);
        rowtitle.getCell(s1).setCellValue("总点数");
        rowtitle.createCell(s2).setCellStyle(cellstyle);
        rowtitle.getCell(s2).setCellValue("合格点数");
        rowtitle.createCell(s3).setCellStyle(cellstyle);
        rowtitle.getCell(s3).setCellValue("不合格点数");
        rowtitle.createCell(s4).setCellStyle(cellstyle);
        rowtitle.getCell(s4).setCellValue("合格率");
    }


    /**
     *
     * @param data
     * @param wb
     * @return
     * @throws ParseException
     */
    private boolean DBtoExcel(List<JjgFbgcSdgcDmpzd> data, XSSFWorkbook wb) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        XSSFSheet sheet = wb.getSheet("二衬大面平整度");
        String testtime = simpleDateFormat.format(data.get(0).getJcsj());
        String name = data.get(0).getSdmc();
        String zy = data.get(0).getZh().substring(0,1);
        int index = 0;
        int tableNum = 0;
        fillTitleCellData(sheet, tableNum, data.get(0));
        for(JjgFbgcSdgcDmpzd row:data){
            if(name.equals(row.getSdmc()) && zy.equals(row.getZh().substring(0,1))){
                testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
                if(index/10 == 3){
                    sheet.getRow(tableNum*35+2).getCell(5).setCellValue(testtime);
                    testtime = simpleDateFormat.format(row.getJcsj());
                    tableNum ++;
                    fillTitleCellData(sheet, tableNum, row);
                    index = 0;
                }
                if(index == 0){
                    sheet.getRow(tableNum*35 + 5).getCell(0).setCellValue(name);//setCellValue(name+"("+(ZY.equals("Z")?"左线":"右线")+")");
                }
                fillCommonCellData(sheet, tableNum, index+5, row);
                index++;
            }
            else{
                name = row.getSdmc();
                zy = row.getZh().substring(0,1);
                index = 0;
                sheet.getRow(tableNum*35+2).getCell(5).setCellValue(testtime);
                testtime = simpleDateFormat.format(row.getJcsj());
                tableNum ++;
                fillTitleCellData(sheet, tableNum, row);
                index = 0;
                sheet.getRow(tableNum*35 + 5 + index).getCell(0).setCellValue(name);//setCellValue(name+"("+(ZY.equals("Z")?"左线":"右线")+")");
                fillCommonCellData(sheet, tableNum, index+5, row);
                index ++;
            }
        }
        sheet.getRow(tableNum*35+2).getCell(5).setCellValue(testtime);
        return true;

    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param row
     */
    private void fillTitleCellData(XSSFSheet sheet, int tableNum,JjgFbgcSdgcDmpzd row) {
        System.out.println(tableNum);
        sheet.getRow(tableNum*35+1).getCell(1).setCellValue(row.getProname());
        sheet.getRow(tableNum*35+1).getCell(5).setCellValue(row.getHtd());
        sheet.getRow(tableNum*35+2).getCell(1).setCellValue(row.getFbgc());

    }


    /**
     *
     * @param sheet
     * @param tableNum
     * @param index
     * @param row
     */
    private void fillCommonCellData(XSSFSheet sheet, int tableNum, int index, JjgFbgcSdgcDmpzd row) {
        if((index+1)%6 == 0){
            sheet.getRow(tableNum*35+index).getCell(1).setCellValue(row.getZh());
        }
        if((index+1)%3 == 0 && (index+1)%6 < 3){
            sheet.getRow(tableNum*35+index).getCell(2).setCellValue("左边墙");
        }else if((index+1)%3 == 0 && (index+1)%6 >= 3){
            sheet.getRow(tableNum*35+index).getCell(2).setCellValue("右边墙");
        }
        sheet.getRow(tableNum*35+index).getCell(3).setCellValue(Double.valueOf(row.getYxps()));
        sheet.getRow(tableNum*35+index).getCell(4).setCellValue(Double.valueOf(row.getScz()));
    }


    /**
     *
     * @param tableNum
     * @param wb
     */
    private void createTable(int tableNum, XSSFWorkbook wb) {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            RowCopy.copyRows(wb, "二衬大面平整度", "二衬大面平整度", 0, 34, i * 35);
        }
        if(record > 1)
            wb.setPrintArea(wb.getSheetIndex("二衬大面平整度"), 0, 6, 0, record * 35-1);
    }

    /**
     *
     * @param size
     * @return
     */
    private int gettableNum(int size) {
        return size%30 ==0 ? size/30 : size/30+1;
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        //String fbgc = commonInfoVo.getFbgc();
        String sheetname = "二衬大面平整度";

        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        //获取鉴定表文件
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"40隧道大面平整度.xlsx");
        if(!f.exists()){
            return null;
        }else {
            XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));
            //读取工作表
            XSSFSheet slSheet = xwb.getSheet(sheetname);
            XSSFCell xmname = slSheet.getRow(1).getCell(1);//陕西高速
            XSSFCell htdname = slSheet.getRow(1).getCell(5);//LJ-1
            XSSFCell hd = slSheet.getRow(2).getCell(1);//涵洞
            List<Map<String,Object>> mapList = new ArrayList<>();
            Map<String,Object> jgmap = new HashMap<>();
            if(proname.equals(xmname.toString()) && htd.equals(htdname.toString())){
                slSheet.getRow(5).getCell(7).setCellType(CellType.STRING);
                slSheet.getRow(5).getCell(8).setCellType(CellType.STRING);
                slSheet.getRow(5).getCell(9).setCellType(CellType.STRING);
                slSheet.getRow(5).getCell(10).setCellType(CellType.STRING);

                jgmap.put("总点数",decf.format(Double.valueOf(slSheet.getRow(5).getCell(7).getStringCellValue())));
                jgmap.put("合格点数",decf.format(Double.valueOf(slSheet.getRow(5).getCell(8).getStringCellValue())));
                jgmap.put("不合格点数",decf.format(Double.valueOf(slSheet.getRow(5).getCell(9).getStringCellValue())));
                jgmap.put("合格率",df.format(Double.valueOf(slSheet.getRow(5).getCell(10).getStringCellValue())));
                mapList.add(jgmap);
                return mapList;
            }else {
                return null;
            }

        }
    }

    @Override
    public void exportsddmpzd(HttpServletResponse response) {
        String fileName = "03隧道大面平整度实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcSdgcDmpzdVo()).finish();

    }

    @Override
    public void importsddmpzd(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcSdgcDmpzdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcSdgcDmpzdVo>(JjgFbgcSdgcDmpzdVo.class) {
                                @Override
                                public void handle(List<JjgFbgcSdgcDmpzdVo> dataList) {
                                    for(JjgFbgcSdgcDmpzdVo sdgcDmpzdVo: dataList)
                                    {
                                        JjgFbgcSdgcDmpzd fbgcSdgcDmpzd = new JjgFbgcSdgcDmpzd();
                                        BeanUtils.copyProperties(sdgcDmpzdVo,fbgcSdgcDmpzd);
                                        fbgcSdgcDmpzd.setCreatetime(new Date());
                                        fbgcSdgcDmpzd.setProname(commonInfoVo.getProname());
                                        fbgcSdgcDmpzd.setHtd(commonInfoVo.getHtd());
                                        fbgcSdgcDmpzd.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcSdgcDmpzdMapper.insert(fbgcSdgcDmpzd);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }

    @Override
    public List<Map<String, Object>> lookjg(CommonInfoVo commonInfoVo) {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String sheetname = "二衬大面平整度";
        DecimalFormat decf = new DecimalFormat("0.##");
        //获取鉴定表文件
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "40隧道大面平整度.xlsx");
        if (!f.exists()) {
            return null;
        }else {
            List<Map<String, Object>> data = new ArrayList<>();

            try (FileInputStream fis = new FileInputStream(f)) {
                Workbook workbook = WorkbookFactory.create(fis);
                Sheet sheet = workbook.getSheet(sheetname); // 假设数据在第一个工作表中

                DataFormatter dataFormatter = new DataFormatter();
                FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();

                for (int i = 1; i <= sheet.getLastRowNum(); i++) { // 循环每一行，从第3行开始（忽略表头）

                    Row row = sheet.getRow(i);
                    Cell cell0 = getMergedCell(sheet, i, 0); // 获取合并单元格 // 第0列
                    Cell cell13 = row.getCell(3); // 第0列
                    Cell cell8 = row.getCell(7); // 第0列
                    Cell cell9 = row.getCell(8); // 第34列


                    if (cell0.getStringCellValue().equals("隧道名称") && dataFormatter.formatCellValue(cell8, formulaEvaluator).equals("总点数") ) { // 判断是否不为空
                        Cell nextRowCell8 = sheet.getRow(i + 1).getCell(7); // 下一行的第0列
                        Cell nextRowCell9 = sheet.getRow(i + 1).getCell(8); // 下一行的第34列

                        Cell nextRowCell0 = sheet.getRow(i + 1).getCell(0); // 下一行的第34列
                        Cell nextRowCell3 = sheet.getRow(i + 1).getCell(3); // 下一行的第34列

                        String data8 = dataFormatter.formatCellValue(nextRowCell8, formulaEvaluator);
                        String data9 = dataFormatter.formatCellValue(nextRowCell9, formulaEvaluator);

                        String data0 = nextRowCell0.getStringCellValue();
                        String data3 = dataFormatter.formatCellValue(nextRowCell3, formulaEvaluator);

                        Map map = new HashMap();
                        map.put("qlmc",data0);
                        map.put("zds",data8);
                        map.put("hgds",data9);
                        map.put("sjqd",data3);

                        data.add(map);

                    }
                }

            } catch (Exception e) {
                e.printStackTrace();
            }
            return data;
        }


    }

    private static Cell getMergedCell(Sheet sheet, int rowIndex, int columnIndex) {
        for (CellRangeAddress range : sheet.getMergedRegions()) {
            if (range.isInRange(rowIndex, columnIndex)) {
                Row mergedRow = sheet.getRow(range.getFirstRow());
                return mergedRow.getCell(range.getFirstColumn());
            }
        }
        return sheet.getRow(rowIndex).getCell(columnIndex);
    }
}
