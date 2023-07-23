package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLjgcLjcj;
import glgc.jjgys.model.project.JjgFbgcLjgcLjwc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.ljgc.JjgFbgcLjgcLjwcVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcLjgcLjwcMapper;
import glgc.jjgys.system.service.JjgFbgcLjgcLjwcService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.service.ProjectService;
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
import java.math.BigDecimal;
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
 * @since 2023-02-27
 */
@Service
public class JjgFbgcLjgcLjwcServiceImpl extends ServiceImpl<JjgFbgcLjgcLjwcMapper, JjgFbgcLjgcLjwc> implements JjgFbgcLjgcLjwcService {

    @Autowired
    private JjgFbgcLjgcLjwcMapper jjgFbgcLjgcLjwcMapper;
    @Autowired
    private ProjectService projectService;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;


    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException {
        XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获取数据
        QueryWrapper<JjgFbgcLjgcLjwc> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("xh");
        List<JjgFbgcLjgcLjwc> data = jjgFbgcLjgcLjwcMapper.selectList(wrapper);
        //查询保证率系数
        /*double bzlxs = 0;
        Integer level = projectService.getlevel(proname);
        if (level == 0){
            bzlxs = 2.0;
        }else {
            bzlxs = 1.645;
        }*/
        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"02路基弯沉(贝克曼梁法).xlsx");
        if (data == null || data.size()==0){
            return;
        }else {
            File fdir = new File(filepath+File.separator+proname+File.separator+htd);
            if(!fdir.exists()){
                //创建文件根目录
                fdir.mkdirs();
            }
            File directory = new File("service-system/src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String path =reportPath +File.separator+ "路基弯沉.xlsx";
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            int num = gettableNum(proname, htd, fbgc);
            createTable(num,wb);
            if(DBtoExcel(data,wb)){
                calculateTempDate(wb.getSheet("路基弯沉"),wb);
                ArrayList<String> totalref = getTotalMark(wb.getSheet("路基弯沉"));
                String time = getLastTime(wb.getSheet("路基弯沉"));

                //创建“评定单元”sheet
                createEvaluateTable(totalref,wb);

                //引用数据，然后计算
                completeTotleTable(wb.getSheet("评定单元"), totalref, time,wb,proname,htd,fbgc);

                for (int j = 0; j < wb.getNumberOfSheets(); j++) {   //表内公式  计算 显示结果
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
     * 将“评定单元”sheet数据引用填好，然后计算
     * @param sheet
     * @param ref
     */
    public void completeTotleTable(XSSFSheet sheet, ArrayList<String> ref, String time,XSSFWorkbook xwb,String proname,String htd,String fbgc) {
        XSSFCellStyle cellstyle = xwb.createCellStyle();
        XSSFFont font=xwb.createFont();
        font.setFontHeightInPoints((short)11);
        font.setFontName("宋体");
        cellstyle.setFont(font);
        cellstyle.setBorderBottom(BorderStyle.THIN); //下边框
        cellstyle.setBorderLeft(BorderStyle.THIN);//左边框
        cellstyle.setBorderTop(BorderStyle.THIN);//上边框
        cellstyle.setBorderRight(BorderStyle.THIN);//右边框
        cellstyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直
        cellstyle.setAlignment(HorizontalAlignment.CENTER);//水平
        cellstyle.setWrapText(true);//自动换行

        XSSFDataFormat df = xwb.createDataFormat();
        cellstyle.setDataFormat(df.getFormat("#,##0.0"));

        XSSFCellStyle cellstyle2 = xwb.createCellStyle();
        XSSFFont font2=xwb.createFont();
        font2.setFontHeightInPoints((short)9);
        font2.setFontName("宋体");
        cellstyle2.setFont(font2);
        cellstyle2.setBorderBottom(BorderStyle.THIN); //下边框
        cellstyle2.setBorderLeft(BorderStyle.THIN);//左边框
        cellstyle2.setBorderTop(BorderStyle.THIN);//上边框
        cellstyle2.setBorderRight(BorderStyle.THIN);//右边框
        cellstyle2.setVerticalAlignment(VerticalAlignment.CENTER);//垂直
        cellstyle2.setAlignment(HorizontalAlignment.CENTER);//水平
        cellstyle2.setWrapText(true);//自动换行
        XSSFRow row;
        int index = 0, rowstart = 5;
        /*
         * 填写“评定单元”sheet的表头数据
         */
        sheet.getRow(1).getCell(2).setCellValue(proname);
        sheet.getRow(1).getCell(7).setCellValue(htd);
        sheet.getRow(2).getCell(2).setCellValue(fbgc);
        sheet.getRow(2).getCell(7).setCellValue(time);
        double value = 0.0;
        for (int i = rowstart; i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if (index < ref.size()) {
                row.createCell(0).setCellValue(String.valueOf(index/3+1));
                row.getCell(0).setCellStyle(cellstyle);

                row.createCell(1).setCellFormula("路基弯沉"+"!" + ref.get(index));
                sheet.addMergedRegion(new CellRangeAddress(i, i, 1, 4));
                row.getCell(1).setCellStyle(cellstyle2);

                row.createCell(5).setCellFormula("路基弯沉"+"!" + ref.get(index + 1));
                row.getCell(5).setCellStyle(cellstyle);
                row.createCell(6).setCellFormula("路基弯沉"+"!" + ref.get(index + 2));
                row.getCell(6).setCellStyle(cellstyle);
                row.createCell(7).setCellFormula(
                        "IF(" + row.getCell(6).getReference()
                                + "=\"\",\"\",IF("
                                + row.getCell(6).getReference() + "<"
                                + row.getCell(5).getReference()
                                + ",\"√\",\"\"))");
                row.getCell(7).setCellStyle(cellstyle);
                // =IF(G6="","",IF(G6>=F6,"×",""))
                row.createCell(8).setCellFormula(
                        "IF(" + row.getCell(6).getReference()
                                + "=\"\",\"\",IF("
                                + row.getCell(6).getReference() + ">="
                                + row.getCell(5).getReference()
                                + ",\"×\",\"\"))");
                if("×".equals(row.getCell(8).getRawValue())){
                    row.getCell(8).setCellStyle(cellstyle);
                }
                row.getCell(8).setCellStyle(cellstyle);
                index += 3;
            }
            if ("合计".equals(row.getCell(0).toString())) {
                //System.out.println("zuihouyihang");
                row.getCell(4).setCellFormula(
                        "COUNT("
                                + sheet.getRow(rowstart).getCell(5)
                                .getReference() + ":"
                                + sheet.getRow(i - 1).getCell(5).getReference()
                                + ")");// =COUNT(F6:F27)

                row.getCell(6).setCellFormula(
                        "COUNTIF("
                                + sheet.getRow(rowstart).getCell(7)
                                .getReference() + ":"
                                + sheet.getRow(i - 1).getCell(7).getReference()
                                + ",\"√\")");// =COUNTIF(H6:H27,"√")

                row.getCell(8).setCellFormula(
                        row.getCell(6).getReference() + "/"
                                + row.getCell(4).getReference() + "*100");// =G28/E28*100
            }
        }
    }


    /**
     * 创建“评定单元”sheet
     * @param ref
     */
    public void createEvaluateTable(ArrayList<String> ref,XSSFWorkbook xwb){
        int record = 0;
        record = ref.size()/3;
        for(int i = 1; i < record/23+1; i++){
            if(i < record/23){
                RowCopy.copyRows(xwb, "评定单元", "评定单元", 5, 27, (i-1)*23+28);
            }
            else{
                RowCopy.copyRows(xwb, "评定单元", "评定单元", 5, 26, (i-1)*23+28);
            }
        }
        if(record/23 == 0){
            xwb.getSheet("评定单元").shiftRows(28, 28, -1, true , false);
        }
        RowCopy.copyRows(xwb, "source", "评定单元", 0, 1, (record/23+1)*23+4);
        xwb.setPrintArea(xwb.getSheetIndex("评定单元"), 0, 8, 0, (record/23+1)*23+4);
    }

    /**
     * 取得时间
     * @param sheet
     * @return
     */
    public String getLastTime(XSSFSheet sheet) {
        String time = "1900.1.1";
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        XSSFRow row;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if (row.getCell(0) != null && "分部工程名称".equals(row.getCell(0).toString()) && row.getCell(11) != null && !"".equals(row.getCell(11).toString())) {
                try {
                    Date dt1 = simpleDateFormat.parse(time);
                    Date dt2 = simpleDateFormat.parse(row.getCell(11).toString());
                    if(dt1.getTime() < dt2.getTime()){
                        time = row.getCell(11).toString();
                    }
                } catch (ParseException e) {
                    e.printStackTrace();
                }
            }
        }
        return time;
    }

    /**
     * 取得汇总表中所需数据的reference
     * @param sheet
     * @return
     */
    public ArrayList<String> getTotalMark(XSSFSheet sheet) {
        ArrayList<String> ref = new ArrayList<String>();
        XSSFRow row;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if ("桩号/部位".equals(row.getCell(0).toString()) && row.getCell(16) != null && !"".equals(row.getCell(16).toString())) {
                ref.add(row.getCell(2).getReference());
                // 验收弯沉值（0.01mm）
                ref.add(row.getCell(11).getReference());
            }
            // 弯沉代表值(0.01mm)
            if ("弯沉代表值(0.01mm)".equals(row.getCell(8).toString()) && row.getCell(16) != null && !"".equals(row.getCell(16).toString())) {
                ref.add(row.getCell(11).getReference());
            }
        }
        return ref;
    }


    /**
     * 将sheet主题表格中的：平均弯沉值(0.01mm)，弯沉代表值(0.01mm)，结论三项进行填写
     * @param rowNum
     * @param cell
     */
    public void fillSheetBody(int rowNum,XSSFWorkbook xwb, XSSFCell... cell) {
        /*
         * 前两项为四个参数，后一项为三个参数 平均弯沉值(0.01mm)L5=IF(Q3<L4,Q8,Q3)
         * 弯沉代表值(0.01mm)L6=IF(Q4>L4,Q9,Q4) 结论L7=IF(L6>L4,"×","√")
         */
        XSSFCellStyle cellstyle = xwb.createCellStyle();
        XSSFFont font=xwb.createFont();
        font.setFontHeightInPoints((short)9);
        font.setColor(Font.COLOR_RED);
        font.setFontName("宋体");
        cellstyle.setFont(font);
        cellstyle.setBorderBottom(BorderStyle.THIN); //下边框
        cellstyle.setBorderLeft(BorderStyle.THIN);//左边框
        cellstyle.setBorderTop(BorderStyle.THIN);//上边框
        cellstyle.setBorderRight(BorderStyle.THIN);//右边框
        cellstyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直
        cellstyle.setAlignment(HorizontalAlignment.CENTER);//水平
        cellstyle.setWrapText(true);//自动换行
        XSSFCell cell1, cell2, cell3, cell4;
        if (cell.length == 4) {
            if(rowNum == 5){
                cell1 = cell[0];
                cell2 = cell[1];
                cell3 = cell[2];
                cell4 = cell[3];
                cell1.setCellFormula("IF(" + cell2.getReference() + ">"
                        + cell3.getReference() + "," + cell4.getReference() + ","
                        + cell2.getReference() + ")");
            }else{
                cell1 = cell[0];
                cell2 = cell[1];
                cell3 = cell[2];
                cell4 = cell[3];
                cell1.setCellFormula("IF(" + cell2.getReference() + ">"
                        + cell3.getReference() + "," + cell4.getReference() + ","
                        + cell2.getReference() + ")");
            }
        }
        if (cell.length == 3) {
            cell1 = cell[0];
            cell2 = cell[1];
            cell3 = cell[2];
            cell1.setCellFormula("IF(" + cell2.getReference() + ">"
                    + cell3.getReference() + ",\"×\",\"√\")");
            XSSFFormulaEvaluator evaluate = new XSSFFormulaEvaluator(xwb);
            if("×".equals(cell1.getRawValue())){
                cell1.setCellStyle(cellstyle);
            }
        }
    }


    /**
     * 计算“特异点数”的函数
     * 就是将大于等于“特异值上限”和小于等于“特异值下限”的数剔除后，计算平均值
     * @param upper
     * @param lower
     * @param left_top
     * @param right_bottom
     * @param evaluator
     * @return
     */
    public int getArrayCount(BigDecimal upper, BigDecimal lower, String left_top, String right_bottom, XSSFFormulaEvaluator evaluator,XSSFWorkbook xwb) {
        BigDecimal value = null;
        int count = 0;
        XSSFSheet sheet = xwb.getSheet("路基弯沉");
        XSSFRow row = null;
        int firstrow = Integer.valueOf(left_top.replaceAll("[^0-9]", "")) - 1;
        int lastrow = Integer.valueOf(right_bottom.replaceAll("[^0-9]", "")) - 1;
        for(int i = firstrow; i <= lastrow; i ++){
            row = sheet.getRow(i);
            if(row == null){
                continue;
            }
            for (int j = 18; j <= 21; j++) {
                if(row.getCell(j) == null || "".equals(row.getCell(j).toString())){
                    continue;
                }
                value = new BigDecimal(evaluator.evaluate(row.getCell(j)).getNumberValue());
                if(upper.compareTo(value) > 0 && lower.compareTo(value) < 0){
                    count ++;
                }

            }
        }
        if(count <= 1){
            return 0;
        }
        return count;
    }

    /**
     * 计算“舍后均方差”的函数
     * 就是将大于等于“特异值上限”和小于等于“特异值下限”的数剔除后，计算均方差
     * @param upper
     * @param lower
     * @param average
     * @param left_top
     * @param right_bottom
     * @param evaluator
     * @param xwb
     * @return
     */
    public BigDecimal getArraySTDEVA(BigDecimal upper, BigDecimal lower, BigDecimal average, String left_top, String right_bottom, XSSFFormulaEvaluator evaluator,XSSFWorkbook xwb) {
        BigDecimal res = new BigDecimal(0);
        BigDecimal value = null;
        int count = 0;
        XSSFSheet sheet = xwb.getSheet("路基弯沉");
        XSSFRow row = null;
        int firstrow = Integer.valueOf(left_top.replaceAll("[^0-9]", "")) - 1;
        int lastrow = Integer.valueOf(right_bottom.replaceAll("[^0-9]", "")) - 1;
        BigDecimal std2 = new BigDecimal(0);
        double single = 0.0;
        BigDecimal std = new BigDecimal(0);
        for(int i = firstrow; i <= lastrow; i ++){
            row = sheet.getRow(i);
            if(row == null){
                continue;
            }
            for (int j = 18; j <= 21; j++) {
                if(row.getCell(j) == null || "".equals(row.getCell(j).toString())){
                    continue;
                }
                value = new BigDecimal(evaluator.evaluate(row.getCell(j)).getNumberValue());
                if(upper.compareTo(value) > 0 && lower.compareTo(value) < 0){
                    single = Math.pow(value.subtract(average).doubleValue(), 2);
                    std2 = std2.add(new BigDecimal(single));
                    count ++;
                }
            }
        }
        System.out.println("count = " + count);
        if(count <= 1){
            return res;
        }
        std = new BigDecimal(Math.sqrt(std2.divide(new BigDecimal(count-1), 1, BigDecimal.ROUND_HALF_UP).doubleValue())).setScale(1, BigDecimal.ROUND_HALF_UP);
        return std;
    }

    /**
     * 计算“舍后平均值”的函数
     * 就是将大于等于“特异值上限”和小于等于“特异值下限”的数剔除后，计算平均值
     * @param upper
     * @param lower
     * @param left_top
     * @param right_bottom
     * @param evaluator
     * @return
     */
    public BigDecimal getArrayAverage(BigDecimal upper, BigDecimal lower, String left_top, String right_bottom, XSSFFormulaEvaluator evaluator,XSSFWorkbook xwb) {
        BigDecimal total = new BigDecimal(0);
        BigDecimal average = new BigDecimal(0);
        BigDecimal value = null;
        int count = 0;
        XSSFSheet sheet = xwb.getSheet("路基弯沉");
        XSSFRow row = null;
        int firstrow = Integer.valueOf(left_top.replaceAll("[^0-9]", "")) - 1;
        int lastrow = Integer.valueOf(right_bottom.replaceAll("[^0-9]", "")) - 1;
        for(int i = firstrow; i <= lastrow; i ++){
            row = sheet.getRow(i);
            if(row == null){
                continue;
            }
            for (int j = 18; j <= 21; j++) {
                if(row.getCell(j) == null || "".equals(row.getCell(j).toString())){
                    continue;
                }
                value = new BigDecimal(evaluator.evaluate(row.getCell(j)).getNumberValue());
                if(upper.compareTo(value) > 0 && lower.compareTo(value) < 0){
                    total = total.add(value);
                    count ++;
                }

            }
        }
        if(count <= 1){
            return average;
        }
        average = total.divide(new BigDecimal(count), 1, BigDecimal.ROUND_HALF_UP);
        System.out.println("average = " + average);
        return average;
    }

    /**
     * 将“路基弯沉”sheet中的临时数据计算出来 这些数据有： 舍前均方差,舍前平均值,舍前代表值,特异值下限,特异值上限,
     * 舍后均方,舍后平均值,舍后代表值,总测点数,特异点数
     * @param evaluator
     * @param sheet
     * @param start_row
     * @param end_row
     * @param left_top
     * @param right_bottom
     */
    public void fillTempDate(XSSFFormulaEvaluator evaluator, XSSFSheet sheet, int start_row, int end_row, String left_top, String right_bottom,XSSFWorkbook xwb) {
        int tables = (end_row - start_row)/35;

        XSSFRow row2, row3, row4, row5, row6, row7, row8, row9, row10, row11;
        row2 = sheet.getRow(start_row + 1);
        row3 = sheet.getRow(start_row + 2);
        row4 = sheet.getRow(start_row + 3);
        row5 = sheet.getRow(start_row + 4);
        row6 = sheet.getRow(start_row + 5);
        row7 = sheet.getRow(start_row + 6);
        row8 = sheet.getRow(start_row + 7);
        row9 = sheet.getRow(start_row + 8);
        row10 = sheet.getRow(start_row + 9);
        row11 = sheet.getRow(start_row + 10);
        sheet.getRow(start_row + 10);
        // 舍前均方差
        row2.getCell(16).setCellFormula(
                "IF(ISERROR(STDEV(" + left_top + ":" + right_bottom
                        + ")),\"\",STDEV(" + left_top + ":" + right_bottom
                        + "))");// =IF(ISERROR(STDEV(S10:V34)),"",STDEV(S10:V34))
        BigDecimal row2value = new BigDecimal(evaluator.evaluate(row2.getCell(16)).getNumberValue());
        row2.getCell(16).setCellFormula(null);
        row2.getCell(16).setCellValue(row2value.setScale(1, BigDecimal.ROUND_HALF_UP).doubleValue());
        // 舍前平均值
        row3.getCell(16).setCellFormula(
                "IF(ISERROR(AVERAGE(" + left_top + ":" + right_bottom
                        + ")),\"\",AVERAGE(" + left_top + ":" + right_bottom
                        + "))");// =IF(ISERROR(AVERAGE(S10:V34)),"",AVERAGE(S10:V34))
        BigDecimal row3value = new BigDecimal(evaluator.evaluate(row3.getCell(16)).getNumberValue());
        row3.getCell(16).setCellFormula(null);
        row3.getCell(16).setCellValue(row3value.setScale(1, BigDecimal.ROUND_HALF_UP).doubleValue());
        // 舍前代表值  =IF(ISERROR((Q3+F6*Q2)*F5*H6*H5),"",(Q3+F6*Q2)*F5*H6*H5)
        row4.getCell(16).setCellFormula(
                "IF(ISERROR((" + row3.getCell(16).getReference()+"+"//舍前平均值
                        +row5.getCell(2).getReference()+ "*"//保证率系数λ
                        +row2.getCell(16).getReference()+ ")*"//舍前均方差
                        +row5.getCell(5).getReference()+ "*"//温度影响系数
                        //+row6.getCell(7).getReference()+ "*"//湿度影响系数
                        +row5.getCell(7).getReference()+"),\"\",("//季节影响系数
                        +row3.getCell(16).getReference()+"+"//舍前平均值
                        +row5.getCell(2).getReference()+ "*"//保证率系数λ
                        +row2.getCell(16).getReference()+ ")*"//舍前均方差
                        +row5.getCell(5).getReference()+ "*"//温度影响系数
                        //+row6.getCell(7).getReference()+"*"//湿度影响系数
                        +row5.getCell(7).getReference()+")");//季节影响系数
        BigDecimal row4value = new BigDecimal(evaluator.evaluate(row4.getCell(16)).getNumberValue());
        row4.getCell(16).setCellFormula(null);
        row4.getCell(16).setCellValue(row4value.setScale(1, BigDecimal.ROUND_HALF_UP).doubleValue());
        // 特异值下限
        row5.getCell(16).setCellFormula(
                "IF(ISERROR(" + row3.getCell(16).getReference() + "-2*"
                        + row2.getCell(16).getReference() + "),\"\","
                        + row3.getCell(16).getReference() + "-2*"
                        + row2.getCell(16).getReference() + ")");// =IF(ISERROR(Q3-2*Q2),"",Q3-2*Q2)
        BigDecimal row5value = new BigDecimal(evaluator.evaluate(row5.getCell(16)).getNumberValue());
        row5.getCell(16).setCellFormula(null);
        row5.getCell(16).setCellValue(row5value.setScale(1, BigDecimal.ROUND_HALF_UP).doubleValue());
        // 特异值上限
        row6.getCell(16).setCellFormula(
                "IF(ISERROR(" + row3.getCell(16).getReference() + "+2*"
                        + row2.getCell(16).getReference() + "),\"\","
                        + row3.getCell(16).getReference() + "+2*"
                        + row2.getCell(16).getReference() + ")");// =IF(ISERROR(Q3+2*Q2),"",Q3+2*Q2)
        BigDecimal row6value = new BigDecimal(evaluator.evaluate(row6.getCell(16)).getNumberValue());
        row6.getCell(16).setCellFormula(null);
        row6.getCell(16).setCellValue(row6value.setScale(1, BigDecimal.ROUND_HALF_UP).doubleValue());
        row5value = new BigDecimal(evaluator.evaluate(row5.getCell(16)).getNumberValue()).setScale(1, BigDecimal.ROUND_HALF_UP);
        row6value = new BigDecimal(evaluator.evaluate(row6.getCell(16)).getNumberValue()).setScale(1, BigDecimal.ROUND_HALF_UP);
        //用函数计算值，然后填入单元格
        BigDecimal average = getArrayAverage(row6value, row5value, left_top, right_bottom, evaluator,xwb);
        row8.getCell(16).setCellValue(average.doubleValue());
        row7.getCell(16).setCellValue(getArraySTDEVA(row6value, row5value, average, left_top, right_bottom, evaluator,xwb).doubleValue());
        // 舍后代表值
        row9.getCell(16).setCellFormula(
                row8.getCell(16).getReference() + "+"
                        + row5.getCell(2).getReference() + "*"
                        + row7.getCell(16).getReference());// =Q8+F6*Q7
        // 总测点数
        row10.getCell(16).setCellFormula(
                "COUNT(" + left_top + ":" + right_bottom + ")");// =COUNT(S10:V34)
        // 特异点数  row5value特异值下限  row6value特异值上限
        //=COUNT(S10:V34)-SUM(IF(S10:V34<Q6*(S10:V34>Q5),1))

        //sheet.setArrayFormula("COUNT(" + left_top + ":" + right_bottom + ")-SUM(IF(" + left_top + ":" + right_bottom+"<"+row6value+"*("+left_top+":"+right_bottom+">"+row5value+"),1))",new CellRangeAddress(start_row + 10,start_row + 10,16,16));

        //row11.getCell(16).setCellFormula("COUNT("+left_top+":"+right_bottom+")-SUM(IF("+left_top+":"+right_bottom+"<"+row6value+"*("+left_top+":"+right_bottom+">"+row5value+"),1))");

        int tycount = getArrayCount(row6value, row5value, left_top, right_bottom, evaluator,xwb);
        //row11.getCell(16).setCellValue(tycount);
        row11.getCell(16).setCellFormula("COUNT(" + left_top + ":" + right_bottom + ")"+"-"+tycount);

        fillSheetBody(4,xwb, row5.getCell(11), row3.getCell(16), row4.getCell(11),
                row8.getCell(16));
        fillSheetBody(5,xwb, row6.getCell(11), row4.getCell(16), row4.getCell(11),
                row9.getCell(16));
        fillSheetBody(6, xwb,row7.getCell(11), row6.getCell(11), row4.getCell(11));
        XSSFRow t_row1, t_row2, t_row3;
        for (int i = 1; i < tables; i++) {
            t_row1 = sheet.getRow(start_row + i*35 + 4);
            t_row2 = sheet.getRow(start_row + i*35 + 5);
            t_row3 = sheet.getRow(start_row + i*35 + 6);
            fillSheetBody(4, xwb,t_row1.getCell(11), row3.getCell(16), row4.getCell(11),
                    row8.getCell(16));
            fillSheetBody(5, xwb,t_row2.getCell(11), row4.getCell(16), row4.getCell(11),
                    row9.getCell(16));
            fillSheetBody(6,xwb, t_row3.getCell(11), row6.getCell(11), row4.getCell(11));
        }
    }

    /**
     * 将“路基弯沉”sheet中的数据计算好，取出来放到右边,以便于计算
     * @param sheet
     */
    public void calculateTempDate(XSSFSheet sheet,XSSFWorkbook xwb) {
        XSSFFormulaEvaluator evaluator = new XSSFFormulaEvaluator(xwb);
        XSSFRow row;
        boolean flag = false;
        int start_row = 0;
        String left_top = "", right_bottom = "";
        String name = "";
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if ("桩号/部位".equals(row.getCell(0).toString()) ) {//&& !name.equals(row.getCell(2).toString())
                if("".equals(name)){
                    name = row.getCell(2).toString();
                    left_top = sheet.getRow(i + 6).getCell(18).getReference();
                    continue;
                }
                else{
                    try{
                        right_bottom = sheet.getRow(i - 5).getCell(21).getReference();
                    }
                    catch(Exception e){
                        right_bottom = sheet.getRow(i - 5).createCell(21).getReference();
                    }
                    fillTempDate(evaluator, sheet, start_row, i - 3, left_top, right_bottom,xwb);
                    start_row = i - 3;
                    name = row.getCell(2).toString();
                    left_top = sheet.getRow(i + 6).getCell(18).getReference();
                }
            }
            if ("桩号".equals(row.getCell(0).toString())) {
                flag = true;
                i++;
                continue;
            }
            if (row.getCell(0) != null && row.getCell(0).toString().startsWith("备注")) {
                flag = false;
                continue;
            }
            // 计算左边的值
            if (flag && row.getCell(2).toString() != null && !"".equals(row.getCell(2).toString()))
            {
                row.getCell(4).setCellFormula(
                        "IF(" + row.getCell(2).getReference() + "<>\"\","
                                + row.getCell(2).getReference() + "*2,\"\")");//=IF(Q3<L4,Q8,Q3)
                row.getCell(5).setCellFormula(
                        "IF(" + row.getCell(3).getReference() + "<>\"\","
                                + row.getCell(3).getReference() + "*2,\"\")");// =IF(C22<>"",C22*2,"")
                row.createCell(18).setCellFormula(row.getCell(4).getReference());// =E22
                row.createCell(19).setCellFormula(row.getCell(5).getReference());
            }
            // 计算右边的值
            if (flag && row.getCell(9).toString() != null && !"".equals(row.getCell(10).toString()))
            {
                row.getCell(11).setCellFormula(
                        "IF(" + row.getCell(9).getReference() + "<>\"\","
                                + row.getCell(9).getReference() + "*2,\"\")");
                row.getCell(12).setCellFormula(
                        "IF(" + row.getCell(10).getReference() + "<>\"\","
                                + row.getCell(10).getReference() + "*2,\"\")");
                row.createCell(20).setCellFormula(row.getCell(11).getReference());
                row.createCell(21).setCellFormula(row.getCell(12).getReference());
            }
        }
        try{
            right_bottom = sheet.getRow(sheet.getPhysicalNumberOfRows() - 2).getCell(21).getReference();
        }
        catch(Exception e){
            right_bottom = sheet.getRow(sheet.getPhysicalNumberOfRows() - 2).createCell(21).getReference();
        }
        fillTempDate(evaluator, sheet, start_row, sheet.getPhysicalNumberOfRows(), left_top, right_bottom,xwb);
    }

    /**
     * 将数据写入sheet为路基弯沉中
     * @param data
     * @return
     * @throws IOException
     */
    public boolean DBtoExcel(List<JjgFbgcLjgcLjwc> data,XSSFWorkbook wb) throws IOException {
        DecimalFormat nf = new DecimalFormat(".00");
        XSSFCellStyle cellstyle = wb.createCellStyle();
        XSSFFont font=wb.createFont();
        font.setFontHeightInPoints((short)9);
        font.setFontName("宋体");
        cellstyle.setFont(font);
        cellstyle.setBorderBottom(BorderStyle.THIN); //下边框
        cellstyle.setBorderLeft(BorderStyle.THIN);//左边框
        cellstyle.setBorderTop(BorderStyle.THIN);//上边框
        cellstyle.setBorderRight(BorderStyle.THIN);//右边框
        cellstyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直
        cellstyle.setAlignment(HorizontalAlignment.CENTER);//水平


        XSSFSheet sheet = wb.getSheet("路基弯沉");

        String zh = data.get(0).getZh();
        String xuhao = data.get(0).getXh();
        HashMap<Integer, String> map = new HashMap<>();
        String zhuanghao = zh+";\n";
        for(JjgFbgcLjgcLjwc row:data){
            if(xuhao.equals(row.getXh())){
                if(!zhuanghao.contains(row.getZh())){
                    zhuanghao += " "+row.getZh()+";\n";
                }
            }
            else{
                if(map.get(Integer.valueOf(xuhao)) == null){
                    map.put(Integer.valueOf(xuhao), zhuanghao.substring(0,zhuanghao.length()-1));
                }
                else{
                    map.put(Integer.valueOf(xuhao), map.get(Integer.valueOf(xuhao))+"\n"+zhuanghao.substring(0,zhuanghao.length()-1));
                }
                xuhao = row.getXh();
                zhuanghao = row.getZh()+";\n";
            }
        }
        if(!"".equals(zhuanghao)){
            if(map.get(Integer.valueOf(xuhao)) == null){
                map.put(Integer.valueOf(xuhao), zhuanghao.substring(0,zhuanghao.length()-1));
            }
            else{
                map.put(Integer.valueOf(xuhao), map.get(Integer.valueOf(xuhao))+"\n"+zhuanghao.substring(0,zhuanghao.length()-1));
            }
        }
        int index = 0;
        int tableNum = 0;
        zh = data.get(0).getZh();
        xuhao = data.get(0).getXh();
        fillTitleCellData(sheet, tableNum, data.get(0), map.get(Integer.valueOf(xuhao)));
        for(JjgFbgcLjgcLjwc row:data){
            if(xuhao.equals(row.getXh())){// 序号相同 在同一表格
                if(index == 50){
                    tableNum ++;
                    fillTitleCellData(sheet, tableNum, data.get(0), map.get(Integer.valueOf(xuhao)));
                    index = 0;
                }
                fillCommonCellData(sheet, tableNum, index, row, cellstyle);
                index ++;
            }
            else{
                zh = row.getZh();
                xuhao = row.getXh();
                tableNum ++;
                index = 0;
                fillTitleCellData(sheet, tableNum, row,map.get(Integer.valueOf(xuhao)));
                fillCommonCellData(sheet, tableNum, index, row, cellstyle);
                index += 1;
            }
        }
        return true;
    }

    /**
     * 填写路基弯沉工作簿中的数据
     * @param sheet
     * @param tableNum
     * @param index
     * @param row
     * @param cellstyle
     */
    public void fillCommonCellData(XSSFSheet sheet, int tableNum, int index, JjgFbgcLjgcLjwc row,XSSFCellStyle cellstyle) {
        sheet.getRow(tableNum*35+9+index%25).getCell(1+7*(index/25)).setCellValue(row.getCd());
        sheet.getRow(tableNum*35+9+index%25).getCell(2+7*(index/25)).setCellValue(Double.valueOf(row.getZz()).intValue());
        sheet.getRow(tableNum*35+9+index%25).getCell(3+7*(index/25)).setCellValue(Double.valueOf(row.getYz()).intValue());
        sheet.getRow(tableNum*35+9+index%25).getCell(6+7*(index/25)).setCellValue(Double.valueOf(row.getLbwd()).intValue());
        sheet.getRow(tableNum*35+9+index%25).getCell(7*(index/25)).setCellValue(row.getCjzh());
        sheet.getRow(tableNum*35+9+index%25).getCell(1+7*(index/25)).setCellStyle(cellstyle);


    }

    /**
     * 根据tableNum创建table
     * @param tableNum
     * @throws IOException
     */
    public void createTable(int tableNum,XSSFWorkbook wb) throws IOException {
        for(int i = 1; i < tableNum; i++){
            RowCopy.copyRows(wb, "路基弯沉", "路基弯沉", 0, 34, i*35);
        }
        if(tableNum > 1){
            wb.setPrintArea(wb.getSheetIndex("路基弯沉"), 0, 13, 0, tableNum*35-1);
        }
    }

    /**
     * 填写“路基弯沉”sheet的表头
     * @param sheet
     * @param tableNum
     * @param position
     */
    public void fillTitleCellData(XSSFSheet sheet, int tableNum,JjgFbgcLjgcLjwc row, String position) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        sheet.getRow(tableNum * 35 + 1).getCell(2).setCellValue(row.getProname());
        sheet.getRow(tableNum * 35 + 1).getCell(11).setCellValue(row.getHtd());
        sheet.getRow(tableNum*35+2).getCell(2).setCellValue("路基土石方");
        sheet.getRow(tableNum*35+2).getCell(11).setCellValue(simpleDateFormat.format(row.getJcsj()));//检测日期
        sheet.getRow(tableNum*35+3).getCell(2).setCellValue(position);//工程部位

        sheet.getRow(tableNum*35+3).getCell(11).setCellType(Cell.CELL_TYPE_NUMERIC);
        sheet.getRow(tableNum*35+3).getCell(11).setCellValue(Double.parseDouble(row.getYswcz()));//(nf2.format(Double.parseDouble(row[7])));//验收弯沉值

        sheet.getRow(tableNum*35+4).getCell(2).setCellValue(Double.valueOf(row.getMbkkzb()));//保证率系数λ  目标可靠指标

        sheet.getRow(tableNum*35+4).getCell(5).setCellValue(Double.valueOf(row.getSdyxxs()));//温度影响系数

        if(row.getSdyxxs() == null || "".equals(row.getSdyxxs())){
            sheet.getRow(tableNum*35+4).getCell(7).setCellValue(1);//季节影响系数
        }
        else{
            sheet.getRow(tableNum*35+4).getCell(7).setCellValue(Double.valueOf(row.getJjyxxs()));//季节影响系数
        }
        sheet.getRow(tableNum*35+5).getCell(2).setCellValue(row.getJgcc());//结构层次
        sheet.getRow(tableNum*35+5).getCell(6).setCellValue(row.getJglx());//结构类型

        sheet.getRow(tableNum*35+6).getCell(2).setCellValue(Double.valueOf(row.getHzz()));//后轴重
        sheet.getRow(tableNum*35+6).getCell(6).setCellValue(Double.valueOf(row.getLtqy()));//轮胎气压
        if (row.getBz()==null || "".equals(row.getBz())){
            sheet.getRow(tableNum*35+34).getCell(0).setCellValue("备注：");//备注

        }else {
            sheet.getRow(tableNum*35+34).getCell(0).setCellValue("备注："+row.getBz());//备注
        }

    }


    /**
     * 获取要生成页数的数量
     * @param proname
     * @param htd
     * @param fbgc
     * @return
     */
    public int gettableNum(String proname,String htd,String fbgc){
        QueryWrapper<JjgFbgcLjgcLjwc> queryWrapper = new QueryWrapper();
        queryWrapper.select("DISTINCT xh").like("proname",proname).like("htd",htd).like("fbgc",fbgc);
        List<JjgFbgcLjgcLjwc> xhnum = jjgFbgcLjgcLjwcMapper.selectList(queryWrapper);
        List<Integer> xuhaoList = new ArrayList<>();
        for (JjgFbgcLjgcLjwc x:xhnum){
            xuhaoList.add(Integer.valueOf(x.getXh()));

        }
        int tableNum = 0;
        for(int i = 0; i < xuhaoList.size(); i++){
            QueryWrapper<JjgFbgcLjgcLjwc> queryWrapperxh = new QueryWrapper();
            queryWrapperxh.like("proname",proname).like("htd",htd).like("fbgc",fbgc).like("xh", xuhaoList.get(i));
            List<JjgFbgcLjgcLjwc> jjgFbgcLjgcLjwcs = jjgFbgcLjgcLjwcMapper.selectList(queryWrapperxh);
            int dataNum = jjgFbgcLjgcLjwcs.size();
            tableNum+= dataNum%50 == 0 ? dataNum/50 : dataNum/50+1;
        }
        return tableNum;
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        String title = "弯沉质量鉴定结果汇总表";
        String sheetname = "评定单元";
        //获取鉴定表文件
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "02路基弯沉(贝克曼梁法).xlsx");
        if (!f.exists()) {
            return null;
        } else {
            XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));
            //读取工作表
            XSSFSheet slSheet = xwb.getSheet(sheetname);
            XSSFCell bt = slSheet.getRow(0).getCell(0);//标题
            XSSFCell xmname = slSheet.getRow(1).getCell(2);//项目名
            XSSFCell htdname = slSheet.getRow(1).getCell(7);//合同段名
            List<Map<String, Object>> mapList = new ArrayList<>();
            Map<String, Object> jgmap = new HashMap<>();
            DecimalFormat df = new DecimalFormat("0.00");
            DecimalFormat decf = new DecimalFormat("0.##");
            if (proname.equals(xmname.toString()) && title.equals(bt.toString()) && htd.equals(htdname.toString())) {
                int lastRowNum = slSheet.getLastRowNum();
                slSheet.getRow(lastRowNum).getCell(4).setCellType(CellType.STRING);//检测单元数
                slSheet.getRow(lastRowNum).getCell(6).setCellType(CellType.STRING);//合格单元
                slSheet.getRow(lastRowNum).getCell(8).setCellType(CellType.STRING);//合格率
                slSheet.getRow(5).getCell(5).setCellType(CellType.STRING);//合格率
                jgmap.put("检测单元数", decf.format(Double.valueOf(slSheet.getRow(lastRowNum).getCell(4).getStringCellValue())));
                jgmap.put("合格单元数", decf.format(Double.valueOf(slSheet.getRow(lastRowNum).getCell(6).getStringCellValue())));
                jgmap.put("合格率", df.format(Double.valueOf(slSheet.getRow(lastRowNum).getCell(8).getStringCellValue())));
                jgmap.put("规定值", slSheet.getRow(5).getCell(5).getStringCellValue());
                mapList.add(jgmap);
                return mapList;
            }
            return null;
        }
    }

    @Override
    public void exportljwc(HttpServletResponse response) {
        String fileName = "02路基弯沉(贝克曼梁法)实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcLjgcLjwcVo()).finish();

    }

    @Override
    public void importljwc(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcLjgcLjwcVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcLjgcLjwcVo>(JjgFbgcLjgcLjwcVo.class) {
                                @Override
                                public void handle(List<JjgFbgcLjgcLjwcVo> dataList) {
                                    for(JjgFbgcLjgcLjwcVo ljwcVo: dataList)
                                    {
                                        JjgFbgcLjgcLjwc ljwc = new JjgFbgcLjgcLjwc();
                                        BeanUtils.copyProperties(ljwcVo,ljwc);
                                        ljwc.setCreatetime(new Date());
                                        ljwc.setProname(commonInfoVo.getProname());
                                        ljwc.setHtd(commonInfoVo.getHtd());
                                        ljwc.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcLjgcLjwcMapper.insert(ljwc);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }
}
