package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLjgcLjwc;
import glgc.jjgys.model.project.JjgFbgcLjgcLjwcLcf;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.ljgc.JjgFbgcLjgcLjwcLcfVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcLjgcLjwcLcfMapper;
import glgc.jjgys.system.service.JjgFbgcLjgcLjwcLcfService;
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
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-03-01
 */
@Service
public class JjgFbgcLjgcLjwcLcfServiceImpl extends ServiceImpl<JjgFbgcLjgcLjwcLcfMapper, JjgFbgcLjgcLjwcLcf> implements JjgFbgcLjgcLjwcLcfService {

    @Autowired
    private JjgFbgcLjgcLjwcLcfMapper jjgFbgcLjgcLjwcLcfMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;


    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException {
        XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获取数据
        QueryWrapper<JjgFbgcLjgcLjwcLcf> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("xh","cjzh");
        List<JjgFbgcLjgcLjwcLcf> data = jjgFbgcLjgcLjwcLcfMapper.selectList(wrapper);
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"02路基弯沉(落锤法).xlsx");
        if (data == null || data.size()==0){
            return;
        }else {
            File fdir = new File(filepath + File.separator + proname + File.separator + htd);
            if (!fdir.exists()) {
                //创建文件根目录
                fdir.mkdirs();
            }
            File directory = new File("service-system/src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String path = reportPath + File.separator + "路基弯沉(落锤法).xlsx";
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            int num = gettableNum(proname, htd, fbgc);
            createTable(num,wb);
            if(DBtoExcel(data,wb)){
                //右侧的数据
                calculateTempDate(wb.getSheet("路基弯沉"),wb);
                ArrayList<String> totalref = getTotalMark(wb.getSheet("路基弯沉"));
                String time = getLastTime(wb.getSheet("路基弯沉"));
                //创建“评定单元”sheet
                createEvaluateTable(totalref,wb);
                //引用数据，然后计算
                completeTotleTable(wb.getSheet("评定单元"), totalref, time,wb,proname,htd);

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
        if (cell.length == 2) {
            cell1 = cell[0];
            cell2 = cell[1];
            cell1.setCellFormula(cell2.getReference());

        }
        //=IF(J6>J4,"×","√")
        if (cell.length == 3) {
            cell1 = cell[0];//J7
            cell2 = cell[1];//J6
            cell3 = cell[2];//J4
            cell1.setCellFormula("IF(" + cell2.getReference() + ">"
                    + cell3.getReference() + ",\"×\",\"√\")");
            XSSFFormulaEvaluator evaluate = new XSSFFormulaEvaluator(xwb);
            if("×".equals(cell1.getRawValue())){
                cell1.setCellStyle(cellstyle);
            }
        }
    }

    public void fillTempDate(XSSFFormulaEvaluator evaluator, XSSFSheet sheet, int start_row, int end_row, String left_top, String right_bottom,XSSFWorkbook xwb) {
        int tables = (end_row - start_row)/34;
        XSSFRow row2, row3, row4, row5, row6, row7;
        row2 = sheet.getRow(start_row + 1);
        row3 = sheet.getRow(start_row + 2);
        row4 = sheet.getRow(start_row + 3);
        row5 = sheet.getRow(start_row + 4);
        row6 = sheet.getRow(start_row + 5);
        row7 = sheet.getRow(start_row + 6);
        sheet.getRow(start_row + 10);
        // 均方差
        row2.getCell(13).setCellFormula(
                "IF(ISERROR(STDEV(" + left_top + ":" + right_bottom
                        + ")),\"\",STDEV(" + left_top + ":" + right_bottom
                        + "))");//  =IF(ISERROR(STDEV(P9:Q33)),"",STDEV(P9:Q33))
        BigDecimal row2value = new BigDecimal(evaluator.evaluate(row2.getCell(13)).getNumberValue());
        //row2.getCell(13).setCellFormula(null);
        row2.getCell(13).setCellValue(row2value.setScale(1, BigDecimal.ROUND_HALF_UP).doubleValue());
        // 平均值
        row3.getCell(13).setCellFormula(
                "IF(ISERROR(AVERAGE(" + left_top + ":" + right_bottom
                        + ")),\"\",AVERAGE(" + left_top + ":" + right_bottom
                        + "))");//  =IF(ISERROR(AVERAGE(P9:Q33)),"",AVERAGE(P9:Q33))
        BigDecimal row3value = new BigDecimal(evaluator.evaluate(row3.getCell(13)).getNumberValue());
        //row3.getCell(13).setCellFormula(null);
        row3.getCell(13).setCellValue(row3value.setScale(1, BigDecimal.ROUND_HALF_UP).doubleValue());
        // 舍前代表值  =IF(ISERROR((Q3+F6*Q2)*F5*H6*H5),"",(Q3+F6*Q2)*F5*H6*H5)   =IF(ISERROR((N3+D6*N2)*D5*F5*F6),"",(N3+D6*N2)*D5*F5*F6)
        row4.getCell(13).setCellFormula(
                "IF(ISERROR((" + row3.getCell(13).getReference()+"+"//平均值
                        +row6.getCell(3).getReference()+ "*"//目标可靠值
                        +row2.getCell(13).getReference()+ ")*"//均方差
                        +row5.getCell(3).getReference()+ "*"//温度影响系数
                        +row5.getCell(5).getReference()+ "*"//湿度影响系数
                        +row6.getCell(5).getReference()+"),\"\",("//季节影响系数
                        +row3.getCell(13).getReference()+"+"//平均值
                        +row6.getCell(3).getReference()+ "*"//保证率系数λ
                        +row2.getCell(13).getReference()+ ")*"//舍前均方差
                        +row5.getCell(3).getReference()+ "*"//温度影响系数
                        +row5.getCell(5).getReference()+"*"//湿度影响系数
                        +row6.getCell(5).getReference()+")");//季节影响系数
        BigDecimal row4value = new BigDecimal(evaluator.evaluate(row4.getCell(13)).getNumberValue());
        //row4.getCell(13).setCellFormula(null);
        row4.getCell(13).setCellValue(row4value.setScale(1, BigDecimal.ROUND_HALF_UP).doubleValue());
        // 特异值下限   =IF(ISERROR(N3-2*N2),"",N3-2*N2)
        row5.getCell(13).setCellFormula(
                "IF(ISERROR(" + row3.getCell(13).getReference() + "-2*"  //N3
                        + row2.getCell(13).getReference() + "),\"\","
                        + row3.getCell(13).getReference() + "-2*"
                        + row2.getCell(13).getReference() + ")");// =IF(ISERROR(Q3-2*Q2),"",Q3-2*Q2)
        BigDecimal row5value = new BigDecimal(evaluator.evaluate(row5.getCell(13)).getNumberValue());
        //row5.getCell(13).setCellFormula(null);
        row5.getCell(13).setCellValue(row5value.setScale(1, BigDecimal.ROUND_HALF_UP).doubleValue());
        // 特异值上限  =IF(ISERROR(N3+2*N2),"",N3+2*N2)
        row6.getCell(13).setCellFormula(
                "IF(ISERROR(" + row3.getCell(13).getReference() + "+2*"
                        + row2.getCell(13).getReference() + "),\"\","
                        + row3.getCell(13).getReference() + "+2*"
                        + row2.getCell(13).getReference() + ")");// =IF(ISERROR(Q3+2*Q2),"",Q3+2*Q2)
        BigDecimal row6value = new BigDecimal(evaluator.evaluate(row6.getCell(13)).getNumberValue());
        //row6.getCell(13).setCellFormula(null);
        row6.getCell(13).setCellValue(row6value.setScale(1, BigDecimal.ROUND_HALF_UP).doubleValue());

        // 总测点数  =COUNT(P9:Q33)
        row7.getCell(13).setCellFormula(
                "COUNT(" + left_top + ":" + right_bottom + ")");// =COUNT(S10:V34)

        fillSheetBody(4,xwb, row5.getCell(9), row3.getCell(13));
        fillSheetBody(5,xwb, row6.getCell(9), row4.getCell(13));
        fillSheetBody(6, xwb,row7.getCell(9), row6.getCell(9), row4.getCell(9));
        XSSFRow t_row1, t_row2, t_row3;
        for (int i = 1; i < tables; i++) {
            t_row1 = sheet.getRow(start_row + i*34 + 4);
            t_row2 = sheet.getRow(start_row + i*34 + 5);
            t_row3 = sheet.getRow(start_row + i*34 + 6);
            fillSheetBody(4, xwb,t_row1.getCell(9), row3.getCell(13));
            fillSheetBody(5, xwb,t_row2.getCell(9), row4.getCell(13));
            fillSheetBody(6, xwb,t_row3.getCell(9), row6.getCell(9), row4.getCell(9));
        }
    }

    /**
     *将“路基弯沉”sheet中的数据，取出来放到右边
     * @param sheet
     * @param xwb
     */
    public void calculateTempDate(XSSFSheet sheet,XSSFWorkbook xwb) {
        XSSFFormulaEvaluator evaluator = new XSSFFormulaEvaluator(xwb);
        XSSFRow row;
        boolean flag = false;
        int start_row = 0;
        String left_top = "", right_bottom = "";
        String name = "";
        DecimalFormat nf = new DecimalFormat();
        System.out.println(sheet.getFirstRowNum());
        System.out.println(sheet.getPhysicalNumberOfRows());
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if ("桩号/部位".equals(row.getCell(0).toString()) ) {//&& !name.equals(row.getCell(2).toString())
                if("".equals(name)){
                    name = row.getCell(1).toString();//桩号
                    left_top = sheet.getRow(i + 5).getCell(15).getReference();//p9
                    continue;
                }
                else{
                    try{
                        right_bottom = sheet.getRow(i - 5).getCell(16).getReference();
                    }
                    catch(Exception e){
                        right_bottom = sheet.getRow(i - 5).createCell(16).getReference();
                    }
                    fillTempDate(evaluator, sheet, start_row, i - 3, left_top, right_bottom,xwb);
                    start_row = i - 3;
                    name = row.getCell(1).toString();
                    left_top = sheet.getRow(i + 5).getCell(15).getReference();
                }
            }
            if ("桩号".equals(row.getCell(0).toString())) {
                flag = true;
                //i++;
                continue;
            }
            if (row.getCell(0) != null && row.getCell(0).toString().startsWith("备注")) {
                flag = false;
                continue;
            }
            if (flag && row.getCell(3).toString() != null && !"".equals(row.getCell(3).toString())) {
                row.createCell(15);
                row.getCell(15).setCellFormula("IF("+row.getCell(3).getReference()+"=\"\",\"\","+row.getCell(3).getReference()+")");
                BigDecimal rowvalue = new BigDecimal(evaluator.evaluate(row.getCell(15)).getNumberValue());
                row.getCell(15).setCellValue(rowvalue.setScale(0, BigDecimal.ROUND_HALF_UP).doubleValue());

                //=IF(D9="","",D9)
                //row.createCell(15).setCellFormula("IF("+row.getCell(3).getReference()+"=\"\",\"\","+row.getCell(3).getReference()+")");



            }
            if (flag && row.getCell(9).toString() != null && !"".equals(row.getCell(9).toString())) {
                row.createCell(16);
                row.getCell(16).setCellFormula("IF("+row.getCell(9).getReference()+"=\"\",\"\","+row.getCell(9).getReference()+")");
                BigDecimal rowvalue = new BigDecimal(evaluator.evaluate(row.getCell(16)).getNumberValue());
                row.getCell(16).setCellValue(rowvalue.setScale(0, BigDecimal.ROUND_HALF_UP).doubleValue());
                //=IF(D9="","",D9)
                //row.createCell(16).setCellFormula("IF("+row.getCell(9).getReference()+"=\"\",\"\","+row.getCell(9).getReference()+")");

            }


        }
        try{
            right_bottom = sheet.getRow(sheet.getPhysicalNumberOfRows() - 2).getCell(16).getReference();
        }
        catch(Exception e){
            right_bottom = sheet.getRow(sheet.getPhysicalNumberOfRows() - 2).createCell(16).getReference();
        }
        fillTempDate(evaluator, sheet, start_row, sheet.getPhysicalNumberOfRows(), left_top, right_bottom,xwb);
    }

    /**
     * 将“评定单元”sheet数据引用填好，然后计算
     * @param sheet
     * @param ref
     */
    public void completeTotleTable(XSSFSheet sheet, ArrayList<String> ref, String time,XSSFWorkbook xwb,String proname,String htd) {
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
        System.out.println(proname+htd+time);
        sheet.getRow(1).getCell(2).setCellValue(proname);
        sheet.getRow(1).getCell(7).setCellValue(htd);
        sheet.getRow(2).getCell(2).setCellValue("路基土石方");
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
                //sheet.addMergedRegion(new CellRangeAddress(i, i, 1, 4));
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
            if (row.getCell(0) != null && "分部工程名称".equals(row.getCell(0).toString()) && row.getCell(9) != null && !"".equals(row.getCell(9).toString())) {
                try {
                    Date dt1 = simpleDateFormat.parse(time);
                    Date dt2 = simpleDateFormat.parse(row.getCell(9).toString());
                    if(dt1.getTime() < dt2.getTime()){
                        time = row.getCell(9).toString();
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
            if ("桩号/部位".equals(row.getCell(0).toString()) && row.getCell(13) != null && !"".equals(row.getCell(13).toString())) {
                ref.add(row.getCell(1).getReference());
                // 验收弯沉值（0.01mm）
                ref.add(row.getCell(9).getReference());
            }
            // 弯沉代表值(0.01mm)
            if ("代表弯沉值(0.01mm)".equals(row.getCell(6).toString()) && row.getCell(13) != null && !"".equals(row.getCell(13).toString())) {
                ref.add(row.getCell(9).getReference());
            }
        }
        return ref;
    }

    /**
     * 填写“路基弯沉”sheet的表头
     * @param sheet
     * @param tableNum
     * @param position
     */
    public void fillTitleCellData(XSSFSheet sheet, int tableNum,JjgFbgcLjgcLjwcLcf row, String position) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        sheet.getRow(tableNum * 34 + 1).getCell(1).setCellValue(row.getProname());
        sheet.getRow(tableNum * 34 + 1).getCell(9).setCellValue(row.getHtd());
        sheet.getRow(tableNum*34+2).getCell(1).setCellValue("路基土石方");
        sheet.getRow(tableNum*34+2).getCell(9).setCellValue(simpleDateFormat.format(row.getJcsj()));//检测日期
        sheet.getRow(tableNum*34+3).getCell(1).setCellValue(position);//工程部位

        sheet.getRow(tableNum*34+3).getCell(9).setCellType(Cell.CELL_TYPE_NUMERIC);
        sheet.getRow(tableNum*34+3).getCell(9).setCellValue(Double.parseDouble(row.getSjwcz()));//(nf2.format(Double.parseDouble(row[7])));//验收弯沉值

        sheet.getRow(tableNum*34+4).getCell(1).setCellValue(row.getJgcc());//结构层次
        sheet.getRow(tableNum*34+4).getCell(3).setCellValue(Double.valueOf(row.getWdyxxs()));//温度影响系数
        sheet.getRow(tableNum*34+4).getCell(5).setCellValue(Double.valueOf(row.getJjyxxs()));//季节影响系数
        sheet.getRow(tableNum*34+5).getCell(1).setCellValue(row.getJglx());//结构类型
        sheet.getRow(tableNum*34+5).getCell(3).setCellValue(row.getMbkkzb());//目标可靠指标
        sheet.getRow(tableNum*34+5).getCell(5).setCellValue(Double.valueOf(row.getSdyxxs()));//湿度影响系数
        sheet.getRow(tableNum*34+6).getCell(1).setCellValue(Double.valueOf(row.getLcz()));//落锤重
        sheet.getRow(tableNum*34+6).getCell(3).setCellValue(row.getYqmc());//仪器名称


        if (row.getBz()==null || "".equals(row.getBz())){
            sheet.getRow(tableNum*34+33).getCell(0).setCellValue("备注：");//备注

        }else {
            sheet.getRow(tableNum*34+33).getCell(0).setCellValue("备注："+row.getBz());//备注
        }

    }

    /**
     * 将数据写入sheet为路基弯沉中
     * @param data
     * @return
     * @throws IOException
     */
    public boolean DBtoExcel(List<JjgFbgcLjgcLjwcLcf> data,XSSFWorkbook wb) throws IOException {
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
        for(JjgFbgcLjgcLjwcLcf row:data){
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
        int index = 0;//每条数据的下标
        int tableNum = 0;//页数
        xuhao = data.get(0).getXh();
        fillTitleCellData(sheet, tableNum, data.get(0),map.get(Integer.valueOf(xuhao)));//填写表头
        for(JjgFbgcLjgcLjwcLcf row:data){
            if (xuhao.equals(row.getXh())){
                if (index == 50){
                    tableNum++;
                    fillTitleCellData(sheet, tableNum, data.get(0),map.get(Integer.valueOf(xuhao)));//填写表头
                    index = 0;
                }
                fillCommonCellData(sheet, tableNum, index, row, cellstyle);
                index += 1;
            }else {
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
    public void fillCommonCellData(XSSFSheet sheet, int tableNum, int index, JjgFbgcLjgcLjwcLcf row,XSSFCellStyle cellstyle) {
        sheet.getRow(tableNum*34+8+index%25).getCell(2+6*(index/25)).setCellValue(row.getCd());
        sheet.getRow(tableNum*34+8+index%25).getCell(3+6*(index/25)).setCellValue(Double.valueOf(row.getScwcz()));
        sheet.getRow(tableNum*34+8+index%25).getCell(4+6*(index/25)).setCellValue(Double.valueOf(row.getLbwd()));
        sheet.getRow(tableNum*34+8+index%25).getCell(6*(index/25)).setCellValue(row.getCjzh());
        sheet.getRow(tableNum*34+8+index%25).getCell(1+6*(index/25)).setCellStyle(cellstyle);

    }

    /**
     * 根据tableNum创建table
     * @param tableNum
     * @throws IOException
     */
    public void createTable(int tableNum,XSSFWorkbook wb) throws IOException {
        for(int i = 1; i < tableNum; i++){
            RowCopy.copyRows(wb, "路基弯沉", "路基弯沉", 0, 33, i*34);
        }
        if(tableNum > 1){
            wb.setPrintArea(wb.getSheetIndex("路基弯沉"), 0, 10, 0, tableNum*34-1);
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
        QueryWrapper<JjgFbgcLjgcLjwcLcf> queryWrapper = new QueryWrapper();
        queryWrapper.select("DISTINCT xh").like("proname",proname).like("htd",htd).like("fbgc",fbgc);
        List<JjgFbgcLjgcLjwcLcf> xhnum = jjgFbgcLjgcLjwcLcfMapper.selectList(queryWrapper);
        List<Integer> xuhaoList = new ArrayList<>();
        for (JjgFbgcLjgcLjwcLcf x:xhnum){
            xuhaoList.add(Integer.valueOf(x.getXh()));
        }
        int tableNum = 0;
        for(int i = 0; i < xuhaoList.size(); i++){
            QueryWrapper<JjgFbgcLjgcLjwcLcf> queryWrapperxh = new QueryWrapper();
            queryWrapperxh.like("proname",proname).like("htd",htd).like("fbgc",fbgc).like("xh", xuhaoList.get(i));
            List<JjgFbgcLjgcLjwcLcf> jjgFbgcLjgcLjwcs = jjgFbgcLjgcLjwcLcfMapper.selectList(queryWrapperxh);
            int dataNum = jjgFbgcLjgcLjwcs.size();
            tableNum+= dataNum%50 == 0 ? dataNum/50 : dataNum/50+1;
        }
        return tableNum;
    }

    /**
     * 导出模板文件
     * @param response
     */
    @Override
    public void exportljwclcf(HttpServletResponse response) {
        String fileName = "路基弯沉（落锤法）实测数据";
        String sheetName = "路基弯沉（落锤法）实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcLjgcLjwcLcfVo()).finish();

    }

    /**
     * 导入数据
     * @param file
     * @param commonInfoVo
     */
    @Override
    public void importljwclcf(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcLjgcLjwcLcfVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcLjgcLjwcLcfVo>(JjgFbgcLjgcLjwcLcfVo.class) {
                                @Override
                                public void handle(List<JjgFbgcLjgcLjwcLcfVo> dataList) {
                                    for(JjgFbgcLjgcLjwcLcfVo ljwclcfVo: dataList)
                                    {
                                        JjgFbgcLjgcLjwcLcf ljwclcf = new JjgFbgcLjgcLjwcLcf();
                                        BeanUtils.copyProperties(ljwclcfVo,ljwclcf);
                                        ljwclcf.setCreatetime(new Date());
                                        ljwclcf.setProname(commonInfoVo.getProname());
                                        ljwclcf.setHtd(commonInfoVo.getHtd());
                                        ljwclcf.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcLjgcLjwcLcfMapper.insert(ljwclcf);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }
}
