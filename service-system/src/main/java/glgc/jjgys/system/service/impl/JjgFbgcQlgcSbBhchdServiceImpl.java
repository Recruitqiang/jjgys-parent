package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLjgcLjwcLcf;
import glgc.jjgys.model.project.JjgFbgcQlgcSbBhchd;
import glgc.jjgys.model.project.JjgFbgcQlgcSbJgcc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.qlgc.JjgFbgcQlgcSbBhchdVo;
import glgc.jjgys.model.projectvo.qlgc.JjgFbgcQlgcSbJgccVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcQlgcSbBhchdMapper;
import glgc.jjgys.system.service.JjgFbgcQlgcSbBhchdService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
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
 * @since 2023-03-20
 */
@Service
public class JjgFbgcQlgcSbBhchdServiceImpl extends ServiceImpl<JjgFbgcQlgcSbBhchdMapper, JjgFbgcQlgcSbBhchd> implements JjgFbgcQlgcSbBhchdService {

    @Autowired
    private JjgFbgcQlgcSbBhchdMapper jjgFbgcQlgcSbBhchdMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    /**
     * 生成鉴定表
     * @param commonInfoVo
     */
    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获得数据
        QueryWrapper<JjgFbgcQlgcSbBhchd> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("qlmc","gjbhjjcbw");
        List<JjgFbgcQlgcSbBhchd> data = jjgFbgcQlgcSbBhchdMapper.selectList(wrapper);

        List<Map<String,Object>> selectnum = jjgFbgcQlgcSbBhchdMapper.selectnum(proname, htd, fbgc);

        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"31桥梁上部保护层厚度.xlsx");
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
            String name = "桥梁保护层厚度鉴定表.xlsx";
            String path = reportPath + File.separator + name;
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream in = new FileInputStream(f);
            wb = new XSSFWorkbook(in);
            createTable(gettableNum(selectnum),wb);
            if(DBtoExcel(data,wb)){
                calculateSizeSheet(wb);
                for (int j = 0; j < wb.getNumberOfSheets(); j++) {   //表内公式  计算 显示结果
                    JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
                }
                FileOutputStream fileOut = new FileOutputStream(f);
                wb.write(fileOut);
                fileOut.flush();
                fileOut.close();

            }
            in.close();
            wb.close();
        }

    }

    /**
     * 对“保护层”sheet进行计算
     * @param xwb
     */
    private void calculateSizeSheet(XSSFWorkbook xwb) {
        XSSFSheet sheet = xwb.getSheet("保护层");
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
            XSSFFormulaEvaluator e = new XSSFFormulaEvaluator(xwb);
            double designValue = 0;//设计值
            double avgValue = 0;//平均值

            // 可以计算
            if (!"".equals(row.getCell(4).toString()) && flag) {

                int temp_avgValue = 0;//平均值

                row.getCell(10).setCellFormula("AVERAGE("+row.getCell(4).getReference()+":"
                        +row.getCell(5).getReference()+")-"+row.getCell(13).getReference());//平均值  减去修正值
                temp_avgValue = Double.valueOf(e.evaluate(row.getCell(10)).getNumberValue()+0.5).intValue();
                row.getCell(10).setCellFormula(null);
                row.getCell(10).setCellValue(temp_avgValue);

                row.getCell(11).setCellFormula("AVERAGE("+row.getCell(6).getReference()+":"
                        +row.getCell(7).getReference()+")-"+row.getCell(13).getReference());//平均值 平均值  减去修正值
                temp_avgValue = Double.valueOf(e.evaluate(row.getCell(11)).getNumberValue()+0.5).intValue();
                row.getCell(11).setCellFormula(null);
                row.getCell(11).setCellValue(temp_avgValue);

                row.getCell(12).setCellFormula("AVERAGE("+row.getCell(8).getReference()+":"
                        +row.getCell(9).getReference()+")-"+row.getCell(13).getReference());//平均值 平均值  减去修正值
                temp_avgValue = Double.valueOf(e.evaluate(row.getCell(12)).getNumberValue()+0.5).intValue();
                row.getCell(12).setCellFormula(null);
                row.getCell(12).setCellValue(temp_avgValue);


                row.getCell(14).setCellFormula("AVERAGE("+row.getCell(10).getReference()+":"
                        +row.getCell(12).getReference()+")");//平均值

                temp_avgValue = Double.valueOf(e.evaluate(row.getCell(14)).getNumberValue()+0.5).intValue();
                row.getCell(14).setCellFormula(null);
                row.getCell(14).setCellValue(temp_avgValue);

                designValue = Double.valueOf(row.getCell(3).toString());
                avgValue = Double.valueOf(e.evaluate(row.getCell(7+7)).getNumberValue()+0.5).intValue();


                if(Math.abs(designValue-avgValue)<= Double.valueOf(analysisOffset(row.getCell(8+7).toString())[1]).intValue()){
                    row.getCell(9+7).setCellValue("√");
                }
                else{
                    row.getCell(10+7).setCellValue("×");
                }
            }
            // 到下一个表了
            if (!"".equals(row.getCell(0).toString())
                    && row.getCell(0).toString().contains("质量鉴定表") && flag) {
                flag = false;
            }
            // 下一座桥
            if (!"".equals(row.getCell(0).toString()) && !name.equals(getBridgeName(row.getCell(0).toString())) && flag) {
                if (start.size() > 0) {
                    rowrecord = start.get(0);
                    setTotalData(xwb, rowrecord, start, end, 7+7, 9+7, 10+7, 11+7, 12+7, 13+7, 14+7);
                }
                start.clear();
                end.clear();
                name = getBridgeName(row.getCell(0).toString());
                rowstart = sheet.getRow(i);
                rowrecord = rowstart;

                rowend = sheet.getRow(getCellEndRow(sheet, i, 0));
                start.add(rowstart);
                end.add(rowend);
                bridgename.add(rowstart.getCell(0).toString());
            }// 同一座桥，但是在不同的单元格
            else if (!"".equals(row.getCell(0).toString()) && name.equals(getBridgeName(row.getCell(0).toString())) && flag && rowstart != row) {
                rowstart = sheet.getRow(i);
                rowend = sheet.getRow(getCellEndRow(sheet, i, 0));
                start.add(rowstart);
                end.add(rowend);
            }
            // 可以计算啦
            if ("桥梁名称".equals(row.getCell(0).toString())) {//找到 桥梁名称
                flag = true;
                i += 2; // 向下两行
                rowstart = sheet.getRow(i + 1);
                rowend = sheet.getRow(getCellEndRow(sheet, i + 1, 0));
                if (name.equals(getBridgeName(rowstart.getCell(0).toString()))) {
                    start.add(rowstart);
                    end.add(rowend);
                } else {
                    if (start.size() > 0) {
                        rowrecord = start.get(0);
                        setTotalData(xwb, rowrecord, start, end, 7+7, 9+7, 10+7, 11+7, 12+7, 13+7, 14+7);
                    }
                    start.clear();
                    end.clear();
                    start.add(rowstart);
                    end.add(rowend);
                    name = getBridgeName(rowstart.getCell(0).toString());
                    bridgename.add(name);
                }
            }
        }
        /*
         * sheet中的所有数据都读取，处理完毕，如果List中还有数据，说明还有桥的数据没有汇总 此处进行最后的计算
         */
        if (start.size() > 0) {
            rowrecord = start.get(0);
            setTotalData(xwb, rowrecord, start, end, 7+7, 9+7, 10+7, 11+7, 12+7, 13+7, 14+7);
        }
        start.clear();
        end.clear();
        row = sheet.getRow(2);
        row.createCell(18).setCellFormula("SUM("
                +sheet.getRow(3).createCell(18).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(18).getReference()+")/2");//=SUM(L7:L300)
        row.createCell(19).setCellFormula("SUM("
                +sheet.getRow(3).createCell(19).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(19).getReference()+")/2");
        /*row.createCell(20).setCellFormula(
                row.createCell(19).getReference()+"*100/"+
                row.createCell(18).getReference());*/
        sheet.getRow(2).createCell(20).setCellFormula(sheet.getRow(2).getCell(19).getReference()+"*100/"
                +sheet.getRow(2).getCell(18).getReference());//合格率

    }

    /*
    给指定的sheet的指定行rowrecord设置汇总数据
     */
    public void setTotalData(XSSFWorkbook xwb, XSSFRow rowrecord, ArrayList<XSSFRow> start, ArrayList<XSSFRow> end, int c1, int c2, int c3, int s1, int s2, int s3, int s4) {
        XSSFSheet sheet = xwb.getSheet("保护层");
        XSSFCellStyle cellstyle = xwb.createCellStyle();
        cellstyle.setAlignment(HorizontalAlignment.CENTER); // 居中
        String H = "";
        String I = "";
        String J = "";

        for (int index = 0; index < start.size(); index++) {
            H = "COUNT(" + start.get(index).getCell(c1).getReference() + ":"
                    + end.get(index).getCell(c1).getReference() + ")+";
            I = "COUNTIF(" + start.get(index).getCell(c2).getReference() + ":"
                    + end.get(index).getCell(c2).getReference() + ",\"√\")+";
            J = "COUNTIF(" + start.get(index).getCell(c3).getReference() + ":"
                    + end.get(index).getCell(c3).getReference() + ",\"×\")+";
            start.get(index).createCell(s1).setCellFormula(H.substring(0, H.length() - 1));// H=COUNT(E216:E245)
            start.get(index).getCell(s1).setCellStyle(cellstyle);

            start.get(index).createCell(s2).setCellFormula(I.substring(0, I.length() - 1));// I=COUNTIF(F41:F60,"√")
            start.get(index).getCell(s2).setCellStyle(cellstyle);

            start.get(index).createCell(s3).setCellFormula(J.substring(0, J.length() - 1));// J=COUNTIF(G216:G245,"×")
            start.get(index).getCell(s3).setCellStyle(cellstyle);

        }
        setTotalTitle(sheet.getRow(rowrecord.getRowNum() - 2), cellstyle, s1,
                s2, s3, s4);

        sheet.getRow(rowrecord.getRowNum() - 1).createCell(s1).setCellFormula("SUM("
                +start.get(0).getCell(s1).getReference()+":"
                +start.get(start.size()-1).getCell(s1).getReference()+")");// H=COUNT(E216:E245)
        sheet.getRow(rowrecord.getRowNum() - 1).getCell(s1).setCellStyle(cellstyle);

        sheet.getRow(rowrecord.getRowNum() - 1).createCell(s2).setCellFormula("SUM("
                +start.get(0).getCell(s2).getReference()+":"
                +start.get(start.size()-1).getCell(s2).getReference()+")");// I=COUNTIF(F41:F60,"√")
        sheet.getRow(rowrecord.getRowNum() - 1).getCell(s2).setCellStyle(cellstyle);



        sheet.getRow(rowrecord.getRowNum() - 1).createCell(s3).setCellFormula("SUM("
                +start.get(0).getCell(s3).getReference()+":"
                +start.get(start.size()-1).getCell(s3).getReference()+")");// J=COUNTIF(G216:G245,"×")
        sheet.getRow(rowrecord.getRowNum() - 1).getCell(s3).setCellStyle(cellstyle);

        sheet.getRow(rowrecord.getRowNum() - 1).createCell(s4).setCellFormula(
                sheet.getRow(rowrecord.getRowNum() - 1).getCell(s2).getReference() + "/"
                        + sheet.getRow(rowrecord.getRowNum() - 1).getCell(s1).getReference() + "*100");// K==I41/H41*100
        sheet.getRow(rowrecord.getRowNum() - 1).getCell(s4).setCellStyle(cellstyle);

    }

    /**
     * 根据给定的行rowtitle，为汇总数据设置标题
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
     * 根据给定的单元格起始行号，得到合并单元格的最后一行行号 如果给定的初始行号不是合并单元格，那么函数返回初始行号
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
     * 桥梁的名称可能后面带有（1），（2）或者(1)，(2),所以要过滤掉括号后面的序号
     * @param name
     * @return
     */
    public String getBridgeName(String name){
        if(name.contains("(")){
            return name.substring(0,name.indexOf("("));
        }
        else if(name.contains("（")){
            return name.substring(0,name.indexOf("（"));
        }
        return name;
    }

    /**
     * 根据传入的字符串，判断允许偏差的数值
     * @param offset
     * @return
     */
    public String[] analysisOffset(String offset) {
        String result[] = new String[2];
        if(offset.contains(",")){
            String [] temp = offset.split("[+,-]");
            if(offset.contains("+") && offset.contains("-")){
                result[0] = temp[1];
                result[1] = temp[temp.length-1];
            }else if(offset.contains("+")){
                result[0] = temp[1];
                result[1] = temp[temp.length-1];
            }else{
                result[0] = temp[0];
                result[1] = temp[temp.length-1];
            }
        }
        else{
            result[0] = offset.substring(1);
            result[1] = result[0];
        }
        return result;
    }

    /**
     * 写入数据
     * @param data
     * @param wb
     * @return
     * @throws ParseException
     */
    private boolean DBtoExcel(List<JjgFbgcQlgcSbBhchd> data, XSSFWorkbook wb) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        XSSFSheet sheet = wb.getSheet("保护层");
        sheet.setColumnHidden(10, true);
        sheet.setColumnHidden(11, true);
        sheet.setColumnHidden(12, true);
        String position = data.get(0).getGjbhjjcbw();
        String name = data.get(0).getQlmc();
        String testtime = simpleDateFormat.format(data.get(0).getJcsj());;
        int index = 0;
        int tableNum = 0;
        fillTitleCellData(sheet, tableNum,data.get(0).getHtd(), data.get(0).getProname());
        sheet.getRow(tableNum*36 + 6 + index).getCell(0).setCellValue(name);
        for(JjgFbgcQlgcSbBhchd row:data){
            if(name.equals(row.getQlmc())){
                testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
                if(!position.equals(row.getGjbhjjcbw())){
                    if(index% 6 != 0){
                        index += 6 - index%6;
                        position = row.getGjbhjjcbw();
                    }
                    else{
                        position = row.getGjbhjjcbw();
                    }
                }
                if(index/10 == 3){
                    sheet.getRow(tableNum*36+2).getCell(8+7).setCellValue(testtime);
                    testtime = simpleDateFormat.format(row.getJcsj());
                    tableNum ++;
                    fillTitleCellData(sheet, tableNum,row.getHtd(), row.getProname());
                    index = 0;
                }
                fillCommonCellData(sheet, tableNum, index+6, row);
                index+=1;
            }
            else{
                name = row.getQlmc();
                position = row.getGjbhjjcbw();
                index = 0;
                sheet.getRow(tableNum*36+2).getCell(8+7).setCellValue(testtime);
                testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
                testtime = simpleDateFormat.format(row.getJcsj());
                tableNum ++;
                fillTitleCellData(sheet, tableNum, row.getHtd(), row.getProname());
                sheet.getRow(tableNum*36 + 6 + index).getCell(0).setCellValue(name);
                fillCommonCellData(sheet, tableNum, index+6, row);
                index +=1;
            }
        }
        sheet.getRow(tableNum*36+2).getCell(8+7).setCellValue(testtime);
        return true;
    }

    /**
     * 填写数据
     * @param sheet
     * @param tableNum
     * @param index
     * @param row
     */
    private void fillCommonCellData(XSSFSheet sheet, int tableNum, int index, JjgFbgcQlgcSbBhchd row) {
        if(index == 6){
            sheet.getRow(tableNum*36+index).getCell(0).setCellValue(row.getQlmc());
        }
        if(index%6 == 0){
            sheet.getRow(tableNum*36+index).getCell(1).setCellValue(row.getGjbhjjcbw());
        }
        sheet.getRow(tableNum*36+index).getCell(2).setCellValue(Double.parseDouble(row.getGjzj()));
        sheet.getRow(tableNum*36+index).getCell(3).setCellValue(Double.parseDouble(row.getSjz()));
        sheet.getRow(tableNum*36+index).getCell(4).setCellValue(Double.parseDouble(row.getScz1()));
        sheet.getRow(tableNum*36+index).getCell(5).setCellValue(Double.parseDouble(row.getScz2()));
        sheet.getRow(tableNum*36+index).getCell(6).setCellValue(Double.parseDouble(row.getScz3()));
        sheet.getRow(tableNum*36+index).getCell(7).setCellValue(Double.parseDouble(row.getScz4()));
        sheet.getRow(tableNum*36+index).getCell(8).setCellValue(Double.parseDouble(row.getScz5()));
        sheet.getRow(tableNum*36+index).getCell(9).setCellValue(Double.parseDouble(row.getScz6()));
        sheet.getRow(tableNum*36+index).getCell(13).setCellValue(Double.parseDouble(row.getXzz()));//14
        if((row.getYxwcz()).equals(row.getYxwcf())){
            sheet.getRow(tableNum*36+index).getCell(8+7).setCellValue("±"+Double.valueOf(row.getYxwcz()).intValue());
        }
        else{
            String piancha = "";
            String piancha_ = "";
            if(new Double(Double.parseDouble(row.getYxwcz())).intValue() != 0){
                piancha = "+"+new Double(Double.parseDouble(row.getYxwcz())).intValue();
            }else{
                piancha = "0";
            }
            if(new Double(Double.parseDouble(row.getYxwcf())).intValue() != 0){
                piancha_ += "-"+new Double(Double.parseDouble(row.getYxwcf())).intValue();
            }else{
                piancha_ = "0";
            }
            sheet.getRow(tableNum*36+index).getCell(8+7).setCellValue(piancha+","+piancha_);
        }
    }

    /**
     * 填写标题
     * @param sheet
     * @param tableNum
     * @param htd
     * @param proname
     */
    private void fillTitleCellData(XSSFSheet sheet, int tableNum, String htd,String proname) {
        sheet.getRow(tableNum*36+1).getCell(2).setCellValue(proname);
        sheet.getRow(tableNum*36+1).getCell(8+7).setCellValue(htd);
        sheet.getRow(tableNum*36+2).getCell(2).setCellValue("桥梁上部");
    }

    /**
     * 创建页
     * @param tableNum
     * @param wb
     */
    private void createTable(int tableNum, XSSFWorkbook wb) {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            RowCopy.copyRows(wb, "保护层", "保护层", 0, 35, i * 36);
        }
        if(record >= 1)
            wb.setPrintArea(wb.getSheetIndex("保护层"), 0, 10+7, 0, record * 36-1);
    }

    /**
     * 获取创建页的数量
     * @return
     */
    private int gettableNum(List<Map<String,Object>> listmap) {
        int size = 0;
        int n = 0;
        for (Map<String, Object> m : listmap)
        {
            for (String k : m.keySet()){
                String qlnum = m.get(k).toString();
                Integer nums = Integer.valueOf(qlnum);
                if (nums % 30 == 0){
                    size += nums/30;
                }else {
                    size += nums/30+1;
                }
            }
        }
        return size;
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        String title = "钢筋保护层厚度质量鉴定表";
        String sheetname = "保护层";

        DecimalFormat df = new DecimalFormat(".00");
        DecimalFormat decf = new DecimalFormat("0.##");
        //获取鉴定表文件
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"31桥梁上部保护层厚度.xlsx");
        if(!f.exists()){
            return null;
        }else {
            XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));
            //读取工作表
            XSSFSheet slSheet = xwb.getSheet(sheetname);
            XSSFCell bt = slSheet.getRow(0).getCell(0);
            XSSFCell xmname = slSheet.getRow(1).getCell(2);//陕西高速
            XSSFCell htdname = slSheet.getRow(1).getCell(15);//LJ-1
            XSSFCell hd = slSheet.getRow(2).getCell(2);//涵洞
            List<Map<String,Object>> mapList = new ArrayList<>();
            Map<String,Object> jgmap = new HashMap<>();
            if(proname.equals(xmname.toString()) && title.equals(bt.toString()) && htd.equals(htdname.toString()) && fbgc.equals(hd.toString())){
                slSheet.getRow(2).getCell(18).setCellType(CellType.STRING);
                slSheet.getRow(2).getCell(19).setCellType(CellType.STRING);
                slSheet.getRow(2).getCell(20).setCellType(CellType.STRING);

                jgmap.put("总点数",decf.format(Double.valueOf(slSheet.getRow(2).getCell(18).getStringCellValue())));
                jgmap.put("合格点数",decf.format(Double.valueOf(slSheet.getRow(2).getCell(19).getStringCellValue())));
                jgmap.put("合格率",df.format(Double.valueOf(slSheet.getRow(2).getCell(20).getStringCellValue())));
                mapList.add(jgmap);
                return mapList;
            }else {
                return null;
            }

        }
    }

    /**
     * 导出数据模板文件
     * @param response
     */
    @Override
    public void export(HttpServletResponse response) {
        String fileName = "07桥梁上部保护层厚度实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcQlgcSbBhchdVo()).finish();

    }

    /**
     * 导入实测数据
     * @param file
     * @param commonInfoVo
     */
    @Override
    public void importbhchd(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcQlgcSbBhchdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcQlgcSbBhchdVo>(JjgFbgcQlgcSbBhchdVo.class) {
                                @Override
                                public void handle(List<JjgFbgcQlgcSbBhchdVo> dataList) {
                                    for(JjgFbgcQlgcSbBhchdVo sbJgccVo: dataList)
                                    {
                                        JjgFbgcQlgcSbBhchd fbgcQlgcSbBhchd = new JjgFbgcQlgcSbBhchd();
                                        BeanUtils.copyProperties(sbJgccVo,fbgcQlgcSbBhchd);
                                        fbgcQlgcSbBhchd.setCreatetime(new Date());
                                        fbgcQlgcSbBhchd.setProname(commonInfoVo.getProname());
                                        fbgcQlgcSbBhchd.setHtd(commonInfoVo.getHtd());
                                        fbgcQlgcSbBhchd.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcQlgcSbBhchdMapper.insert(fbgcQlgcSbBhchd);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }
}
