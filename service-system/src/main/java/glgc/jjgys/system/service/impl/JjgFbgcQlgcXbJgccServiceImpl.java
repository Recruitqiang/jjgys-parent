package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcQlgcXbJgcc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.qlgc.JjgFbgcQlgcXbJgccVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcQlgcXbJgccMapper;
import glgc.jjgys.system.service.JjgFbgcQlgcXbJgccService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
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
 * @since 2023-03-22
 */
@Service
public class JjgFbgcQlgcXbJgccServiceImpl extends ServiceImpl<JjgFbgcQlgcXbJgccMapper, JjgFbgcQlgcXbJgcc> implements JjgFbgcQlgcXbJgccService {


    @Autowired
    private JjgFbgcQlgcXbJgccMapper jjgFbgcQlgcXbJgccMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获得数据
        QueryWrapper<JjgFbgcQlgcXbJgcc> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("dth");
        List<JjgFbgcQlgcXbJgcc> data = jjgFbgcQlgcXbJgccMapper.selectList(wrapper);

        List<Map<String,Object>> selectnum = jjgFbgcQlgcXbJgccMapper.selectnum(proname, htd, fbgc);

        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"26桥梁下部主要结构尺寸.xlsx");
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
            String name = "桥梁上部尺寸.xlsx";
            String path = reportPath + File.separator + name;
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream in = new FileInputStream(f);
            wb = new XSSFWorkbook(in);

            int tablename = gettableNum(selectnum);
            createTable(tablename,wb);
            if(DBtoExcel(data,wb,tablename)){
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
     * 对“尺寸”sheet进行计算处理
     * @param wb
     */
    private void calculateSizeSheet(XSSFWorkbook wb) {
        XSSFSheet sheet = wb.getSheet("尺寸");
        XSSFRow row = null;
        boolean flag = false;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        XSSFRow rowrecord = null;
        ArrayList<String> bridgename = new ArrayList<String>();
        ArrayList<XSSFRow> start = new ArrayList<XSSFRow>();
        ArrayList<XSSFRow> end = new ArrayList<XSSFRow>();
        String name = "";
        for (int i = sheet.getFirstRowNum(); i <= sheet
                .getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            // 计算
            if (row.getCell(1) != null && !"".equals(row.getCell(1).toString()) && flag) {
                row.getCell(7).setCellFormula(
                        "IF(((" + row.getCell(5).getReference() + "-"
                                + row.getCell(4).getReference() + ")<="
                                + analysisOffset(row.getCell(6).toString())[0]+")*AND(("
                                + row.getCell(4).getReference() + "-"
                                + row.getCell(5).getReference() + ")<="
                                + analysisOffset(row.getCell(6).toString())[1]
                                + "),\"√\",\"\")");
                // H=IF(ABS(F6-E6)<=15,"√","")=IF(ABS(F6-E6)<=20,"√","")
                //=IF(((F6-E6)<=15)*AND((E6-F6)<=15),"√","")
                row.getCell(8).setCellFormula(
                        "IF(((" + row.getCell(5).getReference() + "-"
                                + row.getCell(4).getReference() + ")<="
                                + analysisOffset(row.getCell(6).toString())[0]+")*AND(("
                                + row.getCell(4).getReference() + "-"
                                + row.getCell(5).getReference() + ")<="
                                + analysisOffset(row.getCell(6).toString())[1]
                                + "),\"\",\"×\")");// I=IF(ABS(F6-E6)<=15,"","×")=IF(ABS(F6-E6)<=20,"","×")
                row.createCell(13).setCellFormula(
                        row.getCell(5).getReference() + "-"
                                + row.getCell(4).getReference());// =F6-E6
            }
            // 到下一个表了
            if (row.getCell(1) != null && !"".equals(row.getCell(0).toString())
                    && row.getCell(0).toString().contains("质量鉴定表") && flag) {
                flag = false;
            }
            // 下一座桥
            if (row.getCell(0) != null && !"".equals(row.getCell(0).toString())
                    && !name.equals(getBridgeName(row.getCell(0).toString())) && flag) {
                if (start.size() > 0) {
                    rowrecord = start.get(0);
                    setTotalData(sheet, rowrecord, start, end, 5, 7, 8, 9, 10, 11, 12,wb);
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
            }
            // 同一座桥，但是在不同的单元格
            else if (row.getCell(0) != null && !"".equals(row.getCell(0).toString())
                    && name.equals(getBridgeName(row.getCell(0).toString())) && flag
                    && rowstart != row) {
                rowstart = sheet.getRow(i);
                rowend = sheet.getRow(getCellEndRow(sheet, i, 0));
                start.add(rowstart);
                end.add(rowend);

            }
            // 可以计算啦
            if (row.getCell(0) != null && "桥梁   名称".equals(row.getCell(0).toString())) {
                flag = true;
                i += 1;
                rowstart = sheet.getRow(i + 1);
                rowrecord = rowstart;
                rowend = sheet.getRow(getCellEndRow(sheet, i + 1, 0));
                if (name.equals(getBridgeName(rowstart.getCell(0).toString()))) {
                    start.add(rowstart);
                    end.add(rowend);

                } else {
                    if (start.size() > 0) {
                        rowrecord = start.get(0);
                        setTotalData(sheet, rowrecord, start, end, 5, 7, 8, 9, 10, 11, 12,wb);
                    }
                    start.clear();
                    end.clear();
                    name = getBridgeName(rowstart.getCell(0).toString());
                    start.add(rowstart);
                    end.add(rowend);
                    bridgename.add(name);
                }
                System.out.println();
            }
        }
        if (start.size() > 0) {
            rowrecord = start.get(0);
            setTotalData(sheet, rowrecord, start, end, 5, 7, 8, 9, 10, 11, 12,wb);
        }
        row = sheet.getRow(2);
        row.createCell(9).setCellFormula("SUM("
                +sheet.getRow(3).createCell(9).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(9).getReference()+")");//=SUM(L7:L300)
        row.createCell(10).setCellFormula("SUM("
                +sheet.getRow(3).createCell(10).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(10).getReference()+")");//=SUM(L7:L300)
        sheet.getRow(2).createCell(11).setCellFormula(sheet.getRow(2).getCell(10).getReference()+"*100/"
                +sheet.getRow(2).getCell(9).getReference());//合格率
    }

    /**
     * 根据传入的字符串，判断允许偏差的数值
     * @param offset
     * @return
     */
    private String[] analysisOffset(String offset) {
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
     * 根据给定的单元格起始行号，得到合并单元格的最后一行行号 如果给定的初始行号不是合并单元格，那么函数返回初始行号
     * @param sheet
     * @param cellstartrow
     * @param cellstartcol
     * @return
     */
    private int getCellEndRow(XSSFSheet sheet, int cellstartrow, int cellstartcol) {
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
     *给指定的sheet的指定行rowrecord设置汇总数据
     * @param sheet
     * @param rowrecord
     * @param start
     * @param end
     * @param c1  总点数
     * @param c2  合格点数
     * @param c3  不合格点数
     * @param s1 各个子项的列
     * @param s2
     * @param s3
     * @param s4
     * @param wb
     */
    private void setTotalData(XSSFSheet sheet, XSSFRow rowrecord, ArrayList<XSSFRow> start, ArrayList<XSSFRow> end, int c1, int c2, int c3, int s1, int s2, int s3, int s4,XSSFWorkbook wb) {
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
        setTotalTitle(sheet.getRow(rowrecord.getRowNum() - 1), cellstyle, s1,
                s2, s3, s4);
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
     *根据给定的行rowtitle，为汇总数据设置标题
     * @param rowtitle
     * @param cellstyle
     * @param s1
     * @param s2
     * @param s3
     * @param s4
     */
    private void setTotalTitle(XSSFRow rowtitle, XSSFCellStyle cellstyle, int s1, int s2, int s3, int s4) {
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
     *桥梁的名称可能后面带有（1），（2）或者(1)，(2),所以要过滤掉括号后面的序号
     * @param name
     * @return
     */
    private String getBridgeName(String name){
        if(name.contains("(")){
            return name.substring(0,name.indexOf("("));
        }
        else if(name.contains("（")){
            return name.substring(0,name.indexOf("（"));
        }
        return name;
    }

    private boolean DBtoExcel(List<JjgFbgcQlgcXbJgcc> data, XSSFWorkbook wb,int totalTable) throws ParseException {
        XSSFCellStyle cellstyle = wb.createCellStyle();
        XSSFSheet sheet = wb.getSheet("尺寸");
        String name = data.get(0).getQlmc();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        String testtime = simpleDateFormat.format(data.get(0).getJcsj());
        int index = 0;
        int tableNum = 0;
        fillTitleCellData(sheet, tableNum,data.get(0));
        for(JjgFbgcQlgcXbJgcc row:data){
            if(name.equals(row.getQlmc())){
                testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
                if(index/10 == 3){
                    sheet.getRow(tableNum*35+2).getCell(6).setCellValue(testtime);
                    testtime = simpleDateFormat.format(row.getJcsj());
                    tableNum ++;
                    if(totalTable < tableNum){
                        return true;
                    }
                    fillTitleCellData(sheet, tableNum,row);
                    index = 0;
                }
                if(index%10 == 0){
                    sheet.getRow(tableNum*35 + 5 + index).getCell(0).setCellValue(name);
                }
                fillCommonCellData(sheet, tableNum, index+5, row);
                index++;
            }
            else{
                name = row.getQlmc();
                if(index/10 == 0 || index == 10){
                    testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
                    index = 10;
                }
                else if(index/10 == 1 || index == 20){
                    index = 20;
                }
                else if(index/10 == 2 || index == 30){
                    sheet.getRow(tableNum*35+2).getCell(6).setCellValue(testtime);
                    testtime = simpleDateFormat.format(row.getJcsj());
                    testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
                    tableNum ++;
                    fillTitleCellData(sheet, tableNum,row);
                    index = 0;
                }
                sheet.getRow(tableNum*35 + 5 + index).getCell(0).setCellValue(name);
                fillCommonCellData(sheet, tableNum, index+5, row);
                index ++;

            }
        }
        if(totalTable < tableNum){
            return true;
        }
        sheet.getRow(tableNum*35+2).getCell(6).setCellValue(testtime);
        return true;
    }

    /**
     * 填写数据
     * @param sheet
     * @param tableNum
     * @param index
     * @param row
     */
    private void fillCommonCellData(XSSFSheet sheet, int tableNum, int index, JjgFbgcQlgcXbJgcc row) {
        sheet.getRow(tableNum*35+index).getCell(1).setCellValue(row.getLb());
        sheet.getRow(tableNum*35+index).getCell(2).setCellValue(row.getBw());
        sheet.getRow(tableNum*35+index).getCell(3).setCellValue(row.getLb());
        sheet.getRow(tableNum*35+index).getCell(4).setCellValue(Double.parseDouble(row.getSjz()));
        sheet.getRow(tableNum*35+index).getCell(5).setCellValue(Double.parseDouble(row.getScz()));

        if((""+row.getYxwcz()).equals(row.getYxwcf())){
            sheet.getRow(tableNum*35+index).getCell(6).setCellValue("±"+new Double(Double.parseDouble(row.getYxwcz())).intValue());
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
            sheet.getRow(tableNum*35+index).getCell(6).setCellValue(piancha+","+piancha_);
        }
    }

    /**
     * 填写标题
     * @param sheet
     * @param tableNum
     * @param row
     */
    private void fillTitleCellData(XSSFSheet sheet, int tableNum,JjgFbgcQlgcXbJgcc row) {
        sheet.getRow(tableNum*35+1).getCell(2).setCellValue(row.getProname());
        sheet.getRow(tableNum*35+1).getCell(6).setCellValue(row.getHtd());
        sheet.getRow(tableNum*35+2).getCell(2).setCellValue("桥梁下部");

    }

    private void createTable(int gettableNum,XSSFWorkbook wb) {
        int record = 0;
        record = gettableNum;
        for (int i = 1; i < record; i++) {
            RowCopy.copyRows(wb, "尺寸", "尺寸", 0, 34, i * 35);
        }
        if(record >= 1)
            wb.setPrintArea(wb.getSheetIndex("尺寸"), 0, 8, 0, record * 35-1);
    }

    private int gettableNum(List<Map<String,Object>> mapList) {
        int size = 0;
        int n = 0;
        for (Map<String, Object> m : mapList)
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
        String title = "结构（断面）尺寸质量鉴定表";
        String sheetname = "尺寸";

        DecimalFormat df = new DecimalFormat(".00");
        DecimalFormat decf = new DecimalFormat("0.##");
        //获取鉴定表文件
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"26桥梁下部主要结构尺寸.xlsx");
        if(!f.exists()){
            return null;
        }else {
            XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));
            //读取工作表
            XSSFSheet slSheet = xwb.getSheet(sheetname);
            XSSFCell bt = slSheet.getRow(0).getCell(0);
            XSSFCell xmname = slSheet.getRow(1).getCell(2);//陕西高速
            XSSFCell htdname = slSheet.getRow(1).getCell(6);//LJ-1
            XSSFCell hd = slSheet.getRow(2).getCell(2);//涵洞1
            List<Map<String,Object>> mapList = new ArrayList<>();
            Map<String,Object> jgmap = new HashMap<>();
            if(proname.equals(xmname.toString()) && title.equals(bt.toString()) && htd.equals(htdname.toString()) && fbgc.equals(hd.toString())){
                slSheet.getRow(2).getCell(9).setCellType(CellType.STRING);
                slSheet.getRow(2).getCell(10).setCellType(CellType.STRING);
                slSheet.getRow(2).getCell(11).setCellType(CellType.STRING);

                jgmap.put("总点数",decf.format(Double.valueOf(slSheet.getRow(2).getCell(9).getStringCellValue())));
                jgmap.put("合格点数",decf.format(Double.valueOf(slSheet.getRow(2).getCell(10).getStringCellValue())));
                jgmap.put("合格率",df.format(Double.valueOf(slSheet.getRow(2).getCell(11).getStringCellValue())));
                mapList.add(jgmap);
                return mapList;
            }else {
                return null;
            }

        }
    }

    @Override
    public void exportqlxbjgcc(HttpServletResponse response) {
        String fileName = "桥梁下部结构尺寸实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcQlgcXbJgccVo()).finish();

    }

    @Override
    public void importqlxbjgcc(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcQlgcXbJgccVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcQlgcXbJgccVo>(JjgFbgcQlgcXbJgccVo.class) {
                                @Override
                                public void handle(List<JjgFbgcQlgcXbJgccVo> dataList) {
                                    for(JjgFbgcQlgcXbJgccVo xbJgccVo: dataList)
                                    {
                                        JjgFbgcQlgcXbJgcc xbJgcc = new JjgFbgcQlgcXbJgcc();
                                        BeanUtils.copyProperties(xbJgccVo,xbJgcc);
                                        xbJgcc.setCreatetime(new Date());
                                        xbJgcc.setProname(commonInfoVo.getProname());
                                        xbJgcc.setHtd(commonInfoVo.getHtd());
                                        xbJgcc.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcQlgcXbJgccMapper.insert(xbJgcc);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }
}
