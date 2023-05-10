package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcQlgcXbSzd;
import glgc.jjgys.model.project.JjgFbgcQlgcXbTqd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.qlgc.JjgFbgcQlgcXbSzdVo;
import glgc.jjgys.model.projectvo.qlgc.JjgFbgcQlgcXbTqdVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcQlgcXbSzdMapper;
import glgc.jjgys.system.service.JjgFbgcLjgcZdgqdService;
import glgc.jjgys.system.service.JjgFbgcQlgcXbSzdService;
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
 * @since 2023-03-22
 */
@Service
public class JjgFbgcQlgcXbSzdServiceImpl extends ServiceImpl<JjgFbgcQlgcXbSzdMapper, JjgFbgcQlgcXbSzd> implements JjgFbgcQlgcXbSzdService {

    @Autowired
    private JjgFbgcQlgcXbSzdMapper jjgFbgcQlgcXbSzdMapper;

    @Autowired
    private JjgFbgcLjgcZdgqdService jjgFbgcLjgcZdgqdService;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获取数据
        QueryWrapper<JjgFbgcQlgcXbSzd> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("qlmc","dth");
        List<JjgFbgcQlgcXbSzd> data = jjgFbgcQlgcXbSzdMapper.selectList(wrapper);

        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"28桥梁下部墩台垂直度.xlsx");
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
            String name = "桥梁下部竖直度17规范表格.xlsx";
            String path = reportPath + File.separator + name;
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            createTable(gettableNum(data.size()),wb);
            if(DBtoExcel(data,wb)){
                calculateVerticalDegreeSheet(wb);
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
     * 对“竖直度”sheet进行计算 本表中，由于每座桥的数据都比较多，出现了连续多张表的数据是同一个桥的，也可能一张表的前后部分为不同桥，
     * 	所以在计算桥的汇总数据时，要将多张表的数据都加起来 但值得注意的是，每两张表的数据之间还包含了第二张表的表头，所以要把表头的数据跳过
     * 	此处使用两个List来分别存放每座桥在每张表中数据的开始Row和结束Row
     * 	当一座桥的数据计算完毕时，遍历出List中的每张表的开始和结束Row，进行汇总计算 最后，计算完毕，清空List
     * @param wb
     */
    private void calculateVerticalDegreeSheet(XSSFWorkbook wb) {
        XSSFSheet sheet = wb.getSheet("竖直度");
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
            if (row.getCell(2) != null && !"".equals(row.getCell(2).toString()) && flag) {
                //=IF(C6>60000,MIN(C6/3000,30),IF(5000<C6,MIN(C6/1000,20),IF(C6<=5000,5)))
                row.getCell(3).setCellFormula(
                        "IF("+row.getCell(2).getReference()+">60000,MIN("+row.getCell(2).getReference()
                                +"/3000,30),IF(5000<"+row.getCell(2).getReference()+
                                ",MIN("+row.getCell(2).getReference()+"/1000,20),IF("+row.getCell(2).getReference()+"<=5000,5)))");
                XSSFRow r=sheet.getRow(i+1);
                r.getCell(3).setCellFormula(
                        "IF("+row.getCell(2).getReference()+">60000,MIN("+row.getCell(2).getReference()
                                +"/3000,30),IF(5000<"+row.getCell(2).getReference()+
                                ",MIN("+row.getCell(2).getReference()+"/1000,20),IF("+row.getCell(2).getReference()+"<=5000,5)))");
                sheet.addMergedRegion(new CellRangeAddress(i, i+1, 3, 3));
            }
            // 可以计算
            if (row.getCell(5) != null && !"".equals(row.getCell(5).toString()) && flag) {
                row.getCell(6).setCellFormula(
                        "IF(" + row.getCell(5).getReference() + "-"
                                + row.getCell(3).getReference() + "<=" + 0
                                + ",\"√\",\"\")");// F=IF(E6-C6<=0,"√","")
                row.getCell(7).setCellFormula(
                        "IF(" + row.getCell(5).getReference() + "-"
                                + row.getCell(3).getReference() + ">" + 0
                                + ",\"×\",\"\")");// G=IF(E6-C6>0,"×","")
            }
            // 到下一个表了
            if (!"".equals(row.getCell(0).toString())&& !"".equals(row.getCell(0).toString())
                    && row.getCell(0).toString().contains("质量鉴定表") && flag) {
                flag = false;
            }
            // 下一座桥
            if (!"".equals(row.getCell(0).toString())&& !"".equals(row.getCell(0).toString()) && !name.equals(jjgFbgcLjgcZdgqdService.getBridgeName(row.getCell(0).toString())) && flag) {
                if (start.size() > 0) {
                    rowrecord = start.get(0);
                    setTotalData(wb, rowrecord, start, end, 5, 6, 7, 8, 9, 10,11);
                }
                start.clear();
                end.clear();
                name = jjgFbgcLjgcZdgqdService.getBridgeName(row.getCell(0).toString());
                rowstart = sheet.getRow(i);
                rowrecord = rowstart;

                rowend = sheet.getRow(getCellEndRow(sheet, i, 0));
                start.add(rowstart);
                end.add(rowend);
                bridgename.add(rowstart.getCell(0).toString());
            }
            // 同一座桥，但是在不同的单元格
            else if (!"".equals(row.getCell(0).toString())&& !"".equals(row.getCell(0).toString())
                    && name.equals(jjgFbgcLjgcZdgqdService.getBridgeName(row.getCell(0).toString())) && flag
                    && rowstart != row) {
                rowstart = sheet.getRow(i);
                rowend = sheet.getRow(getCellEndRow(sheet, i, 0));
                start.add(rowstart);
                end.add(rowend);
            }
            // 可以计算啦
            if (row.getCell(0) != null &&"桥 梁      名 称".equals(row.getCell(0).toString())) {
                flag = true;
                i += 1;
                rowstart = sheet.getRow(i + 1);
                rowrecord = rowstart;
                rowend = sheet.getRow(getCellEndRow(sheet, i + 1, 0));
                if (name.equals(jjgFbgcLjgcZdgqdService.getBridgeName(rowstart.getCell(0).toString()))) {  //与前边一座桥 相同
                    start.add(rowstart);  //计算 当前
                    end.add(rowend);

                } else {//不同
                    if (start.size() > 0) {
                        rowrecord = start.get(0);
                        setTotalData(wb, rowrecord, start, end,  5, 6, 7, 8, 9, 10,11);
                    }
                    start.clear();
                    end.clear();
                    name = jjgFbgcLjgcZdgqdService.getBridgeName(rowstart.getCell(0).toString());
                    start.add(rowstart);
                    end.add(rowend);
                    bridgename.add(rowstart.getCell(0).toString());
                }
            }
        }
        if (start.size() > 0) {
            rowrecord = start.get(0);
            setTotalData(wb, rowrecord, start, end, 5, 6, 7, 8, 9, 10,11);
        }
        row = sheet.getRow(2);
        row.createCell(8).setCellFormula("SUM("
                +sheet.getRow(3).createCell(8).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(8).getReference()+")");//=SUM(L7:L300)
        row.createCell(9).setCellFormula("SUM("
                +sheet.getRow(3).createCell(9).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(9).getReference()+")");//=SUM(L7:L300)
        sheet.getRow(2).createCell(10).setCellFormula(sheet.getRow(2).getCell(9).getReference()+"*100/"
                +sheet.getRow(2).getCell(8).getReference());//合格率

    }

    public void setTotalData(XSSFWorkbook wb, XSSFRow rowrecord, ArrayList<XSSFRow> start, ArrayList<XSSFRow> end, int c1, int c2, int c3, int s1, int s2, int s3, int s4) {
        XSSFSheet sheet = wb.getSheet("竖直度");
        XSSFCellStyle cellstyle = wb.createCellStyle();
        cellstyle.setAlignment(HorizontalAlignment.CENTER); // 居中
        //String H = "";
        String I = "";
        String J = "";
        String K = "";
        for (int index = 0; index < start.size(); index++) {
            I += "COUNT(" + start.get(index).getCell(c1).getReference() + ":"
                    + end.get(index).getCell(c1).getReference() + ")+";
            J += "COUNTIF(" + start.get(index).getCell(c2).getReference() + ":"
                    + end.get(index).getCell(c2).getReference() + ",\"√\")+";
            K += "COUNTIF(" + start.get(index).getCell(c3).getReference() + ":"
                    + end.get(index).getCell(c3).getReference() + ",\"×\")+";
        }
        setTotalTitle(sheet.getRow(rowrecord.getRowNum() - 1), cellstyle, s1,
                s2, s3, s4);
        rowrecord.createCell(s1).setCellFormula(I.substring(0, I.length() - 1));// H=COUNT(E216:E245)
        rowrecord.getCell(s1).setCellStyle(cellstyle);

        rowrecord.createCell(s2).setCellFormula(J.substring(0, J.length() - 1));// I=COUNTIF(F41:F60,"√")
        rowrecord.getCell(s2).setCellStyle(cellstyle);

        rowrecord.createCell(s3).setCellFormula(K.substring(0, K.length() - 1));// J=COUNTIF(G216:G245,"×")
        rowrecord.getCell(s3).setCellStyle(cellstyle);

        rowrecord.createCell(s4).setCellFormula(
                rowrecord.getCell(s2).getReference() + "/"
                        + rowrecord.getCell(s1).getReference() + "*100");// K==I41/H41*100
        rowrecord.getCell(s4).setCellStyle(cellstyle);

    }

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

    private boolean DBtoExcel(List<JjgFbgcQlgcXbSzd> data, XSSFWorkbook wb) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");

            int totalTable = gettableNum(data.size());
            XSSFSheet sheet = wb.getSheet("竖直度");

            String name = data.get(0).getQlmc();
            String testtime = simpleDateFormat.format(data.get(0).getJcsj());

            int index = 0;
            int tableNum = 0;
            fillTitleCellData(sheet, tableNum,data.get(0));
            for(JjgFbgcQlgcXbSzd row:data){
                if(name.equals(row.getQlmc())){
                    testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
                    if(index/10 == 3){
                        sheet.getRow(tableNum*35+2).getCell(7).setCellValue(testtime);
                        tableNum ++;
                        if(tableNum <totalTable){
                            fillTitleCellData(sheet, tableNum, row);
                        }
                        index = 0;
                    }
                    if(index%10 == 0 && tableNum <totalTable){
                        sheet.getRow(tableNum*35 + 5 + index).getCell(0).setCellValue(name);
                    }
                    if(tableNum <totalTable){
                        fillCommonCellData(sheet, tableNum, index+5, row);
                    }
                    index+=2;
                } else{
                    name = row.getQlmc();
                    index = 0;
                    sheet.getRow(tableNum*35+2).getCell(7).setCellValue(testtime);
                    testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
                    tableNum ++;
                    fillTitleCellData(sheet, tableNum, row);
                    sheet.getRow(tableNum*35 + 5 + index).getCell(0).setCellValue(name);
                    fillCommonCellData(sheet, tableNum, index+5, row);
                    index +=2;
                }
            }
            if(tableNum <totalTable){
                sheet.getRow(tableNum*35+2).getCell(7).setCellValue(testtime);
            }
            return true;
    }

    private void fillCommonCellData(XSSFSheet sheet, int tableNum, int index, JjgFbgcQlgcXbSzd row) {
        sheet.getRow(tableNum*35+index).getCell(1).setCellValue(row.getDth());
        sheet.getRow(tableNum*35+index).getCell(2).setCellValue(Double.parseDouble(row.getDzgd()));
        sheet.getRow(tableNum*35+index).getCell(4).setCellValue("横桥向");
        sheet.getRow(tableNum*35+index).getCell(5).setCellValue(Double.parseDouble(row.getHxscz()));
        sheet.getRow(tableNum*35+index+1).getCell(4).setCellValue("纵桥向");
        sheet.getRow(tableNum*35+index+1).getCell(5).setCellValue(Double.parseDouble(row.getZxscz()));

    }

    public void fillTitleCellData(XSSFSheet sheet, int tableNum, JjgFbgcQlgcXbSzd row) {
        sheet.getRow(tableNum*35+1).getCell(1).setCellValue(row.getProname());
        sheet.getRow(tableNum*35+1).getCell(7).setCellValue(row.getHtd());
        sheet.getRow(tableNum*35+2).getCell(1).setCellValue(row.getFbgc());

    }

    private void createTable(int gettableNum, XSSFWorkbook wb) {
        int record = 0;
        record = gettableNum;
        for (int i = 1; i < record; i++) {
            RowCopy.copyRows(wb, "竖直度", "竖直度", 0, 34, i * 35);
        }
        if(record >= 1)
            wb.setPrintArea(wb.getSheetIndex("竖直度"), 0, 7, 0, record * 35-1);
    }

    private int gettableNum(int size) {
        return size%15 ==0 ? size/15 : size/15+1;
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        String title = "墩台竖直度质量鉴定表";
        String sheetname = "竖直度";

        DecimalFormat df = new DecimalFormat(".00");
        DecimalFormat decf = new DecimalFormat("0.##");
        //获取鉴定表文件
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"28桥梁下部墩台垂直度.xlsx");
        if(!f.exists()){
            return null;
        }else {
            XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));
            //读取工作表
            XSSFSheet slSheet = xwb.getSheet(sheetname);
            XSSFCell bt = slSheet.getRow(0).getCell(0);
            XSSFCell xmname = slSheet.getRow(1).getCell(1);//陕西高速
            XSSFCell htdname = slSheet.getRow(1).getCell(7);//LJ-1
            XSSFCell hd = slSheet.getRow(2).getCell(1);//涵洞1
            List<Map<String,Object>> mapList = new ArrayList<>();
            Map<String,Object> jgmap = new HashMap<>();
            if(proname.equals(xmname.toString()) && title.equals(bt.toString()) && htd.equals(htdname.toString()) && fbgc.equals(hd.toString())){
                slSheet.getRow(2).getCell(8).setCellType(CellType.STRING);
                slSheet.getRow(2).getCell(9).setCellType(CellType.STRING);
                slSheet.getRow(2).getCell(10).setCellType(CellType.STRING);

                jgmap.put("总点数",decf.format(Double.valueOf(slSheet.getRow(2).getCell(8).getStringCellValue())));
                jgmap.put("合格点数",decf.format(Double.valueOf(slSheet.getRow(2).getCell(9).getStringCellValue())));
                jgmap.put("合格率",df.format(Double.valueOf(slSheet.getRow(2).getCell(10).getStringCellValue())));
                mapList.add(jgmap);
                return mapList;
            }else {
                return null;
            }

        }
    }

    @Override
    public void exportqlxbszd(HttpServletResponse response) {
        String fileName = "04桥梁下部竖直度实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcQlgcXbSzdVo()).finish();


    }

    @Override
    public void importqlxbszd(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcQlgcXbSzdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcQlgcXbSzdVo>(JjgFbgcQlgcXbSzdVo.class) {
                                @Override
                                public void handle(List<JjgFbgcQlgcXbSzdVo> dataList) {
                                    for(JjgFbgcQlgcXbSzdVo xbSzdVo: dataList)
                                    {
                                        JjgFbgcQlgcXbSzd fbgcQlgcXbSzd = new JjgFbgcQlgcXbSzd();
                                        BeanUtils.copyProperties(xbSzdVo,fbgcQlgcXbSzd);
                                        fbgcQlgcXbSzd.setCreatetime(new Date());
                                        fbgcQlgcXbSzd.setProname(commonInfoVo.getProname());
                                        fbgcQlgcXbSzd.setHtd(commonInfoVo.getHtd());
                                        fbgcQlgcXbSzd.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcQlgcXbSzdMapper.insert(fbgcQlgcXbSzd);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }


    }
}
