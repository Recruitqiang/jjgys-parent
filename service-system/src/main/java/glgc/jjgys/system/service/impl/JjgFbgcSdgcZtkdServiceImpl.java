package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcQlgcXbTqd;
import glgc.jjgys.model.project.JjgFbgcSdgcZtkd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.qlgc.JjgFbgcQlgcXbTqdVo;
import glgc.jjgys.model.projectvo.sdgc.JjgFbgcSdgcZtkdVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcSdgcZtkdMapper;
import glgc.jjgys.system.service.JjgFbgcSdgcZtkdService;
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
 * @since 2023-03-26
 */
@Service
public class JjgFbgcSdgcZtkdServiceImpl extends ServiceImpl<JjgFbgcSdgcZtkdMapper, JjgFbgcSdgcZtkd> implements JjgFbgcSdgcZtkdService {

    @Autowired
    private JjgFbgcSdgcZtkdMapper jjgFbgcSdgcZtkdMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获取数据
        QueryWrapper<JjgFbgcSdgcZtkd> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("sdmc","zh");
        List<JjgFbgcSdgcZtkd> data = jjgFbgcSdgcZtkdMapper.selectList(wrapper);

        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"41隧道总体宽度.xlsx");
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
            String name = "隧道总体宽度.xlsx";
            String path = reportPath + File.separator + name;
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            createTable(gettableNum(data.size()),wb);
            if(DBtoExcel(data,wb)){
                calculateWidthSheet(wb);
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

    private void calculateWidthSheet(XSSFWorkbook wb) {
        XSSFSheet sheet = wb.getSheet("隧道总体宽度");
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
                System.out.println(i);
                continue;
            }
            // 计算
            if (!"".equals(row.getCell(2).toString()) && flag) {
                row.getCell(5).setCellFormula("IF("
                        +row.getCell(2).getReference()+"-"
                        +row.getCell(3).getReference()+">=0,\"√\",\"\")");//F6=IF(E6-D6<=0,"√","")
                row.getCell(6).setCellFormula("IF("
                        +row.getCell(2).getReference()+"-"
                        +row.getCell(3).getReference()+"<0,\"×\",\"\")");//G6=IF(E6-D6>0,"×","")
            }
            // 到下一个表了
            if (!"".equals(row.getCell(0).toString()) && row.getCell(0).toString().contains("鉴定表") && flag) {
                flag = false;
            }
            // 下一座桥
            if (flag && !"".equals(row.getCell(0).toString()) && !name.equals(row.getCell(0).toString())) {

                System.out.println("下一座桥 name "+name+"  i ="+(i)+ " start.size()="+start.size());
                if (start.size() > 0) {
                    rowrecord = start.get(0);
                    setTotalData(wb, rowrecord, start, end, 2, 5, 6,9,10, 11,12);
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
                        setTotalData(wb, rowrecord, start, end, 2, 5, 6,9,10, 11,12);
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
            setTotalData(wb, rowrecord, start, end, 2, 5, 6,9,10, 11,12);
        }
        start.clear();
        end.clear();
        row = sheet.getRow(2);
        row.createCell(9).setCellFormula("SUM("
                +sheet.getRow(3).createCell(9).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(9).getReference()+")");//=SUM(L7:L300)
        row.createCell(10).setCellFormula("SUM("
                +sheet.getRow(3).createCell(10).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(10).getReference()+")");
        sheet.getRow(2).createCell(11).setCellFormula(sheet.getRow(2).getCell(10).getReference()+"*100/"
                +sheet.getRow(2).getCell(9).getReference());//合格率

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

    private void setTotalData(XSSFWorkbook wb, XSSFRow rowrecord, ArrayList<XSSFRow> start, ArrayList<XSSFRow> end, int c1, int c2, int c3, int s1, int s2, int s3, int s4) {
        XSSFSheet sheet = wb.getSheet("隧道总体宽度");
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

    private boolean DBtoExcel(List<JjgFbgcSdgcZtkd> data, XSSFWorkbook wb) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        XSSFSheet sheet = wb.getSheet("隧道总体宽度");
        String zy = data.get(0).getZh().substring(0,1);
        String testtime = simpleDateFormat.format(data.get(0).getJcsj());
        String name = data.get(0).getSdmc();

        int index = 0;
        int tableNum = 0;
        fillTitleCellData(sheet, tableNum, data.get(0));
        for(JjgFbgcSdgcZtkd row:data){
            if(name.equals(row.getSdmc()) && zy.equals(row.getZh().substring(0,1))){ //当前隧道
                testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
                if(index/15 == 2){ //表示到达表格末尾
                    sheet.getRow(tableNum*35+2).getCell(5).setCellValue(testtime);
                    testtime = simpleDateFormat.format(row.getJcsj());
                    tableNum ++;
                    fillTitleCellData(sheet, tableNum,row);
                    index = 0;
                }
                if(index%15 == 0){ //表格开头  隧道名称
                    System.out.println("index= "+index+" name"+name+" i="+(tableNum*35 + 5 + index));
                    sheet.getRow(tableNum*35 + 5 + index).getCell(0).setCellValue(name);//setCellValue(name+"("+(ZY.equals("Z")?"左线":"右线")+")");
                }
                fillCommonCellData(sheet, tableNum, index+5, row);
                index++;
            }
            else{// 不是当前隧道
                name = row.getSdmc();
                zy = row.getZh().substring(0,1);
                if(index/15 == 0){  //  当前在表格 前15行   index<15
                    testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
                    index = 15;
                }
                else if(index/15 == 1&& index%15 > 0){  //当前位置超过半个表格 填写到下一张表  //  在表格 后15行  index>15
                    sheet.getRow(tableNum*35+2).getCell(5).setCellValue(testtime);
                    testtime = simpleDateFormat.format(row.getJcsj());
                    testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
                    tableNum ++;
                    fillTitleCellData(sheet, tableNum,row);
                    index = 0;
                }
                sheet.getRow(tableNum*35 + 5 + index).getCell(0).setCellValue(name);//setCellValue(name+"("+(ZY.equals("Z")?"左线":"右线")+")");
                fillCommonCellData(sheet, tableNum, index+5, row);
                index ++;

            }
        }
        sheet.getRow(tableNum*35+2).getCell(5).setCellValue(testtime);
        return true;
    }

    public void fillCommonCellData(XSSFSheet sheet, int tableNum, int index, JjgFbgcSdgcZtkd row) {
        sheet.getRow(tableNum*35+index).getCell(1).setCellValue(row.getZh());
        sheet.getRow(tableNum*35+index).createCell(7).setCellValue(Double.parseDouble(row.getZbk()));
        sheet.getRow(tableNum*35+index).createCell(8).setCellValue(Double.parseDouble(row.getYbk()));
        sheet.getRow(tableNum*35+index).getCell(2).setCellFormula(
                sheet.getRow(tableNum*35+index).getCell(7).getReference()+"+"+
                        sheet.getRow(tableNum*35+index).getCell(8).getReference());
        sheet.getRow(tableNum*35+index).getCell(3).setCellValue(Double.valueOf(row.getSjz()));
        sheet.getRow(tableNum*35+index).getCell(4).setCellValue("不小于设计值");
    }

    private void fillTitleCellData(XSSFSheet sheet, int tableNum, JjgFbgcSdgcZtkd row) {
        sheet.getRow(tableNum*35+1).getCell(1).setCellValue(row.getProname());
        sheet.getRow(tableNum*35+1).getCell(5).setCellValue(row.getHtd());
        sheet.getRow(tableNum*35+2).getCell(1).setCellValue("衬砌");

    }

    private void createTable(int tableNum, XSSFWorkbook wb) {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            RowCopy.copyRows(wb, "隧道总体宽度", "隧道总体宽度", 0, 34, i * 35);
        }
        if(record > 1)
            wb.setPrintArea(wb.getSheetIndex("隧道总体宽度"), 0, 6, 0, record * 35-1);
    }

    private int gettableNum(int size) {
        return size%20 ==0 ? size/20 : size/20+1;
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        String title = "隧道宽度质量鉴定表";
        String sheetname = "隧道总体宽度";

        DecimalFormat df = new DecimalFormat(".00");
        DecimalFormat decf = new DecimalFormat("0.##");
        //获取鉴定表文件
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"41隧道总体宽度.xlsx");
        if(!f.exists()){
            return null;
        }else {
            XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));
            //读取工作表
            XSSFSheet slSheet = xwb.getSheet(sheetname);
            XSSFCell bt = slSheet.getRow(0).getCell(0);
            XSSFCell xmname = slSheet.getRow(1).getCell(1);//陕西高速
            XSSFCell htdname = slSheet.getRow(1).getCell(5);//LJ-1
            XSSFCell hd = slSheet.getRow(2).getCell(1);//涵洞
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
    public void exportsdztkd(HttpServletResponse response) {
        String fileName = "04隧道总体宽度实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcSdgcZtkdVo()).finish();

    }

    @Override
    public void importsdztkd(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcSdgcZtkdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcSdgcZtkdVo>(JjgFbgcSdgcZtkdVo.class) {
                                @Override
                                public void handle(List<JjgFbgcSdgcZtkdVo> dataList) {
                                    for(JjgFbgcSdgcZtkdVo ztkdVo: dataList)
                                    {
                                        JjgFbgcSdgcZtkd fbgcSdgcZtkd = new JjgFbgcSdgcZtkd();
                                        BeanUtils.copyProperties(ztkdVo,fbgcSdgcZtkd);
                                        fbgcSdgcZtkd.setCreatetime(new Date());
                                        fbgcSdgcZtkd.setProname(commonInfoVo.getProname());
                                        fbgcSdgcZtkd.setHtd(commonInfoVo.getHtd());
                                        fbgcSdgcZtkd.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcSdgcZtkdMapper.insert(fbgcSdgcZtkd);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }
}
