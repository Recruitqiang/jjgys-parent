package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLmgcGslqlmhdzxf;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.lmgc.JjgFbgcLmgcGslqlmhdzxfVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcLmgcGslqlmhdzxfMapper;
import glgc.jjgys.system.service.JjgFbgcLmgcGslqlmhdzxfService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.service.ProjectService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
 * @since 2023-04-25
 */
@Service
public class JjgFbgcLmgcGslqlmhdzxfServiceImpl extends ServiceImpl<JjgFbgcLmgcGslqlmhdzxfMapper, JjgFbgcLmgcGslqlmhdzxf> implements JjgFbgcLmgcGslqlmhdzxfService {

    @Autowired
    private JjgFbgcLmgcGslqlmhdzxfMapper jjgFbgcLmgcGslqlmhdzxfMapper;

    @Autowired
    private ProjectService projectService;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        String proname = commonInfoVo.getProname();
        Integer level = projectService.getlevel(proname);
        if (level == 0){
            DBtoExcelLqData(commonInfoVo);
        }else {
            DBtoExcelPtData(commonInfoVo,level);
        }

    }

    /**
     * 普通公路沥青路面厚度
     * @param commonInfoVo
     * @param level
     * @throws IOException
     * @throws ParseException
     */
    private void DBtoExcelPtData(CommonInfoVo commonInfoVo,int level) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获取数据
        QueryWrapper<JjgFbgcLmgcGslqlmhdzxf> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("zh");
        List<JjgFbgcLmgcGslqlmhdzxf> data = jjgFbgcLmgcGslqlmhdzxfMapper.selectList(wrapper);
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"22沥青路面厚度-钻芯法.xlsx");
        if (data == null || data.size()==0){
            return;
        }else {
            List<JjgFbgcLmgcGslqlmhdzxf> zxzfdata = jjgFbgcLmgcGslqlmhdzxfMapper.selectzxzf(proname, htd, fbgc);
            List<JjgFbgcLmgcGslqlmhdzxf> zxyfdata = jjgFbgcLmgcGslqlmhdzxfMapper.selectzxyf(proname, htd, fbgc);
            List<JjgFbgcLmgcGslqlmhdzxf> sdzfdata = jjgFbgcLmgcGslqlmhdzxfMapper.selectsdzf(proname, htd, fbgc);
            List<JjgFbgcLmgcGslqlmhdzxf> sdyfdata = jjgFbgcLmgcGslqlmhdzxfMapper.selectsdyf(proname, htd, fbgc);
            List<JjgFbgcLmgcGslqlmhdzxf> qlzfdata = jjgFbgcLmgcGslqlmhdzxfMapper.selectqlzf(proname, htd, fbgc);
            List<JjgFbgcLmgcGslqlmhdzxf> qlyfdata = jjgFbgcLmgcGslqlmhdzxfMapper.selectqlyf(proname, htd, fbgc);
            List<JjgFbgcLmgcGslqlmhdzxf> zddata = jjgFbgcLmgcGslqlmhdzxfMapper.selectzd(proname, htd, fbgc);
            List<JjgFbgcLmgcGslqlmhdzxf> ljxdata = jjgFbgcLmgcGslqlmhdzxfMapper.selectljx(proname, htd, fbgc);
            List<JjgFbgcLmgcGslqlmhdzxf> ljxqdata = jjgFbgcLmgcGslqlmhdzxfMapper.selectljxq(proname, htd, fbgc);
            List<JjgFbgcLmgcGslqlmhdzxf> ljxsddata = jjgFbgcLmgcGslqlmhdzxfMapper.selectljxsd(proname, htd, fbgc);
            File fdir = new File(filepath + File.separator + proname + File.separator + htd);
            if (!fdir.exists()) {
                //创建文件根目录
                fdir.mkdirs();
            }
            File directory = new File("service-system/src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String path = reportPath + File.separator + "普通公路沥青路面厚度-钻芯法.xlsx";
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            //路面左幅
            ptlmhdzxf(wb, zxzfdata, "路面左幅");
            //主线右幅
            ptlmhdzxf(wb, zxyfdata, "路面右幅");
            //隧道左幅
            ptlmhdzxf(wb, sdzfdata, "隧道左幅");
            //隧道右幅
            ptlmhdzxf(wb, sdyfdata, "隧道右幅");
            //桥面左幅
            ptlmhdzxf(wb, qlzfdata, "桥面左幅");
            //桥面右幅
            ptlmhdzxf(wb, qlyfdata, "桥面右幅");
            //路面匝道
            ptlmhdzxf(wb, zddata, "路面匝道");
            //连接线
            ptlmhdzxf(wb, ljxdata, "连接线");
            //连接线桥
            ptlmhdzxf(wb, ljxqdata, "连接线桥");
            //连接线隧道
            ptlmhdzxf(wb, ljxsddata, "连接线隧道");

            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                if (shouldBeCalculate(wb.getSheetAt(j))) {
                    calculateThicknessSheet(wb.getSheetAt(j),level);
                    getTunnelTotalpt(wb.getSheetAt(j));
                }
            }
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {   //表内公式  计算 显示结果
                JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
            }

            JjgFbgcCommonUtils.deletezxfEmptySheets(wb);
            FileOutputStream fileOut = new FileOutputStream(f);
            wb.write(fileOut);
            fileOut.flush();
            fileOut.close();
            out.close();
            wb.close();

        }

    }

    /**
     *
     * @param sheet
     */
    private void getTunnelTotalpt(XSSFSheet sheet) {
        XSSFRow row = null;
        boolean flag = false;
        XSSFRow startrow = null;
        XSSFRow endrow = null;
        String name = "";
        for (int i = sheet.getFirstRowNum(); i <= sheet
                .getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if(flag && row.getCell(1) != null && !"".equals(row.getCell(1).toString())){
                if(!name.equals(row.getCell(1).toString().substring(0, row.getCell(1).toString().indexOf("K"))) && !"".equals(name)){
                    endrow = sheet.getRow(i-1);
                    startrow.createCell(12).setCellFormula("COUNT("
                            +startrow.getCell(8).getReference()+":"
                            +endrow.getCell(8).getReference()+")");//=COUNT(I7:I36)
                    startrow.createCell(13).setCellFormula("COUNTIF("
                            +startrow.getCell(10).getReference()+":"
                            +endrow.getCell(10).getReference()+",\"√\")");//=COUNTIF(K7:K36,"√")
                    startrow.createCell(14).setCellFormula(startrow.getCell(13).getReference()+"*100/"
                            +startrow.getCell(12).getReference());
                    name = row.getCell(1).toString().substring(0, row.getCell(1).toString().indexOf("K"));
                    //System.out.println("name2 = "+name);
                    startrow = row;
                }
                if("".equals(name)){
                    /*
                     * 隧道要分左右幅统计，但渗水系数没有分开统计，所以要根据桩号的z/y来判断
                     */
                    name = row.getCell(1).toString().substring(0, row.getCell(1).toString().indexOf("K"));
                    //System.out.println("name1 = "+name);
                    startrow = row;
                }

            }
            if ("序号".equals(row.getCell(0).toString())) {
                flag = true;
                i++;
            }
            if ("评定".equals(row.getCell(0).toString()) && startrow != null && endrow!=null) {
                //System.out.println("name3 = "+name);
                endrow = sheet.getRow(i-1);
                startrow.createCell(12).setCellFormula("COUNT("
                        +startrow.getCell(8).getReference()+":"
                        +endrow.getCell(8).getReference()+")");//=COUNT(I7:I36)
                startrow.createCell(13).setCellFormula("COUNTIF("
                        +startrow.getCell(10).getReference()+":"
                        +endrow.getCell(10).getReference()+",\"√\")");//=COUNTIF(K7:K36,"√")
                startrow.createCell(14).setCellFormula(startrow.getCell(13).getReference()+"*100/"
                        +startrow.getCell(12).getReference());
                break;
            }
        }
    }

    /**
     *
     * @param sheet
     * @param level
     */
    private void calculateThicknessSheet(XSSFSheet sheet,int level) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        boolean flag = false;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            // 下一张表
            if (!"".equals(row.getCell(0).toString()) && row.getCell(0).toString().contains("质量鉴定表") && flag) {
                int record = rowstart.getRowNum();
                while(record <= rowend.getRowNum()){
                    sheet.getRow(record).getCell(10).setCellFormula("IF($I$31=\"合格\",IF("
                            +sheet.getRow(record).getCell(8).getReference()+">=("
                            +sheet.getRow(record).getCell(9).getReference()+"+$C$4),\"√\",\"\"),\"\")");//K=IF($I$31="合格",IF(I7>=(J7+$C$4),"√",""),"")
                    sheet.getRow(record).getCell(11).setCellFormula("IF("
                            +sheet.getRow(record).getCell(8).getReference()+"=\"\",\"\",IF("
                            +sheet.getRow(record).getCell(10).getReference()+"=\"\",\"×\",\"\"))");//L=IF(I7="","",IF(K7="","×",""))
                    record ++;
                }
                flag = false;
            }
            if(flag && !"".equals(row.getCell(1).toString())){
                rowend = row;
                row.getCell(8).setCellFormula("AVERAGE("+row.getCell(4).getReference()+":"+row.getCell(7).getReference()+")");//I=AVERAGE(E7:H7)
                row.createCell(15).setCellFormula(row.getCell(8).getReference());//P7=MAX(E7:H7)
                row.createCell(16).setCellFormula(row.getCell(8).getReference());
            }
            if ("序号".equals(row.getCell(0).toString())) {
                rowstart = sheet.getRow(i+2);
                rowend = rowstart;
                i += 1;
                flag = true;
                continue;
            }
            if ("评定".equals(row.getCell(0).toString())) {
                row.getCell(3).setCellFormula("AVERAGE("
                        +rowstart.getCell(8).getReference()+":"
                        +rowend.getCell(8).getReference()+")");//D30=AVERAGE(I7:I29)
                row.getCell(6).setCellFormula("STDEV("
                        +rowstart.getCell(8).getReference()+":"
                        +rowend.getCell(8).getReference()+")");//G30=STDEV(I7:I29)

                sheet.getRow(i+2).getCell(3).setCellFormula("COUNT("
                        +rowstart.getCell(8).getReference()+":"
                        +rowend.getCell(8).getReference()+")");//D32=COUNT(I7:I29)

                row.getCell(10).setCellFormula("ROUND("
                        +row.getCell(3).getReference()+"-("
                        +row.getCell(6).getReference()+"*VLOOKUP("
                        +sheet.getRow(i+2).getCell(3).getReference()+",保证率系数!A:D,"+(3+level)+")),1)");//K30=ROUND(D30-(G30*VLOOKUP(D32,保证率系数!A:D,3,)),1)

                //小侯的计算方式
                sheet.getRow(i+1).getCell(4).setCellValue(-5);

                sheet.getRow(i+1).getCell(8).setCellFormula("IF("
                        +row.getCell(10).getReference()+">($C$4+"
                        +sheet.getRow(i+1).getCell(4).getReference()
                        +"),\"合格\",\"不合格\")");//I31=IF(K30>($C$4+E31),"合格","不合格")


                sheet.getRow(i+2).getCell(6).setCellFormula("COUNTIF("
                        +rowstart.getCell(10).getReference()+":"
                        +rowend.getCell(10).getReference()+",\"√\")");//G32=COUNTIF(K7:K29,"√")

                sheet.getRow(i+2).getCell(10).setCellFormula(
                        sheet.getRow(i+2).getCell(6).getReference()+"/"
                                +sheet.getRow(i+2).getCell(3).getReference()+"*100");//K32=G32/D32*100
                if(sheet.getRow(i-1).getCell(15) == null){
                    sheet.getRow(i-1).createCell(15);
                }
                if(sheet.getRow(i-1).getCell(16) == null){
                    sheet.getRow(i-1).createCell(16);
                }
                if(sheet.getRow(6).getCell(15) == null){
                    sheet.getRow(6).createCell(15);
                }
                if(sheet.getRow(6).getCell(16) == null){
                    sheet.getRow(6).createCell(16);
                }
                row.createCell(15).setCellFormula("MAX("+sheet.getRow(6).getCell(15).getReference()+":"
                        +sheet.getRow(i-1).getCell(15).getReference()+")");//P7=MAX(p7:p29)

                row.createCell(16).setCellFormula("MIN("+sheet.getRow(6).getCell(16).getReference()+":"
                        +sheet.getRow(i-1).getCell(16).getReference()+")");
                i+=2;
            }
        }
        int record = rowstart.getRowNum();
        while(record <= rowend.getRowNum()){
            sheet.getRow(record).getCell(10).setCellFormula("IF($I$31=\"合格\",IF("
                    +sheet.getRow(record).getCell(8).getReference()+">=("
                    +sheet.getRow(record).getCell(9).getReference()+"+$C$4),\"√\",\"\"),\"\")");//K=IF($I$31="合格",IF(I7>=(J7+$C$4),"√",""),"")
            sheet.getRow(record).getCell(11).setCellFormula("IF("
                    +sheet.getRow(record).getCell(8).getReference()+"=\"\",\"\",IF("
                    +sheet.getRow(record).getCell(10).getReference()+"=\"\",\"×\",\"\"))");//L=IF(I7="","",IF(K7="","×",""))
            record ++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     * @param sheetname
     * @throws ParseException
     */
    private void ptlmhdzxf(XSSFWorkbook wb, List<JjgFbgcLmgcGslqlmhdzxf> data, String sheetname) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        if(data.size() > 0) {
            createPtTable(wb, getpttableNum(data.size()), sheetname);
            XSSFSheet sheet = wb.getSheet(sheetname);
            String type = data.get(0).getLx();
            int index = 6;
            sheet.getRow(1).getCell(2).setCellValue(data.get(0).getProname());
            sheet.getRow(1).getCell(7).setCellValue(data.get(0).getHtd());
            sheet.getRow(2).getCell(2).setCellValue(data.get(0).getFbgc());
            sheet.getRow(2).getCell(7).setCellValue("路面面层（混凝土路面）");
            if(type.equals("隧")){
                sheet.getRow(2).getCell(7).setCellValue("隧道路面");
            }
            else if(type.equals("匝道"))
            {
                sheet.getRow(2).getCell(7).setCellValue("匝道路面");
            }
            else if(type.equals("连接线"))
            {
                sheet.getRow(2).getCell(7).setCellValue("连接线");
            } else
            {
                sheet.getRow(2).getCell(7).setCellValue("路面面层（主线）");
            }
            String date = simpleDateFormat.format(data.get(0).getJcsj());
            for(int i =1; i < data.size(); i++){
                date = JjgFbgcCommonUtils.getLastDate(date, simpleDateFormat.format(data.get(i).getJcsj()));
            }
            sheet.getRow(3).getCell(2).setCellValue(Double.parseDouble(data.get(0).getZhdsjz()));
            sheet.getRow(3).getCell(7).setCellValue(date);
            for(int i =0; i < data.size(); i++){
                sheet.getRow(index+i).getCell(0).setCellValue(i+1);
                sheet.addMergedRegion(new CellRangeAddress(i+index, i+index, 1, 2));
                sheet.getRow(index+i).getCell(1).setCellValue(data.get(i).getLx()+data.get(i).getZh());
                String zyfname;
                if (data.get(i).getZh().substring(0,1).equals("Z")){
                    zyfname = "左幅";
                }else if (data.get(i).getZh().substring(0,1).equals("Y")){
                    zyfname = "右幅";
                }else {
                    zyfname = "";
                }
                sheet.getRow(index+i).getCell(3).setCellValue(zyfname);
                sheet.getRow(index+i).getCell(4).setCellValue(Double.parseDouble(data.get(i).getZhdcz1()));
                sheet.getRow(index+i).getCell(5).setCellValue(Double.parseDouble(data.get(i).getZhdcz2()));
                sheet.getRow(index+i).getCell(6).setCellValue(Double.parseDouble(data.get(i).getZhdcz3()));
                sheet.getRow(index+i).getCell(7).setCellValue(Double.parseDouble(data.get(i).getZhdcz4()));
                sheet.getRow(index+i).getCell(9).setCellValue(Double.parseDouble(data.get(i).getZhdsjz())<=60?-10:Double.parseDouble(data.get(i).getZhdsjz())*-0.15);
            }

        }
    }

    /**
     *
     * @param wb
     * @param tableNum
     * @param sheetname
     */
    private void createPtTable(XSSFWorkbook wb, int tableNum, String sheetname) {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            if(i < record-1){
                RowCopy.copyRows(wb, sheetname, sheetname, 6, 31, (i - 1) * 26 + 32);
            }
            else{
                RowCopy.copyRows(wb, sheetname, sheetname, 6, 29, (i - 1) * 26 + 32);
            }
        }
        if(record == 1){
            wb.getSheet(sheetname).shiftRows(30, 32, -1);
        }
        RowCopy.copyRows(wb, "source", sheetname, 0, 2,(record) * 26 + 3);
        wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 11, 0,(record) * 26 + 5);
    }

    /**
     *
     * @param size
     * @return
     */
    private int getpttableNum(int size) {
        return size%26 <= 24 ? size/26+1 : size/26+2;
    }

    /**
     * 高速沥青路面厚度
     * @param commonInfoVo
     */
    private void DBtoExcelLqData(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获取数据
        QueryWrapper<JjgFbgcLmgcGslqlmhdzxf> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("zh");
        List<JjgFbgcLmgcGslqlmhdzxf> data = jjgFbgcLmgcGslqlmhdzxfMapper.selectList(wrapper);
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"22沥青路面厚度-钻芯法.xlsx");
        if (data == null || data.size()==0){
            return;
        }else {
            List<JjgFbgcLmgcGslqlmhdzxf> zxzfdata = jjgFbgcLmgcGslqlmhdzxfMapper.selectzxzf(proname, htd, fbgc);
            List<JjgFbgcLmgcGslqlmhdzxf> zxyfdata = jjgFbgcLmgcGslqlmhdzxfMapper.selectzxyf(proname, htd, fbgc);
            List<JjgFbgcLmgcGslqlmhdzxf> sdzfdata = jjgFbgcLmgcGslqlmhdzxfMapper.selectsdzf(proname, htd, fbgc);
            List<JjgFbgcLmgcGslqlmhdzxf> sdyfdata = jjgFbgcLmgcGslqlmhdzxfMapper.selectsdyf(proname, htd, fbgc);
            List<JjgFbgcLmgcGslqlmhdzxf> qlzfdata = jjgFbgcLmgcGslqlmhdzxfMapper.selectqlzf(proname, htd, fbgc);
            List<JjgFbgcLmgcGslqlmhdzxf> qlyfdata = jjgFbgcLmgcGslqlmhdzxfMapper.selectqlyf(proname, htd, fbgc);
            List<JjgFbgcLmgcGslqlmhdzxf> zddata = jjgFbgcLmgcGslqlmhdzxfMapper.selectzd(proname, htd, fbgc);
            //List<JjgFbgcLmgcGslqlmhdzxf> ljxdata = jjgFbgcLmgcGslqlmhdzxfMapper.selectljx(proname, htd, fbgc);
            File fdir = new File(filepath + File.separator + proname + File.separator + htd);
            if (!fdir.exists()) {
                //创建文件根目录
                fdir.mkdirs();
            }
            File directory = new File("service-system/src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String path = reportPath + File.separator + "沥青路面厚度-钻芯法.xlsx";
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            //路面左幅
            lmhdzxf(wb, zxzfdata, "路面左幅");
            //主线右幅
            lmhdzxf(wb, zxyfdata, "路面右幅");
            //隧道左幅
            lmhdzxf(wb, sdzfdata, "隧道左幅");
            //隧道右幅
            lmhdzxf(wb, sdyfdata, "隧道右幅");
            //桥面左幅
            lmhdzxf(wb, qlzfdata, "桥面左幅");
            //桥面右幅
            lmhdzxf(wb, qlyfdata, "桥面右幅");
            //路面匝道
            lmhdzxf(wb, zddata, "路面匝道");
            //连接线
            //lmhdzxf(wb, ljxdata, "连接线");

            /*if (ljxdata.size()>0){
                f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"22沥青路面厚度-连接线-钻芯法.xlsx");
            }*/

            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                if (shouldBeCalculate(wb.getSheetAt(j))) {
                    calculateCompactionSheet(wb.getSheetAt(j), wb.getSheetName(j));
                }
            }
            getTunnelTotal(wb.getSheet("隧道左幅"));
            getTunnelTotal(wb.getSheet("隧道右幅"));

            for (int j = 0; j < wb.getNumberOfSheets(); j++) {   //表内公式  计算 显示结果
                if(!wb.getSheetAt(j).getSheetName().equals("source"))
                {
                    JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
                }
            }
            JjgFbgcCommonUtils.deleteEmptySheets(wb);
            FileOutputStream fileOut = new FileOutputStream(f);
            wb.write(fileOut);
            fileOut.flush();
            fileOut.close();
            out.close();
            wb.close();
        }

    }

    /**
     *
     * @param sheet
     */
    private void getTunnelTotal(XSSFSheet sheet) {
        XSSFRow row = null;
        boolean flag = false;
        XSSFRow startrow = null;
        XSSFRow endrow = null;
        String name = "";
        for (int i = sheet.getFirstRowNum(); i <= sheet
                .getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if ("评定".equals(row.getCell(0).toString())) {
                if(startrow == null){
                    break;
                }
                //System.out.println("name3 = "+name);
                endrow = sheet.getRow(i-1);
                startrow.createCell(20).setCellFormula("COUNT("
                        +startrow.getCell(6).getReference()+":"
                        +endrow.getCell(6).getReference()+")"
                        +"+COUNT("
                        +startrow.getCell(15).getReference()+":"
                        +endrow.getCell(15).getReference()+")");//=COUNT(G7:G9)+=COUNT(P7:P9)

                startrow.createCell(21).setCellFormula("COUNTIF("
                        +startrow.getCell(9).getReference()+":"
                        +endrow.getCell(9).getReference()+",\"√\")"
                        +"+COUNTIF("
                        +startrow.getCell(18).getReference()+":"
                        +endrow.getCell(18).getReference()+",\"√\")");//=COUNTIF(J7:J9,"√")+=COUNTIF(S7:S9,"√")

                startrow.createCell(22).setCellFormula(startrow.getCell(21).getReference()+"*100/"
                        +startrow.getCell(20).getReference());
                break;
            }
            if(flag && row.getCell(0) != null && !"".equals(row.getCell(0).toString())){
                //System.out.println("row.getCell(0).toString() = "+row.getCell(0).toString());
                if(!name.equals(row.getCell(0).toString().substring(0, row.getCell(0).toString().indexOf("K"))) && !"".equals(name)){
                    endrow = sheet.getRow(i-1);
                    startrow.createCell(20).setCellFormula("COUNT("
                            +startrow.getCell(6).getReference()+":"
                            +endrow.getCell(6).getReference()+")"
                            +"+COUNT("
                            +startrow.getCell(15).getReference()+":"
                            +endrow.getCell(15).getReference()+")");//=COUNT(G7:G9)+=COUNT(P7:P9)

                    startrow.createCell(21).setCellFormula("COUNTIF("
                            +startrow.getCell(9).getReference()+":"
                            +endrow.getCell(9).getReference()+",\"√\")"
                            +"+COUNTIF("
                            +startrow.getCell(18).getReference()+":"
                            +endrow.getCell(18).getReference()+",\"√\")");//=COUNTIF(J7:J9,"√")+=COUNTIF(S7:S9,"√")

                    startrow.createCell(22).setCellFormula(startrow.getCell(21).getReference()+"*100/"
                            +startrow.getCell(20).getReference());
                    name = row.getCell(0).toString().substring(0, row.getCell(0).toString().indexOf("K"));
                    //System.out.println("name2 = "+name);
                    startrow = row;
                }
                if("".equals(name)){
                    /*
                     * 隧道要分左右幅统计，但渗水系数没有分开统计，所以要根据桩号的z/y来判断
                     */
                    name = row.getCell(0).toString().substring(0, row.getCell(0).toString().indexOf("K"));
                    //System.out.println("name1 = "+name);
                    startrow = row;
                }

            }
            if ("桩     号".equals(row.getCell(0).toString())) {
                flag = true;
                i+=2;
            }

        }
    }

    /**
     *
     * @param sheet
     * @param name
     */
    private void calculateCompactionSheet(XSSFSheet sheet, String name) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        boolean flag = false;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            // 下一张表
            if (!"".equals(row.getCell(0).toString()) && row.getCell(0).toString().contains("质量鉴定表") && flag) {
                //System.out.println("下一张表！！！");
                flag = false;
            }
            if ("桩     号".equals(row.getCell(0).toString())) {
                rowstart = sheet.getRow(i+3);
                rowend = rowstart;
                i += 2;
                flag = true;
                continue;
            }
            if ("评定".equals(row.getCell(0).toString())) {
                row.getCell(3).setCellFormula("AVERAGE("
                        +rowstart.getCell(6).getReference()+":"
                        +rowend.getCell(6).getReference()+")");//D23=AVERAGE(G7:G20)

                row.getCell(5).setCellFormula("STDEV("
                        +rowstart.getCell(6).getReference()+":"
                        +rowend.getCell(6).getReference()+")");//F23=STDEV(G7:G20)

                //-----------------
                sheet.getRow(i+1).getCell(5).setCellFormula("COUNT("
                        +rowstart.getCell(6).getReference()+":"
                        +rowend.getCell(6).getReference()+")");//F24=COUNT(G7:G20)

                sheet.getRow(i+1).getCell(14).setCellFormula("COUNT("
                        +rowstart.getCell(15).getReference()+":"
                        +rowend.getCell(15).getReference()+")");//O24=COUNT(P7:P20)
                //------------------

                row.getCell(7).setCellFormula("ROUND("
                        +row.getCell(3).getReference()+"-("
                        +row.getCell(5).getReference()+"*VLOOKUP("
                        +sheet.getRow(i+1).getCell(5).getReference()+",保证率系数!A:D,3)),1)");//H23=ROUND(D23-(F23*VLOOKUP(F24,保证率系数!A:D,3,)),1)

                row.getCell(10).setCellFormula(rowstart.getCell(7).getReference()+"*-0.1");//K23=-H7*0.1

                if(name.contains("桥面") || name.contains("隧道") || name.contains("匝道")|| name.contains("连接线")){
                    row.getCell(12).setCellFormula("AVERAGE("
                            +rowstart.getCell(15).getReference()+":"
                            +rowend.getCell(15).getReference()+")");//M23=AVERAGEIF(P7:P20,"<150")
                }else{
                    row.getCell(12).setCellFormula("AVERAGE("
                            +rowstart.getCell(15).getReference()+":"
                            +rowend.getCell(15).getReference()+")");//M23=AVERAGEIF(P7:P20,">=150")

                }

                row.getCell(14).setCellFormula("ROUND(STDEV("
                        +rowstart.getCell(15).getReference()+":"
                        +rowend.getCell(15).getReference()+"),1)");//O23=STDEV(P7:P20)

                row.getCell(16).setCellFormula("ROUND("
                        +row.getCell(12).getReference()+"-("
                        +row.getCell(14).getReference()+"*VLOOKUP("
                        +sheet.getRow(i+1).getCell(14).getReference()+",保证率系数!A:D,3)),1)");//Q23=ROUND(M23-(O23*VLOOKUP(O24,保证率系数!A:D,3,)),1)

                row.getCell(19).setCellFormula(rowstart.getCell(16).getReference()+"*-0.05");//T23=Q7*-0.08  总厚度

                sheet.getRow(i+1).getCell(3).setCellFormula("IF("
                        +row.getCell(7).getReference()+">=("
                        +rowend.getCell(7).getReference()+"+"
                        +row.getCell(10).getReference()+"),\"合格\",\"\")");//D24=IF(H23>=(H7+K23),"合格","")



                sheet.getRow(i+1).getCell(7).setCellFormula("COUNTIF("
                        +rowstart.getCell(9).getReference()+":"
                        +rowend.getCell(9).getReference()+",\"√\")");//H24=COUNTIF(J7:J20,"√")

                sheet.getRow(i+1).getCell(9).setCellFormula(sheet.getRow(i+1).getCell(7)+"/"
                        +sheet.getRow(i+1).getCell(5)+"*100");//J24=H24/F24*100

                sheet.getRow(i+1).getCell(12).setCellFormula("IF("
                        +row.getCell(16).getReference()+">=("
                        +rowstart.getCell(16).getReference()+"+"
                        +row.getCell(19).getReference()+"),\"合格\",\"\")");//M24=IF(Q23>=(Q7+T23),"合格","")



                sheet.getRow(i+1).getCell(16).setCellFormula("COUNTIF("
                        +rowstart.getCell(18).getReference()+":"
                        +rowend.getCell(18).getReference()+",\"√\")");//Q24=COUNTIF(S7:S20,"√")

                sheet.getRow(i+1).getCell(18).setCellFormula(sheet.getRow(i+1).getCell(16)+"/"
                        +sheet.getRow(i+1).getCell(14)+"*100");//S24=Q24/O24*100
                flag = false;

                //计算最大值最小值
                if(sheet.getRow(i-1).getCell(20+3) == null){
                    sheet.getRow(i-1).createCell(20+3);
                }
                if(sheet.getRow(i-1).getCell(21+3) == null){
                    sheet.getRow(i-1).createCell(21+3);
                }
                if(sheet.getRow(6).getCell(20+3) == null){
                    sheet.getRow(6).createCell(20+3);
                }
                if(sheet.getRow(6).getCell(21+3) == null){
                    sheet.getRow(6).createCell(21+3);
                }
                row.createCell(20+3).setCellFormula("MAX("+sheet.getRow(6).getCell(20+3).getReference()+":"
                        +sheet.getRow(i-1).getCell(20+3).getReference()+")");//P7=MAX(p7:p29)

                row.createCell(21+3).setCellFormula("MIN("+sheet.getRow(6).getCell(21+3).getReference()+":"
                        +sheet.getRow(i-1).getCell(21+3).getReference()+")");

                if(sheet.getRow(i-1).getCell(22+3) == null){
                    sheet.getRow(i-1).createCell(22+3);
                }
                if(sheet.getRow(i-1).getCell(23+3) == null){
                    sheet.getRow(i-1).createCell(23+3);
                }
                if(sheet.getRow(6).getCell(22+3) == null){
                    sheet.getRow(6).createCell(22+3);
                }
                if(sheet.getRow(6).getCell(23+3) == null){
                    sheet.getRow(6).createCell(23+3);
                }
                row.createCell(22+3).setCellFormula("MAX("+sheet.getRow(6).getCell(22+3).getReference()+":"
                        +sheet.getRow(i-1).getCell(22+3).getReference()+")");//P7=MAX(p7:p29)

                row.createCell(23+3).setCellFormula("MIN("+sheet.getRow(6).getCell(23+3).getReference()+":"
                        +sheet.getRow(i-1).getCell(23+3).getReference()+")");

            }
            if(flag && !"".equals(row.getCell(0).toString())){
                rowend = row;
                //System.out.println("计算平均值："+row.getRowNum());
                if(row.getCell(2) != null && !"".equals(row.getCell(2).toString()) && !"-".equals(row.getCell(2).toString())){
                    row.getCell(6).setCellFormula("ROUND(AVERAGE("+row.getCell(2).getReference()+":"+row.getCell(5).getReference()+"),1)");//G=AVERAGE(C7:F7)
                    row.getCell(8).setCellFormula(row.getCell(7).getReference()+"*-0.2");//G=AVERAGE(C7:F7)
                    row.getCell(9).setCellFormula("IF(AND("+row.getCell(6).getReference()+">="
                            +row.getCell(7).getReference()+"+"
                            +row.getCell(8).getReference()+"),\"√\",\"\")");//J=IF(AND(G7>=H7+I7),"√","")
                    row.getCell(10).setCellFormula("IF(OR("+row.getCell(6).getReference()+"<"
                            +row.getCell(7).getReference()+"+"
                            +row.getCell(8).getReference()+"),\"×\",\"\")");//K=IF(OR(G7<32),"×","")
                }
                else{
                    row.getCell(6).setCellValue("-");
                    row.getCell(8).setCellValue("-");
                    row.getCell(9).setCellValue("-");
                    row.getCell(10).setCellValue("-");
                }

                if(row.getCell(11) != null && !"".equals(row.getCell(11).toString()) && !"-".equals(row.getCell(11).toString())){
                    row.getCell(15).setCellFormula("ROUND(AVERAGE("+row.getCell(11).getReference()+":"+row.getCell(14).getReference()+"),1)");//P=AVERAGE(L7:O7)
                    row.getCell(17).setCellFormula(row.getCell(16).getReference()+"*-0.1");
                    row.getCell(18).setCellFormula("IF("+row.getCell(15).getReference()+">="
                            +row.getCell(16).getReference()+"+"
                            +row.getCell(17).getReference()+",\"√\",\"\")");//S=IF(P7>=Q7+R7,"√","")
                    row.getCell(19).setCellFormula("IF("+row.getCell(15).getReference()+"<"
                            +row.getCell(16).getReference()+"+"
                            +row.getCell(17).getReference()+",\"×\",\"\")");//T=IF(P7<Q7+R7,"×","")
                }
                else{
                    row.getCell(15).setCellValue("-");
                    row.getCell(17).setCellValue("-");
                    row.getCell(18).setCellValue("-");
                    row.getCell(19).setCellValue("-");
                }
                row.createCell(20+3).setCellFormula(row.getCell(6).getReference());//P7=MAX(E7:H7)
                row.createCell(21+3).setCellFormula(row.getCell(6).getReference());
                row.createCell(22+3).setCellFormula(row.getCell(15).getReference());//P7=MAX(E7:H7)
                row.createCell(23+3).setCellFormula(row.getCell(15).getReference());
            }
        }
    }

    /**
     *
     * @param sheet
     * @return
     */
    private boolean shouldBeCalculate(XSSFSheet sheet) {
        String title = null;
        title = sheet.getRow(0).getCell(0).getStringCellValue();
        if (title.contains("鉴定表")) {
            return true;
        }
        return false;
    }

    /**
     *
     * @param wb
     * @param data
     * @param sheetname
     */
    private void lmhdzxf(XSSFWorkbook wb, List<JjgFbgcLmgcGslqlmhdzxf> data, String sheetname) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        if(data.size() > 0) {
            createTable(wb, gettableNum(data.size()), sheetname);
            XSSFSheet sheet = wb.getSheet(sheetname);
            String type = data.get(0).getLx();
            int index = 6;
            sheet.getRow(1).getCell(1).setCellValue(data.get(0).getProname());
            sheet.getRow(1).getCell(17).setCellValue(data.get(0).getHtd());
            sheet.getRow(2).getCell(1).setCellValue(data.get(0).getFbgc());
            if(type.equals("隧")){
                sheet.getRow(2).getCell(9).setCellValue("隧道路面");
            }
            else if(type.equals("匝道"))
            {
                sheet.getRow(2).getCell(9).setCellValue("匝道路面");
            }
            else if(type.contains("连接线"))
            {
                sheet.getRow(2).getCell(9).setCellValue("连接线");
            } else
            {
                sheet.getRow(2).getCell(9).setCellValue("路面面层（主线）");
            }

            String date = simpleDateFormat.format(data.get(0).getJcsj());
            for(int i =1; i < data.size(); i++){
                date = JjgFbgcCommonUtils.getLastDate(date, simpleDateFormat.format(data.get(i).getJcsj()));
            }
            sheet.getRow(2).getCell(17).setCellValue(date);
            for(int i =0; i < data.size(); i++){
                sheet.getRow(index+i).getCell(0).setCellValue(data.get(i).getLx()+data.get(i).getZh());
                String zyfname;
                if (data.get(i).getZh().substring(0,1).equals("Z")){
                    zyfname = "左幅";
                }else if (data.get(i).getZh().substring(0,1).equals("Y")){
                    zyfname = "右幅";
                }else {
                    zyfname = "";
                }
                sheet.getRow(index+i).getCell(1).setCellValue(zyfname);
                sheet.getRow(index+i).getCell(2).setCellValue(Double.parseDouble(data.get(i).getSmccz1()));
                sheet.getRow(index+i).getCell(3).setCellValue(Double.parseDouble(data.get(i).getSmccz2()));
                sheet.getRow(index+i).getCell(4).setCellValue(Double.parseDouble(data.get(i).getSmccz3()));
                sheet.getRow(index+i).getCell(5).setCellValue(Double.parseDouble(data.get(i).getSmccz4()));
                sheet.getRow(index+i).getCell(7).setCellValue(Double.parseDouble(data.get(i).getSmcsjz()));

                sheet.getRow(index+i).getCell(11).setCellValue(Double.parseDouble(data.get(i).getZhdcz1()));
                sheet.getRow(index+i).getCell(12).setCellValue(Double.parseDouble(data.get(i).getZhdcz2()));
                sheet.getRow(index+i).getCell(13).setCellValue(Double.parseDouble(data.get(i).getZhdcz3()));
                sheet.getRow(index+i).getCell(14).setCellValue(Double.parseDouble(data.get(i).getZhdcz4()));
                sheet.getRow(index+i).getCell(16).setCellValue(Double.parseDouble(data.get(i).getZhdsjz()));
            }
        }


    }

    /**
     *
     * @param wb
     * @param tableNum
     * @param sheetname
     */
    private void createTable(XSSFWorkbook wb, int tableNum, String sheetname) {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            if(i < record-1){
                RowCopy.copyRows(wb, sheetname, sheetname, 6, 23, (i - 1) * 18 + 24);
            }
            else{
                RowCopy.copyRows(wb, sheetname, sheetname, 6, 22, (i - 1) * 18 + 24);
            }
        }
        if(record == 1){
            wb.getSheet(sheetname).shiftRows(23, 24, -1);
        }
        RowCopy.copyRows(wb, "source", sheetname, 0, 1,(record) * 18 + 4);
        wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 19, 0,(record) * 18 + 5);
    }

    /**
     *
     * @param size
     * @return
     */
    private int gettableNum(int size) {
        return size%18 <= 16 ? size/18+1 : size/18+2;
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();

        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "22沥青路面厚度-钻芯法.xlsx");
        if (!f.exists()) {
            return null;
        } else {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
            List<Map<String, Object> > jgmap = new ArrayList<>();
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                if (!wb.isSheetHidden(wb.getSheetIndex(wb.getSheetAt(j)))) {
                    XSSFSheet slSheet = wb.getSheetAt(j);
                    XSSFCell xmname = slSheet.getRow(1).getCell(1);//项目名
                    XSSFCell htdname = slSheet.getRow(1).getCell(17);//合同段名
                    Map map = new HashMap();

                    if (proname.equals(xmname.toString()) && htd.equals(htdname.toString())) {
                        //获取到最后一行
                        int lastRowNum = slSheet.getLastRowNum();
                        slSheet.getRow(lastRowNum).getCell(5).setCellType(CellType.STRING);//总点数
                        slSheet.getRow(lastRowNum).getCell(7).setCellType(CellType.STRING);//合格点数
                        slSheet.getRow(lastRowNum).getCell(9).setCellType(CellType.STRING);//合格率

                        slSheet.getRow(lastRowNum).getCell(14).setCellType(CellType.STRING);//总点数
                        slSheet.getRow(lastRowNum).getCell(16).setCellType(CellType.STRING);//合格点数
                        slSheet.getRow(lastRowNum).getCell(18).setCellType(CellType.STRING);//合格率

                        slSheet.getRow(6).getCell(7).setCellType(CellType.STRING);
                        slSheet.getRow(6).getCell(16).setCellType(CellType.STRING);

                        slSheet.getRow(lastRowNum-1).getCell(7).setCellType(CellType.STRING);
                        slSheet.getRow(lastRowNum-1).getCell(16).setCellType(CellType.STRING);

                        slSheet.getRow(lastRowNum-1).getCell(23).setCellType(CellType.STRING);
                        slSheet.getRow(lastRowNum-1).getCell(24).setCellType(CellType.STRING);
                        slSheet.getRow(lastRowNum-1).getCell(25).setCellType(CellType.STRING);
                        slSheet.getRow(lastRowNum-1).getCell(26).setCellType(CellType.STRING);


                        double zds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(5).getStringCellValue());
                        double hgds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(7).getStringCellValue());
                        double hgl = Double.valueOf(slSheet.getRow(lastRowNum).getCell(9).getStringCellValue());

                        double zds1 = Double.valueOf(slSheet.getRow(lastRowNum).getCell(14).getStringCellValue());
                        double hgds1 = Double.valueOf(slSheet.getRow(lastRowNum).getCell(16).getStringCellValue());
                        double hgl1 = Double.valueOf(slSheet.getRow(lastRowNum).getCell(18).getStringCellValue());
                        String zdsz1 = decf.format(zds);
                        String hgdsz1 = decf.format(hgds);
                        String hglz1 = df.format(hgl);
                        map.put("路面类型", wb.getSheetName(j));
                        map.put("上面层厚度检测点数", zdsz1);
                        map.put("上面层厚度合格点数", hgdsz1);
                        map.put("上面层厚度合格率", hglz1);
                        map.put("上面层设计值", slSheet.getRow(6).getCell(7).getStringCellValue());
                        map.put("上面层代表值", slSheet.getRow(lastRowNum-1).getCell(7).getStringCellValue());

                        map.put("上面层平均值最大值", slSheet.getRow(lastRowNum-1).getCell(23).getStringCellValue());
                        map.put("上面层平均值最小值", slSheet.getRow(lastRowNum-1).getCell(24).getStringCellValue());

                        map.put("总厚度检测点数", decf.format(zds1));
                        map.put("总厚度合格点数", decf.format(hgds1));
                        map.put("总厚度合格率", df.format(hgl1));
                        map.put("总厚度设计值", slSheet.getRow(6).getCell(16).getStringCellValue());
                        map.put("总厚度代表值", slSheet.getRow(lastRowNum-1).getCell(16).getStringCellValue());

                        map.put("总厚度平均值最大值", slSheet.getRow(lastRowNum-1).getCell(25).getStringCellValue());
                        map.put("总厚度平均值最小值", slSheet.getRow(lastRowNum-1).getCell(26).getStringCellValue());
                    }
                    jgmap.add(map);

                }
            }
            return jgmap;
        }
    }

    @Override
    public void exportlqlmhd(HttpServletResponse response) {
        String fileName = "09高速沥青路面厚度钻芯法实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcLmgcGslqlmhdzxfVo()).finish();

    }

    @Override
    public void importLqlmhd(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcLmgcGslqlmhdzxfVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcLmgcGslqlmhdzxfVo>(JjgFbgcLmgcGslqlmhdzxfVo.class) {
                                @Override
                                public void handle(List<JjgFbgcLmgcGslqlmhdzxfVo> dataList) {
                                    for(JjgFbgcLmgcGslqlmhdzxfVo gslqlmhdzxfVo: dataList)
                                    {
                                        JjgFbgcLmgcGslqlmhdzxf gslqlmhdzxf = new JjgFbgcLmgcGslqlmhdzxf();
                                        BeanUtils.copyProperties(gslqlmhdzxfVo,gslqlmhdzxf);
                                        gslqlmhdzxf.setCreatetime(new Date());
                                        gslqlmhdzxf.setProname(commonInfoVo.getProname());
                                        gslqlmhdzxf.setHtd(commonInfoVo.getHtd());
                                        gslqlmhdzxf.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcLmgcGslqlmhdzxfMapper.insert(gslqlmhdzxf);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }
}
