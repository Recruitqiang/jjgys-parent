package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLjgcHdgqd;
import glgc.jjgys.model.project.JjgFbgcLjgcLjbp;
import glgc.jjgys.model.project.JjgFbgcLjgcLjtsfysdHt;
import glgc.jjgys.model.project.JjgFbgcLjgcLjtsfysdSl;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.ljgc.JjgFbgcLjgcLjbpVo;
import glgc.jjgys.model.projectvo.ljgc.JjgFbgcLjgcLjcjVo;
import glgc.jjgys.model.projectvo.ljgc.JjgFbgcLjgcLjtsfysdHtVo;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcLjgcLjtsfysdHtMapper;
import glgc.jjgys.system.service.JjgFbgcLjgcLjtsfysdHtService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.service.JjgFbgcLjgcLjtsfysdSlService;
import glgc.jjgys.system.service.ProjectService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import io.swagger.annotations.Authorization;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import javax.swing.*;
import java.io.*;
import java.lang.reflect.Field;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * <p>
 *  路基土石方服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-02-21
 */
@Service
public class JjgFbgcLjgcLjtsfysdHtServiceImpl extends ServiceImpl<JjgFbgcLjgcLjtsfysdHtMapper, JjgFbgcLjgcLjtsfysdHt> implements JjgFbgcLjgcLjtsfysdHtService {


    @Autowired
    private JjgFbgcLjgcLjtsfysdHtMapper jjgFbgcLjgcLjtsfysdHtMapper;
    @Autowired
    private JjgFbgcLjgcLjtsfysdSlService jjgFbgcLjgcLjtsfysdSlService;
    @Autowired
    private ProjectService projectService;

    private static XSSFWorkbook xwb = null;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    private LinkedHashMap<String,ArrayList<String>> evaluateDataht = new LinkedHashMap<>();

    private int index = 1;


    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException {
        //XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = "路基土石方";
        //获取数据
        QueryWrapper<JjgFbgcLjgcLjtsfysdHt> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("xh","qyzhjwz");
        List<JjgFbgcLjgcLjtsfysdHt> htdata = jjgFbgcLjgcLjtsfysdHtMapper.selectList(wrapper);
        List<JjgFbgcLjgcLjtsfysdSl> sldata = jjgFbgcLjgcLjtsfysdSlService.getdata(proname, htd, fbgc);
        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"01路基土石方压实度.xlsx");
        //健壮性判断如果没有数据返回"请导入数据"
        if (htdata.size() <= 0 && sldata.size() <= 0){
            return;
        }else {
            //存放鉴定表的目录
            File fdir = new File(filepath+File.separator+proname+File.separator+htd);
            if(!fdir.exists()){
                //创建文件根目录
                fdir.mkdirs();
            }
            File directory = new File("service-system/src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String ss = "路基压实度.xlsx";
            String path =reportPath+File.separator+ss;
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            xwb = new XSSFWorkbook(out);
            //xwb.setSheetHidden(xwb.getSheetIndex("保证率系数"), true);
            //沙砾
            LinkedHashMap<String, ArrayList<String>> EvalSLdata = jjgFbgcLjgcLjtsfysdSlService.writeAndGetData(xwb,proname,htd,fbgc);
            //灰土
            List<String> numsList = jjgFbgcLjgcLjtsfysdHtMapper.selectNums(proname,htd,fbgc);
            createTable(gettableNum(numsList));
            DBtoExcel(htdata,proname,htd);
            LinkedHashMap<String, ArrayList<String>> EvalHTdata = evaluateDataht;
            //根据鉴定表的数据生成评定单元
            if(dataToExcel(xwb,EvalSLdata,EvalHTdata,proname,htd)){
                //删除空的sheet
                JjgFbgcCommonUtils.deleteEmptySheets(xwb);
                if(xwb.getSheet("压实度单点(灰土)") == null || xwb.getSheet("压实度单点(灰土)").getRow(1) == null
                        || xwb.getSheet("压实度单点(灰土)").getRow(1).getCell(1) == null || "".equals(xwb.getSheet("压实度单点(灰土)").getRow(1).getCell(1).toString())){

                    xwb.setSheetHidden(xwb.getSheetIndex("评定单元(2)"), 2);//把sheet工作表都给隐藏掉,0=显示 1=隐藏 2=非常隐秘(在excel里就看不见了，等于删除)
                    xwb.setSheetHidden(xwb.getSheetIndex("压实度单点(灰土)"), 2);
                }
                if(xwb.getSheet("压实度单点(砂砾)") == null || xwb.getSheet("压实度单点(砂砾)").getRow(1) == null
                        || xwb.getSheet("压实度单点(砂砾)").getRow(1).getCell(2) == null || "".equals(xwb.getSheet("压实度单点(砂砾)").getRow(1).getCell(2).toString())){

                    xwb.setSheetHidden(xwb.getSheetIndex("评定单元"), true);
                    xwb.setSheetHidden(xwb.getSheetIndex("压实度单点(砂砾)"), true);
                }
                FileOutputStream fileOut = new FileOutputStream(f);
                xwb.write(fileOut);
                fileOut.flush();
                fileOut.close();
            }
            out.close();
            xwb.close();

        }
    }

    /**
     * 根据鉴定表的数据生成评定单元
     * @param SLData
     * @param HTData
     * @return
     * @throws IOException
     */
    public boolean dataToExcel(XSSFWorkbook wb,LinkedHashMap<String,ArrayList<String>> SLData, LinkedHashMap<String,ArrayList<String>> HTData,String proname,String htd) throws IOException{
        try{
            createTable(wb,SLData , 0);
            createTable(wb,HTData , 1);
            calculateSecondSheet(wb.getSheet("评定单元(2)"), calculateFirstSheet(wb.getSheet("评定单元"),proname,htd),proname,htd);
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                JjgFbgcCommonUtils.updateFormula(wb,wb.getSheet("评定单元(2)"));
            }
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                JjgFbgcCommonUtils.updateFormula(wb,wb.getSheet("评定单元"));
            }
        }catch(Exception e){
            e.printStackTrace();
            return false;
        }
        return true;
    }

    /**
     * 根据oneOrtwo参数判断是“评定单元”还是“评定单元(2)”,然后生成表格
     * @param Data
     * @param oneOrtwo  0表示：“评定单元”；1表示“评定单元(2)”
     * @throws IOException
     */
    public void createTable(XSSFWorkbook wb,LinkedHashMap<String,ArrayList<String>> Data, int oneOrtwo) throws IOException{
        XSSFCellStyle cellstyle = wb.createCellStyle();
        XSSFFont font=wb.createFont();
        font.setFontHeightInPoints((short)10);
        font.setFontName("宋体");
        cellstyle.setFont(font);
        cellstyle.setBorderBottom(BorderStyle.THIN); //下边框
        cellstyle.setBorderLeft(BorderStyle.THIN);//左边框
        cellstyle.setBorderTop(BorderStyle.THIN);//上边框
        cellstyle.setBorderRight(BorderStyle.THIN);//右边框
        cellstyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直
        cellstyle.setAlignment(HorizontalAlignment.CENTER);//水平
        cellstyle.setWrapText(true);//自动换行
        int record = 0;
        int remainder = 0;
        int count = 0;
        for(Map.Entry<String, ArrayList<String>> entry : Data.entrySet()){
            record += entry.getValue().size();
            count += entry.getValue().size();
            if(count > 29){
                count = entry.getValue().size();
                count = count % 29;
            }
        }
        remainder = record % 29;
        //System.out.println("remainder = "+remainder);
        if(oneOrtwo == 0){
            if(remainder <= 26){
                for(int i = 1; i < record/29+1; i++){
                    if(i < record/29){
                        RowCopy.copyRows(wb, "评定单元", "评定单元", 5, 33, (i-1)*29+34);
                    }
                    else{
                        RowCopy.copyRows(wb, "评定单元", "评定单元", 5, 30, (i-1)*29+34);
                    }
                }
                if(record/29 == 0){
                    wb.getSheet("评定单元").shiftRows(31, 33, -1, true , false);
                }
                RowCopy.copyRows(wb, "source", "评定单元", 0, 2, (record/29+1)*29+2);
                wb.getSheet("评定单元").getRow((record/29+1)*29+3).getCell(3).setCellValue(96);//D
                wb.setPrintArea(wb.getSheetIndex("评定单元"), 0, 7, 0, (record/29+1)*29+4);
            }
            else{
                //System.out.println("else");
                for(int i = 1; i < record/29+2; i++){
                    //System.out.println("正在生成表格 -> "+i);
                    if(i < record/29+1){
                        RowCopy.copyRows(wb, "评定单元", "评定单元", 5, 33, (i-1)*29+34);
                    }
                    else{
                        RowCopy.copyRows(wb, "评定单元", "评定单元", 5, 30, (i-1)*29+34);
                    }
                }
                RowCopy.copyRows(wb, "source", "评定单元", 0, 2, (record/29+2)*29+2);
                wb.getSheet("评定单元").getRow((record/29+2)*29+3).getCell(3).setCellValue(96);//D
                wb.setPrintArea(wb.getSheetIndex("评定单元"), 0, 7, 0, (record/29+2)*29+4);
            }
            /*
             * 上面已经生成了表格，此处进行表格数据的填写
             */
            fillData(Data, wb,"评定单元","压实度单点(砂砾)", cellstyle);
            JjgFbgcCommonUtils.updateFormula(wb, wb.getSheet("评定单元"));
        }
        else if(oneOrtwo == 1){
            if(remainder <= 25){
                for(int i = 1; i < record/29+1; i++){
                    //System.out.println("正在生成表格 -> "+i);
                    if(i < record/29){
                        RowCopy.copyRows(wb, "评定单元(2)", "评定单元(2)", 5, 33, (i-1)*29+34);
                    }
                    else{
                        RowCopy.copyRows(wb, "评定单元(2)", "评定单元 (2)", 5, 29, (i-1)*29+34);
                    }
                }
                if(record/29 == 0){
                    wb.getSheet("评定单元(2)").shiftRows(30, 33, -1, true , false);
                }
                RowCopy.copyRows(wb, "source", "评定单元(2)", 3, 6, (record/29+1)*29+1);
                wb.getSheet("评定单元(2)").getRow((record/29+1)*29+2).getCell(3).setCellValue(96);//D
                wb.setPrintArea(wb.getSheetIndex("评定单元(2)"), 0, 7, 0, (record/29+1)*29+4);
            }
            else{
                //System.out.println("else");
                for(int i = 1; i < record/29+2; i++){
                    //System.out.println("正在生成表格 -> "+i);
                    if(i < record/29+1){
                        RowCopy.copyRows(wb, "评定单元(2)", "评定单元(2)", 5, 33, (i-1)*29+34);
                    }
                    else{
                        RowCopy.copyRows(wb, "评定单元(2)", "评定单元(2)", 5, 29, (i-1)*29+34);
                    }
                }
                RowCopy.copyRows(wb, "source", "评定单元(2)", 3, 6, (record/29+2)*29+1);
                wb.getSheet("评定单元(2)").getRow((record/29+2)*29+2).getCell(3).setCellValue(96);//D
                wb.setPrintArea(wb.getSheetIndex("评定单元(2)"), 0, 7, 0, (record/29+2)*29+4);
            }
            /*
             * 上面已经生成了表格，此处进行表格数据的填写
             */
            fillData(Data, wb,"评定单元(2)","压实度单点(灰土)", cellstyle);
            JjgFbgcCommonUtils.updateFormula(wb, wb.getSheet("评定单元(2)"));
        }
        else{
            throw new JjgysException(20001,"生成评定单元错误");
        }
    }

    /**
     * 给两个评定单元填入数据
     * @param Data
     * @param sheetName
     * @param dataSheetName
     * @param cellstyle
     */
    public void fillData(LinkedHashMap<String,ArrayList<String>> Data,XSSFWorkbook wb,String sheetName, String dataSheetName, XSSFCellStyle cellstyle){
        int startrow = 5, endrow = 5;
        for(Map.Entry<String, ArrayList<String>> entry : Data.entrySet()){
            endrow += entry.getValue().size() -1;
            if((endrow-5)/29 > (startrow-5)/29){
                if(startrow < (startrow-5)/29*29+29+5-1){
                    wb.getSheet(sheetName).addMergedRegion(new CellRangeAddress(startrow, (startrow-5)/29*29+29+5-1, 0, 0));
                }
                wb.getSheet(sheetName).getRow(startrow).getCell(0).setCellValue(index);
                wb.getSheet(sheetName).getRow(startrow).getCell(0).setCellStyle(cellstyle);

                if((startrow-5)/29*29+29+5 < endrow){
                    wb.getSheet(sheetName).addMergedRegion(new CellRangeAddress((startrow-5)/29*29+29+5, endrow, 0, 0));
                }
                wb.getSheet(sheetName).getRow((startrow-5)/29*29+29+5).getCell(0).setCellValue(index++);
                wb.getSheet(sheetName).getRow((startrow-5)/29*29+29+5).getCell(0).setCellStyle(cellstyle);
            }
            else{
                if(startrow < endrow){
                    wb.getSheet(sheetName).addMergedRegion(new CellRangeAddress(startrow, endrow, 0, 0));
                }
                wb.getSheet(sheetName).getRow(startrow).getCell(0).setCellValue(index++);
                wb.getSheet(sheetName).getRow(startrow).getCell(0).setCellStyle(cellstyle);
            }

            if((endrow-5)/29 > (startrow-5)/29){
                wb.getSheet(sheetName).addMergedRegion(new CellRangeAddress(startrow, (startrow-5)/29*29+29+5-1, 1, 4));
                String pile = entry.getKey().substring(0, entry.getKey().lastIndexOf("@"));
                wb.getSheet(sheetName).getRow(startrow).getCell(1).setCellValue(pile.substring(0, pile.length()-1));
                wb.getSheet(sheetName).getRow(startrow).getCell(1).setCellStyle(cellstyle);

                wb.getSheet(sheetName).addMergedRegion(new CellRangeAddress((startrow-5)/29*29+29+5, endrow, 1, 4));
                pile = entry.getKey().substring(0, entry.getKey().lastIndexOf("@"));
                wb.getSheet(sheetName).getRow((startrow-5)/29*29+29+5).getCell(1).setCellValue(pile.substring(0, pile.length()-1));
                wb.getSheet(sheetName).getRow((startrow-5)/29*29+29+5).getCell(1).setCellStyle(cellstyle);
            }
            else{
                wb.getSheet(sheetName).addMergedRegion(new CellRangeAddress(startrow, endrow, 1, 4));
                String pile = entry.getKey().substring(0, entry.getKey().lastIndexOf("@"));
                wb.getSheet(sheetName).getRow(startrow).getCell(1).setCellValue(pile.substring(0, pile.length()-1));
                wb.getSheet(sheetName).getRow(startrow).getCell(1).setCellStyle(cellstyle);
            }

            XSSFDataFormat df = wb.createDataFormat();
            XSSFCellStyle cellstyle2 = (XSSFCellStyle) cellstyle.clone();
            cellstyle2.setDataFormat(df.getFormat("0.0"));
            for(String s : entry.getValue()){
                wb.getSheet(sheetName).getRow(startrow).getCell(5).setCellFormula("ROUND('"+dataSheetName+"'"+"!"+s+",1)");
                wb.getSheet(sheetName).getRow(startrow++).getCell(5).setCellStyle(cellstyle2);
            }
            startrow = endrow + 1;
            endrow = startrow;
        }
    }

    /**
     * 计算“评定单元”
     * @param sheet
     * @return
     */
    private String[] calculateFirstSheet(XSSFSheet sheet,String proname,String htd) {
        Integer level = projectService.getlevel(proname);
        String[] usedata = new String[2];
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        boolean flag = false;
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        /*
         * 此处先填写“评定单元”sheet的表头数据
         */
        sheet.getRow(1).getCell(2).setCellValue(proname);
        sheet.getRow(1).getCell(7).setCellValue(htd);
        sheet.getRow(2).getCell(2).setCellValue("路基土石方");
        sheet.getRow(2).getCell(7).setCellValue(simpleDateFormat.format(new Date()));

        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if(flag && "评定（1）".equals(row.getCell(0).toString())){
                rowend = sheet.getRow(i-1);
                //System.out.println(i+" -> ");
                row.getCell(3).setCellFormula("AVERAGE("+rowstart.getCell(5).getReference()+":"
                        +rowend.getCell(5).getReference()+")");//D61=AVERAGE(F6:F57)

                row.getCell(5).setCellFormula("STDEV("+rowstart.getCell(5).getReference()+":"
                        +rowend.getCell(5).getReference()+")");//F61=STDEV(F6:F57)

                sheet.getRow(i+2).getCell(3).setCellFormula("COUNT("+rowstart.getCell(5).getReference()+":"
                        +rowend.getCell(5).getReference()+")");//D63=COUNT(F6:F57)
                usedata[0] = sheet.getRow(i+2).getCell(3).getReference();

                row.createCell(9).setCellFormula("VLOOKUP("
                        +sheet.getRow(i+2).getCell(3).getReference()+",保证率系数!A5:D104,"+(3+level)+",FALSE)");//J61=VLOOKUP(D63,保证率系数!A5:D104,3,FALSE)

                row.getCell(7).setCellFormula(row.getCell(3).getReference()+"-"
                        +row.getCell(5).getReference()+"*"
                        +row.getCell(9).getReference());//H61=D61-F61*J61

                sheet.getRow(i+1).getCell(6).setCellFormula("IF("
                        +row.getCell(7).getReference()+">="
                        +sheet.getRow(i+1).getCell(3).getReference()+",\"合格\",\"不合格\")");//G62=IF(H61>=D62,"合格","不合格")

                for(int k = rowstart.getRowNum(); k <= rowend.getRowNum(); k++){
                    sheet.getRow(k).getCell(6).setCellFormula("IF("
                            +sheet.getRow(k).getCell(5).getReference()+"=\"\",\"\",IF("
                            +sheet.getRow(i+1).getCell(6).getReference()+"=\"合格\",IF("
                            +sheet.getRow(k).getCell(5).getReference()+">="
                            +sheet.getRow(i+1).getCell(3).getReference()+"-2,\"√\",\"\"),\"\"))");//G=IF(F6="","",IF($G$62="合格",IF(F6>=$D$62-2,"√",""),""))

                    sheet.getRow(k).getCell(7).setCellFormula("IF("
                            +sheet.getRow(k).getCell(5).getReference()+"=\"\",\"\",IF("
                            +sheet.getRow(i+1).getCell(6).getReference()+"=\"合格\",IF("
                            +sheet.getRow(k).getCell(5).getReference()+"<"
                            +sheet.getRow(i+1).getCell(3).getReference()+"-2,\"×\",\"\"),\"×\"))");//H=IF(F6="","",IF($G$62="合格",IF(F6<$D$62-2,"×",""),"×"))
                }
                sheet.getRow(i+2).getCell(5).setCellFormula("COUNTIF("+rowstart.getCell(6).getReference()+":"
                        +rowend.getCell(6).getReference()+",\"√\")");//f63=COUNTIF(G6:G57,"√")
                usedata[1] = sheet.getRow(i+2).getCell(5).getReference();

                sheet.getRow(i+2).getCell(7).setCellFormula("ROUND("
                        +sheet.getRow(i+2).getCell(5).getReference()+"/"
                        +sheet.getRow(i+2).getCell(3).getReference()+"*100,1)");//H63=ROUND(F63/D63*100,1)
            }
            if ("序号".equals(row.getCell(0).toString())) {
                rowstart = sheet.getRow(i+2);
                rowend = rowstart;
                i++;
                flag = true;
            }
        }
        return usedata;
    }

    /**
     * 计算“评定单元(2)”
     * @param sheet
     * @param usedata
     */
    private void calculateSecondSheet(XSSFSheet sheet, String[] usedata,String proname,String htd) {
        Integer level = projectService.getlevel(proname);
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        boolean flag = false;
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        /*
         * 此处先填写“评定单元”sheet的表头数据
         */
        sheet.getRow(1).getCell(2).setCellValue(proname);
        sheet.getRow(1).getCell(7).setCellValue(htd);
        sheet.getRow(2).getCell(2).setCellValue("路基土石方");
        sheet.getRow(2).getCell(7).setCellValue(simpleDateFormat.format(new Date()));

        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if(flag && "评定（2）".equals(row.getCell(0).toString())){
                rowend = sheet.getRow(i-1);
                //System.out.println(i+" -> ");
                row.getCell(3).setCellFormula("AVERAGE("+rowstart.getCell(5).getReference()+":"
                        +rowend.getCell(5).getReference()+")");//D61=AVERAGE(F6:F57)

                row.getCell(5).setCellFormula("STDEV("+rowstart.getCell(5).getReference()+":"
                        +rowend.getCell(5).getReference()+")");//F61=STDEV(F6:F57)

                sheet.getRow(i+2).getCell(3).setCellFormula("COUNT("+rowstart.getCell(5).getReference()+":"
                        +rowend.getCell(5).getReference()+")");//D63=COUNT(F6:F57)

                row.createCell(9).setCellFormula("VLOOKUP("
                        +sheet.getRow(i+2).getCell(3).getReference()+",保证率系数!A5:D104,"+(3+level)+",FALSE)");//J61=VLOOKUP(D63,保证率系数!A5:D104,3,FALSE)

                row.getCell(7).setCellFormula(row.getCell(3).getReference()+"-"
                        +row.getCell(5).getReference()+"*"
                        +row.getCell(9).getReference());//H61=D61-F61*J61

                sheet.getRow(i+1).getCell(6).setCellFormula("IF("
                        +row.getCell(7).getReference()+">="
                        +sheet.getRow(i+1).getCell(3).getReference()+",\"合格\",\"不合格\")");//G62=IF(H61>=D62,"合格","不合格")

                for(int k = rowstart.getRowNum(); k <= rowend.getRowNum(); k++){
                    sheet.getRow(k).getCell(6).setCellFormula("IF("
                            +sheet.getRow(k).getCell(5).getReference()+"=\"\",\"\",IF("
                            +sheet.getRow(i+1).getCell(6).getReference()+"=\"合格\",IF("
                            +sheet.getRow(k).getCell(5).getReference()+">="
                            +sheet.getRow(i+1).getCell(3).getReference()+"-2,\"√\",\"\"),\"\"))");//G=IF(F6="","",IF($G$62="合格",IF(F6>=$D$62-2,"√",""),""))

                    sheet.getRow(k).getCell(7).setCellFormula("IF("
                            +sheet.getRow(k).getCell(5).getReference()+"=\"\",\"\",IF("
                            +sheet.getRow(i+1).getCell(6).getReference()+"=\"合格\",IF("
                            +sheet.getRow(k).getCell(5).getReference()+"<"
                            +sheet.getRow(i+1).getCell(3).getReference()+"-2,\"×\",\"\"),\"×\"))");//H=IF(F6="","",IF($G$62="合格",IF(F6<$D$62-2,"×",""),"×"))
                }
                sheet.getRow(i+2).getCell(5).setCellFormula("COUNTIF("+rowstart.getCell(6).getReference()+":"
                        +rowend.getCell(6).getReference()+",\"√\")");//f63=COUNTIF(G6:G57,"√")

                sheet.getRow(i+2).getCell(7).setCellFormula("ROUND("
                        +sheet.getRow(i+2).getCell(5).getReference()+"/"
                        +sheet.getRow(i+2).getCell(3).getReference()+"*100,1)");//H63=ROUND(F63/D63*100,1)

                sheet.getRow(i+3).getCell(3).setCellFormula(
                        sheet.getRow(i+2).getCell(3).getReference()+"+评定单元!"+
                                usedata[0]);//监测点数=D58+评定单元!D63
                sheet.getRow(i+3).getCell(5).setCellFormula(
                        sheet.getRow(i+2).getCell(5).getReference()+"+评定单元!"+
                                usedata[1]);//合格点数=F58+评定单元!F63
                sheet.getRow(i+3).getCell(7).setCellFormula("ROUND("
                        +sheet.getRow(i+3).getCell(5).getReference()+"/"
                        +sheet.getRow(i+3).getCell(3).getReference()+"*100,1)");//H63=ROUND(F63/D63*100,1)
            }
            if ("序号".equals(row.getCell(0).toString())) {
                rowstart = sheet.getRow(i+2);
                rowend = rowstart;
                i++;
                flag = true;
            }
        }
    }


    /**
     * 给压实度单点（灰土）写入数据
     * @param data
     * @param proname
     * @param htd
     * @throws IOException
     */
    public void DBtoExcel(List<JjgFbgcLjgcLjtsfysdHt> data,String proname,String htd) throws IOException{
        if(data.size() > 0){
            DecimalFormat nf = new DecimalFormat(".00");
            XSSFCellStyle cellstyle = JjgFbgcCommonUtils.dBtoExcelUtils(xwb);
            XSSFSheet sheet = xwb.getSheet("压实度单点(灰土)");

            String zh = data.get(0).getZh();
            String xuhao = data.get(0).getXh();
            int i =Integer.valueOf(xuhao);
            System.out.println(i);
            HashMap<Integer, String> map = new HashMap<>();

            for(JjgFbgcLjgcLjtsfysdHt row:data){
                if(xuhao.equals(row.getXh())){
                    if(!zh.equals(row.getZh())){
                        zh = row.getZh();
                    }
                }
                else{
                    if(map.get(Integer.valueOf(xuhao)) == null){
                        map.put(Integer.valueOf(xuhao), zh);
                    }
                    else{
                        map.put(Integer.valueOf(xuhao), map.get(Integer.valueOf(xuhao))+"\n"+zh);
                    }
                    xuhao = row.getXh();
                    zh = row.getZh()+";\n";
                }
            }
            if(!"".equals(zh)){
                if(map.get(Integer.valueOf(xuhao)) == null){
                    map.put(Integer.valueOf(xuhao), zh);
                }
                else{
                    map.put(Integer.valueOf(xuhao), map.get(Integer.valueOf(xuhao))+"\n"+zh);
                }
            }
            int index = 0;
            int tableNum = 0;
            String pileNO = "";
            zh = data.get(0).getZh();
            xuhao = data.get(0).getXh();
            pileNO = map.get(Integer.valueOf(xuhao));
            ArrayList<String> compaction = new ArrayList<String>();
            fillTitleCellData(sheet, tableNum, data.get(0), zh,proname,htd);
            for(JjgFbgcLjgcLjtsfysdHt row:data){
                if(xuhao.equals(row.getXh())){
                    if(index == 10){
                        tableNum ++;
                        fillTitleCellData(sheet, tableNum, data.get(0), row.getZh(),proname,htd);
                        index = 0;
                    }
                    if(index%2 == 0){
                        fillMergeCellData(sheet, tableNum, index, row);
                    }
                    fillCommonCellData(xwb,sheet, tableNum, index, row, cellstyle);
                    if(index%2 == 1){
                        fillCalculateCellData(xwb,sheet, tableNum, index);
                        System.out.println(sheet.getRow(tableNum*29+25).getCell(1+index-1).getReference());
                        compaction.add(sheet.getRow(tableNum*29+25).getCell(1+index-1).getReference());
                    }
                    index ++;
                }
                else{
                    ArrayList<String> temp = new ArrayList<String>();
                    temp = (ArrayList<String>) compaction.clone();
                    evaluateDataht.put(pileNO+"@"+xuhao, temp);
                    compaction.clear();
                    if(index%2 == 1){
                        fillCalculateCellData(xwb,sheet, tableNum, index);
                        compaction.add(sheet.getRow(tableNum*29+25).getCell(1+index-1).getReference());
                    }
                    zh = row.getZh();
                    xuhao = row.getXh();
                    tableNum ++;
                    index = 0;
                    fillTitleCellData(sheet, tableNum, row, zh,proname,htd);
                    fillMergeCellData(sheet, tableNum, index, row);
                    fillCommonCellData(xwb,sheet, tableNum, index, row, cellstyle);
                    index = 1;
                }
            }
            if(index%2 == 1){
                fillCalculateCellData(xwb,sheet, tableNum, index);
                compaction.add(sheet.getRow(tableNum*29+25).getCell(1+index-1).getReference());
            }
            evaluateDataht.put(pileNO+"@"+xuhao, compaction);
        }
    }

    /**
     * 填写下方的需要计算的合并单元格
     * @param sheet
     * @param tableNum
     * @param index
     */
    public void fillCalculateCellData(XSSFWorkbook xwb,XSSFSheet sheet, int tableNum, int index){
        sheet.getRow(tableNum*29+23).getCell(1+index-1).setCellFormula("SUMIF("
                +sheet.getRow(tableNum*29+22).getCell(1+index-1).getReference()+":"
                +sheet.getRow(tableNum*29+22).getCell(1+index).getReference()+",\">0\")/2");//B24=SUMIF(B23:C23,">0")/2

        sheet.getRow(tableNum*29+24).getCell(1+index-1).setCellFormula("IF("
                +sheet.getRow(tableNum*29+23).getCell(1+index-1).getReference()+"=0,\"\","
                +sheet.getRow(tableNum*29+15).getCell(1+index-1).getReference()+"/(1+"
                +sheet.getRow(tableNum*29+23).getCell(1+index-1).getReference()+"/100))");//B25=IF(B24=0," ",B16/(1+B24/100))

        sheet.getRow(tableNum*29+25).getCell(1+index-1).setCellFormula("IF(ISERROR("
                +sheet.getRow(tableNum*29+24).getCell(1+index-1).getReference()+"/"
                +sheet.getRow(tableNum*29+6).getCell(1).getReference()+"*100),\"\","
                +sheet.getRow(tableNum*29+24).getCell((1+index-1)).getReference()+"/"
                +sheet.getRow(tableNum*29+6).getCell((1)).getReference()+"*100)");//B26=IF(ISERROR(B25/B7*100),"",B25/B7*100)

        XSSFFormulaEvaluator evaluate = new XSSFFormulaEvaluator(xwb);

        double value=evaluate.evaluate(sheet.getRow(tableNum*29+25).getCell(1+index-1)).getNumberValue();

        sheet.getRow(tableNum*29+25).getCell(1+index-1).setCellFormula(null);
        sheet.getRow(tableNum*29+25).getCell(1+index-1).setCellValue(value);

        XSSFCellStyle cellstyle2 = xwb.createCellStyle();
        XSSFFont font=xwb.createFont();
        font.setFontHeightInPoints((short)9);
        font.setFontName("宋体");
        cellstyle2.setFont(font);
        cellstyle2.setBorderBottom(BorderStyle.THIN); //下边框
        cellstyle2.setBorderLeft(BorderStyle.THIN);//左边框
        cellstyle2.setBorderTop(BorderStyle.THIN);//上边框
        cellstyle2.setBorderRight(BorderStyle.THIN);//右边框
        cellstyle2.setVerticalAlignment(VerticalAlignment.CENTER);//垂直
        cellstyle2.setAlignment(HorizontalAlignment.CENTER);//水平

        sheet.getRow(tableNum*29+26).getCell((1+index-1)).setCellFormula("IF("
                +sheet.getRow(tableNum*29+25).getCell((1+index-1)).getReference()+">=94,\"√\",\"×\")");//B27=IF(B26>=94,"√","×")

        System.out.println("row.getCell(8).getRawValue() = "+sheet.getRow(tableNum*29+26).getCell((1+index-1)).getReference()+" - > "
                +evaluate.evaluateFormulaCell(sheet.getRow(tableNum*29+26).getCell((1+index-1)))+" -> "
                +sheet.getRow(tableNum*29+26).getCell((1+index-1)).getRawValue());

        if("×".equals(sheet.getRow(tableNum*29+26).getCell((1+index-1)).getRawValue())){
            sheet.getRow(tableNum*29+26).getCell((1+index-1)).setCellStyle(cellstyle2);
        }
    }

    /**
     * 填写中间的合并单元格
     * @param sheet
     * @param tableNum
     * @param index
     * @param row
     */
    public void fillMergeCellData(XSSFSheet sheet, int tableNum, int index, JjgFbgcLjgcLjtsfysdHt row){
        sheet.getRow(tableNum*29+7).getCell(1+index).setCellValue(row.getQyzhjwz());//取样桩号及位置
        sheet.getRow(tableNum*29+8).getCell(1+index).setCellValue((Double.parseDouble(row.getSksd())));//试坑深度
        sheet.getRow(tableNum*29+9).getCell(1+index).setCellValue((Double.parseDouble(row.getZtjjbhbmjszl())));//锥体及基板和表面间砂质量（g）
        sheet.getRow(tableNum*29+10).getCell(1+index).setCellValue((Double.parseDouble(row.getGsqtszl())));//灌砂前筒+砂质量（g）
        sheet.getRow(tableNum*29+11).getCell(1+index).setCellValue((Double.parseDouble(row.getGshtszl())));//灌砂后筒+砂质量（g）
        sheet.getRow(tableNum*29+12).getCell(1+index).setCellType(Cell.CELL_TYPE_FORMULA);
        sheet.getRow(tableNum*29+12).getCell(1+index).setCellFormula(sheet.getRow(tableNum*29+10).getCell(1+index).getReference()+"-"
                +sheet.getRow(tableNum*29+11).getCell(1+index).getReference()+"-"
                +sheet.getRow(tableNum*29+9).getCell(1+index).getReference());//B13=B11-B12-B10
        sheet.getRow(tableNum*29+13).getCell(1+index).setCellValue((Double.parseDouble(row.getSyzl())));//试样质量（g）
        sheet.getRow(tableNum*29+14).getCell(1+index).setCellFormula("IF("
                +sheet.getRow(tableNum*29+6).getCell(8).getReference()+"=0,\"\","
                +sheet.getRow(tableNum*29+12).getCell(1+index).getReference()+"/"
                +sheet.getRow(tableNum*29+6).getCell(8).getReference()+")");//B15=IF(I7=0," ",B13/I7)
        sheet.getRow(tableNum*29+15).getCell(1+index).setCellFormula("IF("
                +sheet.getRow(tableNum*29+14).getCell(1+index).getReference()+"=0,\"\","
                +sheet.getRow(tableNum*29+13).getCell(1+index).getReference()+"/"
                +sheet.getRow(tableNum*29+14).getCell(1+index).getReference()+")");//B16=IF(B15=0," ",B14/B15)
    }

    /**
     * 填写中间下方的没有合并单元格的普通单元格
     * @param sheet
     * @param tableNum
     * @param index
     * @param row
     * @param cellstyle
     */
    public void fillCommonCellData(XSSFWorkbook xwb,XSSFSheet sheet, int tableNum, int index, JjgFbgcLjgcLjtsfysdHt row, XSSFCellStyle cellstyle){
        try{
            sheet.getRow(tableNum*29+16).getCell(1+index).setCellValue(Double.parseDouble(row.getHh()));//盒号
        }
        catch(Exception e){
            sheet.getRow(tableNum*29+16).getCell(1+index).setCellValue(row.getHh());//盒号
        }
        XSSFFont font=xwb.createFont();
        font.setFontHeightInPoints((short)9);
        font.setFontName("宋体");
        cellstyle.setFont(font);
        sheet.getRow(tableNum*29+17).getCell(1+index).setCellStyle(cellstyle);
        sheet.getRow(tableNum*29+17).getCell(1+index).setCellValue((Double.parseDouble(row.getHzl())));//盒质量（g）
        sheet.getRow(tableNum*29+18).getCell(1+index).setCellValue((Double.parseDouble(row.getHsshzl())));//盒+湿试样质量（g）
        sheet.getRow(tableNum*29+19).getCell(1+index).setCellValue((Double.parseDouble(row.getHgsyzl())));//盒+干试样质量（g）
        sheet.getRow(tableNum*29+20).getCell(1+index).setCellType(Cell.CELL_TYPE_FORMULA);
        sheet.getRow(tableNum*29+20).getCell(1+index).setCellFormula(sheet.getRow(tableNum*29+18).getCell(1+index).getReference()+"-"
                +sheet.getRow(tableNum*29+19).getCell(1+index).getReference());//B21=B19-B20
        sheet.getRow(tableNum*29+21).getCell(1+index).setCellFormula(sheet.getRow(tableNum*29+19).getCell(1+index).getReference()+"-"
                +sheet.getRow(tableNum*29+17).getCell(1+index).getReference());//B22=B20-B18

        sheet.getRow(tableNum*29+22).getCell(1+index).setCellFormula("IF("
                +sheet.getRow(tableNum*29+21).getCell(1+index).getReference()+"=0,\"\","
                +sheet.getRow(tableNum*29+20).getCell(1+index).getReference()+"/"
                +sheet.getRow(tableNum*29+21).getCell(1+index).getReference()+"*100)");//B23=IF(B22=0," ",B21/B22*100)
    }

    /**
     * 填写表头
     * @param sheet
     * @param tableNum
     * @param row
     * @param position
     * @param proname
     * @param htd
     */
    public void fillTitleCellData(XSSFSheet sheet, int tableNum, JjgFbgcLjgcLjtsfysdHt row,String position,String proname,String htd){
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        sheet.getRow(tableNum*29+1).getCell(1).setCellValue(proname);
        sheet.getRow(tableNum*29+1).getCell(8).setCellValue(htd);
        sheet.getRow(tableNum*29+2).getCell(1).setCellValue(position);
        sheet.getRow(tableNum*29+3).getCell(1).setCellValue("路基土石方");
        sheet.getRow(tableNum*29+3).getCell(8).setCellValue(simpleDateFormat.format(row.getSysj()));
        sheet.getRow(tableNum*29+5).getCell(1).setCellValue(row.getJgcc());
        sheet.getRow(tableNum*29+5).getCell(8).setCellValue(row.getJglx());
        sheet.getRow(tableNum*29+6).getCell(1).setCellValue((Double.parseDouble(row.getZdgmd())));
        sheet.getRow(tableNum*29+6).getCell(8).setCellValue((Double.parseDouble(row.getBzsmd())));
    }

    /**
     * 获取要生成的页数
     * @param numlist
     * @return
     */
    public int gettableNum(List<String> numlist){
        int tableNum = 0;
        for (String i: numlist){
            int dataNum=0;
            dataNum = Integer.parseInt(i);
            tableNum+= dataNum%10 == 0 ? dataNum/10 : dataNum/10+1;

        }
        return tableNum;
    }

    /**
     * 根据页数复制模板
     * @param tableNum
     * @throws IOException
     */
    public void createTable(int tableNum) throws IOException{
        for(int i = 1; i < tableNum; i++){
            RowCopy.copyRows(xwb, "压实度单点(灰土)", "压实度单点(灰土)", 0, 28, i*29);
        }
        if(tableNum > 1){
            xwb.setPrintArea(xwb.getSheetIndex("压实度单点(灰土)"), 0, 10, 0, tableNum*29-1);
        }
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        String title = "路基压实度质量鉴定表（评定单元）";
        String slsheetname = "评定单元";
        String htsheetname = "评定单元(2)";
        DecimalFormat df = new DecimalFormat(".00");
        DecimalFormat bzcdf = new DecimalFormat(".000");
        DecimalFormat decf = new DecimalFormat("0.##");
        //获取鉴定表文件
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"01路基土石方压实度.xlsx");
        if(!f.exists()){
            return null;
        }else {
            XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));
            XSSFSheet slSheet = xwb.getSheet(slsheetname);
            XSSFCell bt = slSheet.getRow(0).getCell(0);
            XSSFCell xmname = slSheet.getRow(1).getCell(2);
            XSSFCell htdname = slSheet.getRow(1).getCell(7);
            XSSFCell hd = slSheet.getRow(2).getCell(2);
            boolean slsheetHidden = xwb.isSheetHidden(0);
            boolean htsheetHidden = xwb.isSheetHidden(1);
            List<Map<String,Object>> mapList = new ArrayList<>();
            Map<String,Object> sljgmap = new HashMap<>();
            Map<String,Object> htjgmap = new HashMap<>();
            if (!slsheetHidden && proname.equals(xmname.toString()) && title.equals(bt.toString()) && htd.equals(htdname.toString()) && fbgc.equals(hd.toString())){
                int sllastRowNum = slSheet.getLastRowNum();
                slSheet.getRow(sllastRowNum-2).getCell(3).setCellType(CellType.STRING);//平均值
                slSheet.getRow(sllastRowNum-2).getCell(5).setCellType(CellType.STRING);//标准差
                slSheet.getRow(sllastRowNum-2).getCell(7).setCellType(CellType.STRING);//代表值
                slSheet.getRow(sllastRowNum-1).getCell(3).setCellType(CellType.STRING);//规定值
                slSheet.getRow(sllastRowNum-1).getCell(6).setCellType(CellType.STRING);//结果
                slSheet.getRow(sllastRowNum).getCell(3).setCellType(CellType.STRING);//检测点数
                slSheet.getRow(sllastRowNum).getCell(5).setCellType(CellType.STRING);//合格点数
                slSheet.getRow(sllastRowNum).getCell(7).setCellType(CellType.STRING);//合格率（%）

                sljgmap.put("压实度项目","沙砾");
                sljgmap.put("检测点数",decf.format(Double.valueOf(slSheet.getRow(sllastRowNum).getCell(3).getStringCellValue())));
                sljgmap.put("合格点数",decf.format(Double.valueOf(slSheet.getRow(sllastRowNum).getCell(5).getStringCellValue())));
                sljgmap.put("合格率",df.format(Double.valueOf(slSheet.getRow(sllastRowNum).getCell(7).getStringCellValue())));
                sljgmap.put("平均值",df.format(Double.valueOf(slSheet.getRow(sllastRowNum-2).getCell(3).getStringCellValue())));
                sljgmap.put("标准差",Double.valueOf(slSheet.getRow(sllastRowNum-2).getCell(5).getStringCellValue()));
                sljgmap.put("代表值",df.format(Double.valueOf(slSheet.getRow(sllastRowNum-2).getCell(7).getStringCellValue())));
                sljgmap.put("规定值",decf.format(Double.valueOf(slSheet.getRow(sllastRowNum-1).getCell(3).getStringCellValue())));
                sljgmap.put("结果",slSheet.getRow(sllastRowNum-1).getCell(6).getStringCellValue());
                mapList.add(sljgmap);
            }else {
                sljgmap.put("压实度项目","沙砾");
                sljgmap.put("检测点数",0);
                sljgmap.put("合格点数",0);
                sljgmap.put("合格率",0);
                sljgmap.put("平均值",0);
                sljgmap.put("标准差",0);
                sljgmap.put("代表值",0);
                sljgmap.put("规定值",0);
                sljgmap.put("结果","无结果");
                mapList.add(sljgmap);

            }
            if (!htsheetHidden && proname.equals(xmname.toString()) && title.equals(bt.toString()) && htd.equals(htdname.toString()) && fbgc.equals(hd.toString())){
                XSSFSheet htSheet = xwb.getSheet(htsheetname);
                int htlastRowNum = htSheet.getLastRowNum();

                htSheet.getRow(htlastRowNum-3).getCell(3).setCellType(CellType.STRING);//平均值
                htSheet.getRow(htlastRowNum-3).getCell(5).setCellType(CellType.STRING);//标准差
                htSheet.getRow(htlastRowNum-3).getCell(7).setCellType(CellType.STRING);//代表值
                htSheet.getRow(htlastRowNum-2).getCell(3).setCellType(CellType.STRING);//规定值
                htSheet.getRow(htlastRowNum-2).getCell(6).setCellType(CellType.STRING);//结果
                htSheet.getRow(htlastRowNum-1).getCell(3).setCellType(CellType.STRING);//检测点数
                htSheet.getRow(htlastRowNum-1).getCell(5).setCellType(CellType.STRING);//合格点数
                htSheet.getRow(htlastRowNum-1).getCell(7).setCellType(CellType.STRING);//合格率（%）
                htjgmap.put("压实度项目","灰土");
                htjgmap.put("检测点数",decf.format(Double.valueOf(htSheet.getRow(htlastRowNum-1).getCell(3).getStringCellValue())));
                htjgmap.put("合格点数",decf.format(Double.valueOf(htSheet.getRow(htlastRowNum-1).getCell(5).getStringCellValue())));
                htjgmap.put("合格率",df.format(Double.valueOf(htSheet.getRow(htlastRowNum-1).getCell(7).getStringCellValue())));
                htjgmap.put("平均值",df.format(Double.valueOf(htSheet.getRow(htlastRowNum-3).getCell(3).getStringCellValue())));
                htjgmap.put("标准差",bzcdf.format(Double.valueOf(htSheet.getRow(htlastRowNum-3).getCell(5).getStringCellValue())));
                htjgmap.put("代表值",df.format(Double.valueOf(htSheet.getRow(htlastRowNum-3).getCell(7).getStringCellValue())));
                htjgmap.put("规定值",decf.format(Double.valueOf(htSheet.getRow(htlastRowNum-2).getCell(3).getStringCellValue())));
                htjgmap.put("结果",htSheet.getRow(htlastRowNum-2).getCell(6).getStringCellValue());
                mapList.add(htjgmap);
            }else {
                htjgmap.put("压实度项目","灰土");
                htjgmap.put("检测点数",0);
                htjgmap.put("合格点数",0);
                htjgmap.put("合格率",0);
                htjgmap.put("平均值",0);
                htjgmap.put("标准差",0);
                htjgmap.put("代表值",0);
                htjgmap.put("规定值",0);
                htjgmap.put("结果","无结果");
                mapList.add(htjgmap);

            }
            return mapList;
        }
    }

    /**
     * 导出实测数据模板文件
     * @param response
     * @throws IOException
     */
    @Override
    public void exportysdht(HttpServletResponse response) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();// 创建一个Excel文件
        XSSFCellStyle columnHeadStyle = JjgFbgcCommonUtils.tsfCellStyle(workbook);

        XSSFSheet sheet = workbook.createSheet("实测数据");// 创建一个Excel的Sheet
        sheet.setColumnWidth(0,28*256);
        List<String> checklist = Arrays.asList("路基压实度_灰土规定值","试验时间","桩号","结构层次",
                "结构类型","最大干密度","标准砂密度","取样桩号及位置","试坑深度(cm)",
                "锥体及基板和表面间砂质量（g）","灌砂前筒+砂质量（g）","灌砂后筒+砂质量（g）",
                "试样质量（g）","盒号","盒质量（g）","盒+湿试样质量（g）","盒+干试样质量（g）","序号");
        for (int i=0;i< checklist.size();i++){
            XSSFRow row = sheet.createRow(i);// 创建第一行
            XSSFCell cell = row.createCell(0);// 创建第一行第一列
            cell.setCellValue(new XSSFRichTextString(checklist.get(i)));
            cell.setCellStyle(columnHeadStyle);
        }
        String filename = "01路基土石方压实度_灰土实测数据.xls";// 设置下载时客户端Excel的名称
        filename = new String((filename).getBytes("GBK"), "ISO8859_1");
        response.setContentType("application/octet-stream;charset=UTF-8");
        response.addHeader("Content-Disposition", "attachment;filename=" + filename);
        OutputStream ouputStream = response.getOutputStream();
        workbook.write(ouputStream);
        ouputStream.flush();
        ouputStream.close();


    }

    /**
     * 导入实测数据
     * @param file
     * @param commonInfoVo
     * @throws IOException
     * @throws ParseException
     */
    @Override
    public void importysdht(MultipartFile file, CommonInfoVo commonInfoVo) throws IOException, ParseException {
        // 将文件流传过来，变成workbook对象
        XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
        //获得文本
        XSSFSheet sheet = workbook.getSheetAt(0);
        //获得行数
        int rows = sheet.getPhysicalNumberOfRows();
        //获得列数
        int columns = 0;
        for(int i=1;i<rows;i++){
            XSSFRow row = sheet.getRow(i);
            columns  = row.getPhysicalNumberOfCells();
        }
        JjgFbgcLjgcLjtsfysdHtVo jjgFbgcLjgcLjtsfysdHtVo = new JjgFbgcLjgcLjtsfysdHtVo();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        List<String> titlelist = new ArrayList();
        for (int n=0;n<1;n++){
            for (int m=0;m<rows;m++){
                XSSFRow row = sheet.getRow(m);
                XSSFCell cell = row.getCell(n);
                titlelist.add(cell.toString());
            }
        }
        List<String> checklist = Arrays.asList("路基压实度_灰土规定值","试验时间","桩号","结构层次",
                "结构类型","最大干密度","标准砂密度","取样桩号及位置","试坑深度(cm)",
                "锥体及基板和表面间砂质量（g）","灌砂前筒+砂质量（g）","灌砂后筒+砂质量（g）",
                "试样质量（g）","盒号","盒质量（g）","盒+湿试样质量（g）","盒+干试样质量（g）","序号");
        if(checklist.equals(titlelist)){
            for (int j = 1;j<columns;j++){//列
                Map<String,Object> map = new HashMap<>();
                Field[] fields = jjgFbgcLjgcLjtsfysdHtVo.getClass().getDeclaredFields();
                JjgFbgcLjgcLjtsfysdHt jjgFbgcLjgcLjtsfysdHt = new JjgFbgcLjgcLjtsfysdHt();
                for(int k=0;k<rows;k++){//行
                    //列是不变的 行增加
                    XSSFRow row = sheet.getRow(k);
                    XSSFCell cell = row.getCell(j);

                    //cell.setCellType(CellType.STRING);
                    System.out.println(cell);
                    switch (cell.getCellType()){
                        case XSSFCell.CELL_TYPE_STRING ://String
                            cell.setCellType(CellType.STRING);
                            map.put(fields[k].getName(),cell.getStringCellValue());//属性赋值
                            break;
                        case XSSFCell.CELL_TYPE_BOOLEAN ://bealean
                            //cell.setCellType(XSSFCell.CELL_TYPE_STRING);
                            cell.setCellType(CellType.STRING);
                            //map.put(fields[k].getName(),Boolean.valueOf(cell.getBooleanCellValue()).toString());//属性赋值
                            map.put(fields[k].getName(),String.valueOf(cell.getStringCellValue()));//属性赋值
                            break;
                        case XSSFCell.CELL_TYPE_NUMERIC ://number
                            //默认日期读取出来是数字，判断是否是日期格式的数字
                            if(DateUtil.isCellDateFormatted(cell)){
                                //读取的数字是日期，转换一下格式
                                DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                                Date date = cell.getDateCellValue();
                                System.out.println(date);
                                map.put(fields[k].getName(),dateFormat.format(date));//属性赋值
                            }else {//不是日期直接赋值
                                cell.setCellType(CellType.STRING);
                                //map.put(fields[k].getName(),Double.valueOf(cell.getNumericCellValue()).toString());//属性赋值
                                map.put(fields[k].getName(),String.valueOf(cell.getStringCellValue()));//属性赋值
                            }
                            break;
                        case XSSFCell.CELL_TYPE_BLANK :
                            cell.setCellType(CellType.STRING);
                            map.put(fields[k].getName(),"");//属性赋值
                            break;
                        default:
                            System.out.println("未知类型------>"+cell);
                    }
                }
                jjgFbgcLjgcLjtsfysdHt.setHtgdz((String) map.get("htgdz"));
                jjgFbgcLjgcLjtsfysdHt.setSysj(simpleDateFormat.parse((String) map.get("sysj")));
                jjgFbgcLjgcLjtsfysdHt.setZh((String) map.get("zh"));
                jjgFbgcLjgcLjtsfysdHt.setJgcc((String) map.get("jgcc"));
                jjgFbgcLjgcLjtsfysdHt.setJglx((String) map.get("jglx"));
                jjgFbgcLjgcLjtsfysdHt.setZdgmd((String) map.get("zdgmd"));
                jjgFbgcLjgcLjtsfysdHt.setBzsmd((String) map.get("bzsmd"));
                jjgFbgcLjgcLjtsfysdHt.setQyzhjwz((String) map.get("qyzhjwz"));
                jjgFbgcLjgcLjtsfysdHt.setSksd((String) map.get("sksd"));
                jjgFbgcLjgcLjtsfysdHt.setZtjjbhbmjszl((String) map.get("ztjjbhbmjszl"));
                jjgFbgcLjgcLjtsfysdHt.setGsqtszl((String) map.get("gsqtszl"));
                jjgFbgcLjgcLjtsfysdHt.setGshtszl((String) map.get("gshtszl"));
                jjgFbgcLjgcLjtsfysdHt.setSyzl((String) map.get("syzl"));
                jjgFbgcLjgcLjtsfysdHt.setHh((String) map.get("hh"));
                jjgFbgcLjgcLjtsfysdHt.setHzl((String) map.get("hzl"));
                jjgFbgcLjgcLjtsfysdHt.setHsshzl((String) map.get("hsshzl"));
                jjgFbgcLjgcLjtsfysdHt.setHgsyzl((String) map.get("hgsyzl"));
                jjgFbgcLjgcLjtsfysdHt.setXh((String) map.get("xh"));
                jjgFbgcLjgcLjtsfysdHt.setProname(commonInfoVo.getProname());
                jjgFbgcLjgcLjtsfysdHt.setHtd(commonInfoVo.getHtd());
                jjgFbgcLjgcLjtsfysdHt.setFbgc("路基土石方");
                jjgFbgcLjgcLjtsfysdHt.setCreatetime(new Date());
                jjgFbgcLjgcLjtsfysdHtMapper.insert(jjgFbgcLjgcLjtsfysdHt);
            }
        }else {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }
    }
}
