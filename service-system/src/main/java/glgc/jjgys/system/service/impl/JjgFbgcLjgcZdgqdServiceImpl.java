package glgc.jjgys.system.service.impl;

import cn.hutool.core.io.resource.ClassPathResource;
import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLjgcZdgqd;
import glgc.jjgys.model.project.JjgHtd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.ljgc.JjgFbgcLjgcZdgqdVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcLjgcZdgqdMapper;
import glgc.jjgys.system.service.JjgFbgcLjgcZdgqdService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.poi.ss.usermodel.*;
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
 *  支档砼强度
 * </p>
 *
 * @author wq
 * @since 2023-02-15
 */
@Service
public class JjgFbgcLjgcZdgqdServiceImpl extends ServiceImpl<JjgFbgcLjgcZdgqdMapper, JjgFbgcLjgcZdgqd> implements JjgFbgcLjgcZdgqdService {


    @Autowired
    private JjgFbgcLjgcZdgqdMapper jjgFbgcLjgcZdgqdMapper;


    //private static XSSFWorkbook wb = null;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = "支挡工程";
        //获取数据
        QueryWrapper<JjgFbgcLjgcZdgqd> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("zh","bw1");
        List<JjgFbgcLjgcZdgqd> data = jjgFbgcLjgcZdgqdMapper.selectList(wrapper);
        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"10路基支挡砼强度.xlsx");
        //健壮性判断如果没有数据返回"请导入数据"
        if (data == null || data.size()==0){
            return;
        }else {
            //存放鉴定表的目录
            File fdir = new File(filepath+File.separator+proname+File.separator+htd);
            if(!fdir.exists()){
                //创建文件根目录
                fdir.mkdirs();
            }
            //File directory = new File("service-system/src/main/resources/static");
            File directory = new File("service-system/src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String name = "支挡砼强度.xlsx";
            String path =reportPath+File.separator+name;
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            createTable(gettableNum(data.size()),wb);
            if(DBtoExcel(data,proname,htd,fbgc,wb)){
                //设置公式,计算合格点数
                calculateSheet(wb);
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

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String,Object>> mapList = new ArrayList<>();
        Map<String,Object> jgmap = new HashMap<>();
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();

        String sheetname = "原始数据";
        //获取鉴定表文件
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"10路基支挡砼强度.xlsx");
        if(!f.exists()){
            return null;
        }else {
            //创建工作簿
            XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));
            //读取工作表
            XSSFSheet slSheet = xwb.getSheet(sheetname);
            slSheet.getRow(2).getCell(34).setCellType(CellType.STRING);
            slSheet.getRow(2).getCell(35).setCellType(CellType.STRING);
            slSheet.getRow(2).getCell(36).setCellType(CellType.STRING);
            double zds= Double.valueOf(slSheet.getRow(2).getCell(34).getStringCellValue());
            double hgds= Double.valueOf(slSheet.getRow(2).getCell(35).getStringCellValue());
            double hgl= Double.valueOf(slSheet.getRow(2).getCell(36).getStringCellValue());
            String zdsz = decf.format(zds);
            String hgdsz = decf.format(hgds);
            String hglz = df.format(hgl);
            jgmap.put("总点数",zdsz);
            jgmap.put("合格点数",hgdsz);
            jgmap.put("合格率",hglz);
            mapList.add(jgmap);

        }
        return mapList;

    }

    @Override
    public void exportzdgqd(HttpServletResponse response) {
        String fileName = "10支挡砼强度实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcLjgcZdgqdVo()).finish();

    }

    @Override
    public void importzdgqd(MultipartFile file,CommonInfoVo commonInfoVo) {

        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcLjgcZdgqdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcLjgcZdgqdVo>(JjgFbgcLjgcZdgqdVo.class) {
                                @Override
                                public void handle(List<JjgFbgcLjgcZdgqdVo> dataList) {
                                    for(JjgFbgcLjgcZdgqdVo zdgqdVo: dataList)
                                    {
                                        JjgFbgcLjgcZdgqd fbgcLjgcZdgqd = new JjgFbgcLjgcZdgqd();
                                        BeanUtils.copyProperties(zdgqdVo,fbgcLjgcZdgqd);
                                        fbgcLjgcZdgqd.setCreatetime(new Date());
                                        fbgcLjgcZdgqd.setProname(commonInfoVo.getProname());
                                        fbgcLjgcZdgqd.setHtd(commonInfoVo.getHtd());
                                        fbgcLjgcZdgqd.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcLjgcZdgqdMapper.insert(fbgcLjgcZdgqd);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }


    }

    public void calculateSheet(XSSFWorkbook wb) {
        XSSFSheet sheet = wb.getSheet("原始数据");
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        boolean flag = false;
        String name = "";
        ArrayList<XSSFRow> startRowList = new ArrayList<XSSFRow>();
        ArrayList<XSSFRow> endRowList = new ArrayList<XSSFRow>();
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
        double value29 = 0.0;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            // 下一张表
            if (flag && !"".equals(row.getCell(0).toString()) && row.getCell(0).toString().contains("质量鉴定表")) {
                flag = false;
            }
            // 可以计算
            if (row.getCell(3).getCellType() == Cell.CELL_TYPE_NUMERIC && flag) {
                //此处为计算平均值，因为poi不支持excel公式TRIMMEAN，所以要自己写代码解决
                int[] value = new int[16];
                for (int index = 0; index < 16; index++) {
                    try {
                        value[index] = (int) row.getCell(3 + index).getNumericCellValue();
                    } catch (Throwable th) {
                        continue;
                    }
                }
                Arrays.sort(value);
                String array = "";
                for (int index = 0; index < 10; index++) {
                    array += value[index + 3] + ",";
                }
                row.getCell(19).setCellFormula(
                        "IF(" + row.getCell(3).getReference()
                                + ">0,ROUND(AVERAGE("
                                + array.substring(0, array.lastIndexOf(","))
                                + "),1),\" \")");
                // T=IF(D6>0,ROUND(TRIMMEAN(D6:S6,0.375),1)," ")
                evaluator.evaluateInCell(row.getCell(19));
                row.getCell(19).getCellStyle().setLocked(true);
                // V=IF(D6="","",INDEX(修正数据!$O$2:$X$403,MATCH(原始数据!T6,修正数据!$O$2:$O$403,0),MATCH(原始数据!U6,修正数据!$O$2:$X$2,0)))
                row.getCell(21).setCellValue(getjdxzz(wb.getSheet("修正数据"), row.getCell(19).getNumericCellValue(), row.getCell(20).toString()));
                row.getCell(21).getCellStyle().setLocked(true);

                row.getCell(22).setCellFormula(
                        "IF(" + row.getCell(3).getReference()
                                + "=\"\",\"\",ROUND("
                                + row.getCell(19).getReference() + "+"
                                + row.getCell(21).getReference() + ",1))");// W=IF(D6="","",ROUND(T6+V6,1))
                evaluator.evaluateInCell(row.getCell(22));
                row.getCell(22).getCellStyle().setLocked(true);
                // Y=IF(D6="","",INDEX(修正数据!$Y$2:$AB$403,MATCH(原始数据!W6,修正数据!$Y$2:$Y$403,0),MATCH(原始数据!X6,修正数据!$Y$2:$AB$2,0)))
                row.getCell(24).setCellValue(getjzmxzz(wb.getSheet("修正数据"), row.getCell(22).getNumericCellValue(), row.getCell(23).toString()));
                row.getCell(24).getCellStyle().setLocked(true);
                row.getCell(25).setCellFormula(
                        "IF(" + row.getCell(22).getReference()
                                + "=\"\",\"\",ROUND("
                                + row.getCell(22).getReference() + "+"
                                + row.getCell(24).getReference() + ",1))");
                // Z=IF(W6="","",ROUND(W6+Y6,1))
                evaluator.evaluateInCell(row.getCell(25));
                row.getCell(25).getCellStyle().setLocked(true);

                // AB=IF(Z6="","",INDEX(修正数据!$A$4:$N$405,MATCH(原始数据!Z6,修正数据!$A$4:$A$405,0),MATCH(原始数据!AA6,修正数据!$A$4:$N$4,0)))
                if (getthsdhsz(wb.getSheet("修正数据"), row.getCell(25).getNumericCellValue(), row.getCell(26).toString()) == 60) {
                    XSSFCellStyle cellstyle = wb.createCellStyle();
                    XSSFFont font = wb.createFont();
                    font.setFontHeightInPoints((short) 10);
                    font.setFontName("宋体");
                    cellstyle.setFont(font);
                    cellstyle.setBorderBottom(BorderStyle.THIN); //下边框
                    cellstyle.setBorderLeft(BorderStyle.THIN);//左边框
                    cellstyle.setBorderTop(BorderStyle.THIN);//上边框
                    cellstyle.setBorderRight(BorderStyle.THIN);//右边框
                    cellstyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直
                    cellstyle.setAlignment(HorizontalAlignment.CENTER);//水平
                    XSSFDataFormat format = wb.createDataFormat();
                    cellstyle.setDataFormat(format.getFormat(">##0"));
                    row.getCell(29).setCellStyle(cellstyle);
                }
                row.getCell(27).setCellValue(getthsdhsz(wb.getSheet("修正数据"), row.getCell(25).getNumericCellValue(), row.getCell(26).toString()));
                row.getCell(27).getCellStyle().setLocked(true);

                if (row.getCell(3) == null || "".equals(row.getCell(3).toString())) {
                    row.getCell(29).setCellValue("");
                } else {
                    if ("是".equals(row.getCell(28).toString())) {
                        double Z = evaluator.evaluate(row.getCell(25)).getNumberValue();
                        double AA = Double.valueOf(row.getCell(26).toString());
                        Double v = 0.034488 * Math.pow(Z, 1.94) * Math.pow(10, (-0.0176 * AA));
                        if (v.intValue() > 60) {
                            row.getCell(29).setCellValue(60.0);
                        } else {
                            row.getCell(29).setCellValue(v);
                        }
                    } else {
                        row.getCell(29).setCellFormula(row.getCell(27).getReference());
                    }
                }
                row.getCell(29).getCellStyle().setLocked(true);
                row.getCell(32).setCellFormula(
                        "IF(" + row.getCell(29).getReference() + ">="
                                + row.getCell(31).getReference()
                                + ",\"√\",\"\")");
                // AG=IF(AD6>=AF6,"√","")
                row.getCell(32).getCellStyle().setLocked(true);

                row.getCell(33).setCellFormula(
                        "IF(" + row.getCell(32).getReference()
                                + "=\"\",\"×\",\"\")");
                // AH=IF(AG6="","×","")
                row.getCell(33).getCellStyle().setLocked(true);
            }

            // 可以计算啦
            if(flag && !"".equals(row.getCell(0).toString()) && !name.equals(getBridgeName(row.getCell(0).toString()))){
                //计算这座桥的总计
                String column29 = "";
                String column32 = "";
                startRowList.get(0).createCell(34);
                startRowList.get(0).createCell(35);
                startRowList.get(0).createCell(36);
                for(int r = 0;r < startRowList.size(); r++){
                    column29 += "COUNT("+startRowList.get(0).getCell(29).getReference()+":"+endRowList.get(0).getCell(29).getReference()+")+";
                    column32 += "COUNTIF("+startRowList.get(0).getCell(32).getReference()+":"+endRowList.get(0).getCell(32).getReference()+",\"√\")+";
                }
                startRowList.get(0).getCell(34).setCellFormula(column29.substring(0, column29.lastIndexOf('+')));
                startRowList.get(0).getCell(35).setCellFormula(column32.substring(0, column32.lastIndexOf('+')));//=COUNTIF(AG6:AG50,"√")
                startRowList.get(0).getCell(36).setCellFormula(startRowList.get(0).getCell(35).getReference()+"*100/"
                        +startRowList.get(0).getCell(34).getReference());//=COUNTIF(AH6:AH50,"×")
                startRowList.clear();
                endRowList.clear();
                startRowList.add(sheet.getRow(i));
                endRowList.add(sheet.getRow(i+9));
                name = getBridgeName(row.getCell(0).toString());
            } else if(flag && !"".equals(row.getCell(0).toString()) && name.equals(getBridgeName(row.getCell(0).toString()))){
                System.out.println(name);
                startRowList.add(sheet.getRow(i));
                endRowList.add(sheet.getRow(i+9));
            }
            if ("桩号/\n结构名称".equals(row.getCell(0).toString())) {
                flag = true;
                i++;
                rowstart = sheet.getRow(i+1);
                //计算每座桥的总点数，合格点数，不合格点数
                if(!name.equals(getBridgeName(sheet.getRow(i+1).getCell(0).toString())) && !"".equals(name)){
                    String column29 = "";
                    String column32 = "";
                    startRowList.get(0).createCell(34);
                    startRowList.get(0).createCell(35);
                    startRowList.get(0).createCell(36);
                    for(int r = 0;r < startRowList.size(); r++){
                        column29 += "COUNT("+startRowList.get(r).getCell(29).getReference()+":"+endRowList.get(r).getCell(29).getReference()+")+";
                        column32 += "COUNTIF("+startRowList.get(r).getCell(32).getReference()+":"+endRowList.get(r).getCell(32).getReference()+",\"√\")+";
                    }
                    startRowList.get(0).getCell(34).setCellFormula(column29.substring(0, column29.lastIndexOf('+')));
                    startRowList.get(0).getCell(35).setCellFormula(column32.substring(0, column32.lastIndexOf('+')));//=COUNTIF(AG6:AG50,"√")
                    startRowList.get(0).getCell(36).setCellFormula(startRowList.get(0).getCell(35).getReference()+"*100/"+startRowList.get(0).getCell(34).getReference());//=COUNTIF(AH6:AH50,"×")
                    startRowList.clear();
                    endRowList.clear();
                }
                name = getBridgeName(sheet.getRow(i+1).getCell(0).toString());
                rowend = sheet.getRow(i+10);
            }

        }
            /*
             * 计算最后一座桥的总点数，合格点数，不合格点数
             */
        String column29 = "";
        String column32 = "";
        startRowList.get(0).createCell(34);
        startRowList.get(0).createCell(35);
        startRowList.get(0).createCell(36);
        for(int r = 0;r < startRowList.size(); r++){
            column29 = "COUNT("+startRowList.get(r).getCell(29).getReference()+":"+endRowList.get(r).getCell(29).getReference()+")+";
            column32 = "COUNTIF("+startRowList.get(r).getCell(32).getReference()+":"+endRowList.get(r).getCell(32).getReference()+",\"√\")+";
        }
        startRowList.get(0).getCell(34).setCellFormula(column29.substring(0, column29.lastIndexOf('+')));
        startRowList.get(0).getCell(35).setCellFormula(column32.substring(0, column32.lastIndexOf('+')));//=COUNTIF(AG6:AG50,"√")
        startRowList.get(0).getCell(36).setCellFormula(startRowList.get(0).getCell(35).getReference()+"*100/"+startRowList.get(0).getCell(34).getReference());//=COUNTIF(AH6:AH50,"×")
        //计算总点数，合格点数，合格率
        sheet.getRow(2).createCell(34).setCellFormula("SUM("
                +sheet.getRow(5).getCell(34).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(34).getReference()+")");//总点数=SUM(AI5:AI50)
        sheet.getRow(2).createCell(35).setCellFormula("SUM("
                +sheet.getRow(5).getCell(35).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(35).getReference()+")");//合格点数=SUM(AJ5:AJ50)
        sheet.getRow(2).createCell(36).setCellFormula(sheet.getRow(2).getCell(35).getReference()+"*100/"
                +sheet.getRow(2).getCell(34).getReference());//合格率

    }

    @Override
    public List<Map<String, Object>> selectsjqd(String proname, String htd) {
        List<Map<String, Object>> list = jjgFbgcLjgcZdgqdMapper.selectsjqd(proname,htd);
        return list;
    }

    @Override
    public Map<String, Object> selectchs(String proname, String htd) {
        Map<String, Object> map = jjgFbgcLjgcZdgqdMapper.selectchs(proname,htd);
        return map;
    }

    @Override
    public int selectnum(String proname, String htd) {
        int selectnum = jjgFbgcLjgcZdgqdMapper.selectnum(proname, htd);
        return selectnum;
    }

    public int gettableNum(int size){
        return size%20 ==0 ? size/20 : size/20+1;
    }

    public void createTable(int tableNum,XSSFWorkbook wb){
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            RowCopy.copyRows(wb, "原始数据", "原始数据", 0, 24, i * 25);
        }
        wb.getSheet("原始数据").setColumnHidden(21, true);
        wb.getSheet("原始数据").setColumnHidden(22, true);
        wb.getSheet("原始数据").setColumnHidden(24, true);
        wb.getSheet("原始数据").setColumnHidden(25, true);
        wb.getSheet("原始数据").setColumnHidden(27, true);
        wb.getSheet("原始数据").setColumnHidden(30, true);
        if(record > 1)
            wb.setPrintArea(wb.getSheetIndex("原始数据"), 0, 33, 0, record * 25-1);
    }

    public boolean DBtoExcel(List<JjgFbgcLjgcZdgqd> data,String proname,String htd,String fbgc,XSSFWorkbook wb) throws ParseException {
        XSSFCellStyle cellstyle = JjgFbgcCommonUtils.dBtoExcelUtils(wb);
        //表格数据填写
        XSSFSheet sheet = wb.getSheet("原始数据");
        //获取一下检测时间
        Date jcsj = data.get(0).getJcsj();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        String testtime = simpleDateFormat.format(jcsj);
        int index = 0;
        int tableNum = 0;
        //填写表头
        fillTitleCellData(sheet,tableNum,proname,htd,fbgc);
        //给每个table填写表头
        for(JjgFbgcLjgcZdgqd row:data){
            //比较检测时间，拿到最新的检测时间
            testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
            if(index/10 == 2){
                sheet.getRow(tableNum*25+2).getCell(29).setCellValue(testtime);
                testtime = simpleDateFormat.format(row.getJcsj());
                tableNum ++;
                fillTitleCellData(sheet, tableNum,proname,htd,fbgc);
                index = 0;
            }
            if(index%10 == 0){
                sheet.getRow(tableNum*25 + 5 + index).getCell(0).setCellValue(row.getZh());
            }
            fillCommonCellData(sheet, tableNum, index+5, row,cellstyle);
            index++;
        }
        sheet.getRow(tableNum*25+2).getCell(29).setCellValue(testtime);
        return true;
    }


    public void fillTitleCellData(XSSFSheet sheet, int tableNum, String proname,String htd,String fbgc) {
        if(sheet.getRow(tableNum*25+1) == null || sheet.getRow(tableNum*25+1).getCell(2) == null){
            return;
        }
        sheet.getRow(tableNum*25+1).getCell(2).setCellValue(proname);
        sheet.getRow(tableNum*25+1).getCell(29).setCellValue(htd);
        sheet.getRow(tableNum*25+2).getCell(2).setCellValue(fbgc);
    }


    public void fillCommonCellData(XSSFSheet sheet, int tableNum, int index,JjgFbgcLjgcZdgqd row, XSSFCellStyle cellstyle) {
        //sheet.getRow(tableNum*25+index).getCell(0).setCellValue(row.getZh());
        sheet.getRow(tableNum*25+index).getCell(1).setCellValue(row.getBw1());
        sheet.getRow(tableNum*25+index).getCell(2).setCellValue(row.getBw2());

        if(!"".equals(row.getCdz1()))
            sheet.getRow(tableNum*25+index).getCell(3).setCellValue(Double.valueOf(row.getCdz1()).intValue());
        if(!"".equals(row.getZ2()))
            sheet.getRow(tableNum*25+index).getCell(4).setCellValue(Double.valueOf(row.getZ2()).intValue());
        if(!"".equals(row.getZ3()))
            sheet.getRow(tableNum*25+index).getCell(5).setCellValue(Double.valueOf(row.getZ3()).intValue());
        if(!"".equals(row.getZ4()))
            sheet.getRow(tableNum*25+index).getCell(6).setCellValue(Double.valueOf(row.getZ4()).intValue());
        if(!"".equals(row.getZ5()))
            sheet.getRow(tableNum*25+index).getCell(7).setCellValue(Double.valueOf(row.getZ5()).intValue());
        if(!"".equals(row.getZ6()))
            sheet.getRow(tableNum*25+index).getCell(8).setCellValue(Double.valueOf(row.getZ6()).intValue());
        if(!"".equals(row.getZ7()))
            sheet.getRow(tableNum*25+index).getCell(9).setCellValue(Double.valueOf(row.getZ7()).intValue());
        if(!"".equals(row.getZ8()))
            sheet.getRow(tableNum*25+index).getCell(10).setCellValue(Double.valueOf(row.getZ8()).intValue());
        if(!"".equals(row.getZ9()))
            sheet.getRow(tableNum*25+index).getCell(11).setCellValue(Double.valueOf(row.getZ9()).intValue());
        if(!"".equals(row.getZ10()))
            sheet.getRow(tableNum*25+index).getCell(12).setCellValue(Double.valueOf(row.getZ10()).intValue());
        if(!"".equals(row.getZ11()))
            sheet.getRow(tableNum*25+index).getCell(13).setCellValue(Double.valueOf(row.getZ11()).intValue());
        if(!"".equals(row.getZ12()))
            sheet.getRow(tableNum*25+index).getCell(14).setCellValue(Double.valueOf(row.getZ12()).intValue());
        if(!"".equals(row.getZ13()))
            sheet.getRow(tableNum*25+index).getCell(15).setCellValue(Double.valueOf(row.getZ13()).intValue());
        if(!"".equals(row.getZ14()))
            sheet.getRow(tableNum*25+index).getCell(16).setCellValue(Double.valueOf(row.getZ14()).intValue());
        if(!"".equals(row.getZ15()))
            sheet.getRow(tableNum*25+index).getCell(17).setCellValue(Double.valueOf(row.getZ15()).intValue());
        if(!"".equals(row.getZ16()))
            sheet.getRow(tableNum*25+index).getCell(18).setCellValue(Double.valueOf(row.getZ16()).intValue());
        if("水平".equals(row.getHtjd())){
            sheet.getRow(tableNum*25+index).getCell(20).setCellValue(row.getHtjd());//U回弹角度
        }
        else{
            sheet.getRow(tableNum*25+index).getCell(20).setCellValue(Double.valueOf(row.getHtjd()).intValue());//U回弹角度
        }
        sheet.getRow(tableNum*25+index).getCell(23).setCellValue(row.getJzm());//X浇筑面
        sheet.getRow(tableNum*25+index).getCell(26).setCellValue(Double.valueOf(row.getThsd()));//AA碳化深度
        sheet.getRow(tableNum*25+index).getCell(28).setCellValue(row.getSfbs());//AC是否泵送
        sheet.getRow(tableNum*25+index).getCell(31).setCellValue(Double.valueOf(row.getSjqd()).intValue());//AF设计强度
        sheet.getRow(tableNum*25+index).getCell(31).setCellStyle(cellstyle);
    }



    public double getjdxzz(XSSFSheet sheet, double T, String U) {

        double res = 0.0;
        int rowNum = new BigDecimal(T*10).setScale(0, BigDecimal.ROUND_HALF_UP).intValue() - 200 + 3 - 1;

        int cellNum = 0;
        if(U.contains(".")){
            U = U.substring(0, U.indexOf("."));
        }
        switch (U){
            case "Rm":cellNum = 14;break;
            case "90":cellNum = 15;break;
            case "60":cellNum = 16;break;
            case "45":cellNum = 17;break;
            case "30":cellNum = 18;break;
            case "-30":cellNum = 19;break;
            case "-45":cellNum = 20;break;
            case "-60":cellNum = 21;break;
            case "-90":cellNum = 22;break;
            case "水平":cellNum = 23;break;
            default:break;
        }
        res = Double.valueOf(String.format("%.3f", sheet.getRow(rowNum).getCell(cellNum).getNumericCellValue()));
        return res;
    }

    public double getjzmxzz(XSSFSheet sheet, double W, String X) {
        double res =  0.0;
        int rowNum = new BigDecimal(W*10).setScale(0, BigDecimal.ROUND_HALF_UP).intValue() - 200 + 3 - 1;
        int cellNum = 0;
        if(X.contains(".")){
            X = X.substring(0, X.indexOf("."));
        }
        switch (X){
            case "表面":cellNum = 25;break;
            case "底面":cellNum = 26;break;
            case "侧面":cellNum = 27;break;
            default:break;
        }

        System.out.println("W= "+W+" X= "+X+" rowNum ="+rowNum+" cellNum= "+cellNum);
        res = Double.valueOf(String.format("%.3f", sheet.getRow(rowNum).getCell(cellNum).getNumericCellValue()));
        return res;
    }

    public double getthsdhsz(XSSFSheet sheet, double Z, String AA) {
        double res = 0.0;
        int rowNum = new BigDecimal(Z*10).setScale(0, BigDecimal.ROUND_HALF_UP).intValue() - 200 + 5 - 1;
        int cellNum = 0;
        if(!AA.contains("≥")){
            AA = String.format("%.1f", Double.valueOf(AA));
        }

        switch (AA){
            case "0.0":cellNum = 1;break;
            case "0.5":cellNum = 2;break;
            case "1.0":cellNum = 3;break;
            case "1.5":cellNum = 4;break;
            case "2.0":cellNum = 5;break;
            case "2.5":cellNum = 6;break;
            case "3.0":cellNum = 7;break;
            case "3.5":cellNum = 8;break;
            case "4.0":cellNum = 9;break;
            case "4.5":cellNum = 10;break;
            case "5.0":cellNum = 11;break;
            case "5.5":cellNum = 12;break;
            case "≥6":cellNum = 13;break;
            default:break;
        }

        System.out.println("rowNum ="+rowNum+" cellNum= "+cellNum+" Z= "+Z+" AA= "+AA);

        String temp="";
        if(sheet.getRow(rowNum).getCell(cellNum).toString().contains(">"))
        {
            System.out.println("temp="+sheet.getRow(rowNum).getCell(cellNum).toString());
            temp=sheet.getRow(rowNum).getCell(cellNum).toString();
            temp=temp.substring(1,temp.length());
            System.out.println("temp="+temp);
            res = Double.valueOf(String.format("%.3f", Double.parseDouble(temp)));
        }
        else
        {
            res = Double.valueOf(String.format("%.3f", sheet.getRow(rowNum).getCell(cellNum).getNumericCellValue()));
        }

        return res;
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
}
