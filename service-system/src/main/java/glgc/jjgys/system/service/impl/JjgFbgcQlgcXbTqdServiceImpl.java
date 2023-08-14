package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcQlgcXbTqd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.qlgc.JjgFbgcQlgcXbTqdVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcQlgcXbTqdMapper;
import glgc.jjgys.system.service.JjgFbgcLjgcZdgqdService;
import glgc.jjgys.system.service.JjgFbgcQlgcXbTqdService;
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
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
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
public class JjgFbgcQlgcXbTqdServiceImpl extends ServiceImpl<JjgFbgcQlgcXbTqdMapper, JjgFbgcQlgcXbTqd> implements JjgFbgcQlgcXbTqdService {

    @Autowired
    private JjgFbgcQlgcXbTqdMapper jjgFbgcQlgcXbTqdMapper;

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
        QueryWrapper<JjgFbgcQlgcXbTqd> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("qlmc","bw1");
        List<JjgFbgcQlgcXbTqd> data = jjgFbgcQlgcXbTqdMapper.selectList(wrapper);

        //List<Map<String,Object>> selectnum = jjgFbgcQlgcSbTqdMapper.selectnum(proname, htd, fbgc);
        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"25桥梁下部墩台砼强度.xlsx");
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
            String name = "涵洞砼强度.xlsx";
            String path = reportPath + File.separator + name;
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            createTable(gettableNum(data.size()),wb);
            if(DBtoExcel(data,wb)){
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

    private void calculateSheet(XSSFWorkbook wb) {
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
                row.getCell(21).setCellValue(jjgFbgcLjgcZdgqdService.getjdxzz(wb.getSheet("修正数据"), row.getCell(19).getNumericCellValue(), row.getCell(20).toString()));
                row.getCell(21).getCellStyle().setLocked(true);

                row.getCell(22).setCellFormula(
                        "IF(" + row.getCell(3).getReference()
                                + "=\"\",\"\",ROUND("
                                + row.getCell(19).getReference() + "+"
                                + row.getCell(21).getReference() + ",1))");// W=IF(D6="","",ROUND(T6+V6,1))
                evaluator.evaluateInCell(row.getCell(22));
                row.getCell(22).getCellStyle().setLocked(true);
                // Y=IF(D6="","",INDEX(修正数据!$Y$2:$AB$403,MATCH(原始数据!W6,修正数据!$Y$2:$Y$403,0),MATCH(原始数据!X6,修正数据!$Y$2:$AB$2,0)))
                row.getCell(24).setCellValue(jjgFbgcLjgcZdgqdService.getjzmxzz(wb.getSheet("修正数据"), row.getCell(22).getNumericCellValue(), row.getCell(23).toString()));
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
                if (jjgFbgcLjgcZdgqdService.getthsdhsz(wb.getSheet("修正数据"), row.getCell(25).getNumericCellValue(), row.getCell(26).toString()) == 60) {
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
                row.getCell(27).setCellValue(jjgFbgcLjgcZdgqdService.getthsdhsz(wb.getSheet("修正数据"), row.getCell(25).getNumericCellValue(), row.getCell(26).toString()));
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
            if(flag && !"".equals(row.getCell(0).toString()) && !name.equals(jjgFbgcLjgcZdgqdService.getBridgeName(row.getCell(0).toString()))){
                //计算这座桥的总计
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
                startRowList.get(0).getCell(36).setCellFormula(startRowList.get(0).getCell(35).getReference()+"*100/"
                        +startRowList.get(0).getCell(34).getReference());//=COUNTIF(AH6:AH50,"×")
                startRowList.clear();
                endRowList.clear();
                startRowList.add(sheet.getRow(i));
                endRowList.add(sheet.getRow(i));
                name = jjgFbgcLjgcZdgqdService.getBridgeName(row.getCell(0).toString());
            } else if(flag && !"".equals(row.getCell(0). toString()) && row.getCell(0).toString()!=null && name.equals(jjgFbgcLjgcZdgqdService.getBridgeName(row.getCell(0).toString()))){
                startRowList.add(sheet.getRow(i));
                endRowList.add(sheet.getRow(i));
            }
            if ("桩号/\n结构名称".equals(row.getCell(0).toString())) {
                flag = true;
                i++;
                rowstart = sheet.getRow(i+1);
                //计算每座桥的总点数，合格点数，不合格点数
                if(!name.equals(jjgFbgcLjgcZdgqdService.getBridgeName(sheet.getRow(i+1).getCell(0).toString())) && !"".equals(name)){
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
                name = jjgFbgcLjgcZdgqdService.getBridgeName(sheet.getRow(i+1).getCell(0).toString());
                rowend = sheet.getRow(i+1);
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
            column29 += "COUNT("+startRowList.get(r).getCell(29).getReference()+":"+endRowList.get(r).getCell(29).getReference()+")+";
            column32 += "COUNTIF("+startRowList.get(r).getCell(32).getReference()+":"+endRowList.get(r).getCell(32).getReference()+",\"√\")+";
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

    private boolean DBtoExcel(List<JjgFbgcQlgcXbTqd> data,XSSFWorkbook wb) throws ParseException {
        String proname = data.get(0).getProname();
        String htd = data.get(0).getHtd();
        String fbgc = data.get(0).getFbgc();
        XSSFCellStyle cellstyle = JjgFbgcCommonUtils.dBtoExcelUtils(wb);
        XSSFSheet sheet = wb.getSheet("原始数据");
        Date jcsj = data.get(0).getJcsj();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        String testtime = simpleDateFormat.format(jcsj);
        int index = 0;
        int tableNum = 0;
        //填写表头
        fillTitleCellData(sheet,tableNum,proname,htd,fbgc);
        for(JjgFbgcQlgcXbTqd row:data){
            //比较检测时间，拿到最新的检测时间
            testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
            if(index/10 == 2){
                sheet.getRow(tableNum*25+2).getCell(29).setCellValue(testtime);
                testtime = simpleDateFormat.format(row.getJcsj());
                tableNum ++;
                fillTitleCellData(sheet,tableNum,proname,htd,fbgc);
                index = 0;
            }
            //填写中间下方的普通单元格
            fillCommonCellData(sheet, tableNum, index+5, row,cellstyle);
            index++;
        }
        sheet.getRow(tableNum*25+2).getCell(29).setCellValue(testtime);
        return true;
    }

    public void fillTitleCellData(XSSFSheet sheet, int tableNum,String proname,String htd,String fbgc) {
        if(sheet.getRow(tableNum*25+1) == null || sheet.getRow(tableNum*25+1).getCell(2) == null){
            return;
        }
        sheet.getRow(tableNum*25+1).getCell(2).setCellValue(proname);
        sheet.getRow(tableNum*25+1).getCell(29).setCellValue(htd);
        sheet.getRow(tableNum*25+2).getCell(2).setCellValue(fbgc);
    }

    public void fillCommonCellData(XSSFSheet sheet, int tableNum, int index,JjgFbgcQlgcXbTqd row, XSSFCellStyle cellstyle) {
        sheet.getRow(tableNum*25+index).getCell(0).setCellValue(row.getQlmc());
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

    private void createTable(int tableNum,XSSFWorkbook wb) {
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
        if(record >= 1)
            wb.setPrintArea(wb.getSheetIndex("原始数据"), 0, 33, 0, record * 25-1);
    }

    private int gettableNum(int size) {
        return size%20 ==0 ? size/20 : size/20+1;
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        //String fbgc = commonInfoVo.getFbgc();
        String title = "混凝土强度质量鉴定表（回弹法）";
        String sheetname = "原始数据";
        List<Map<String,Object>> mapList = new ArrayList<>();
        Map<String,Object> jgmap = new HashMap<>();
        //获取鉴定表文件
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"25桥梁下部墩台砼强度.xlsx");
        if(!f.exists()){
            return null;
        }else {
            //创建工作簿
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
            XSSFSheet slSheet = wb.getSheet(sheetname);
            slSheet.getRow(2).getCell(34).setCellType(CellType.STRING);
            slSheet.getRow(2).getCell(35).setCellType(CellType.STRING);
            slSheet.getRow(2).getCell(36).setCellType(CellType.STRING);
            String bt = slSheet.getRow(0).getCell(0).getStringCellValue();//混凝土强度质量鉴定表（回弹法）
            String xmname = slSheet.getRow(1).getCell(2).getStringCellValue();//陕西高速
            String htdname = slSheet.getRow(1).getCell(29).getStringCellValue();//LJ-1
            if (proname.equals(xmname) && htd.equals(htdname)) {
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
    }

    @Override
    public void importqlxbtqd(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcQlgcXbTqdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcQlgcXbTqdVo>(JjgFbgcQlgcXbTqdVo.class) {
                                @Override
                                public void handle(List<JjgFbgcQlgcXbTqdVo> dataList) {
                                    for(JjgFbgcQlgcXbTqdVo xbTqdVo: dataList)
                                    {
                                        JjgFbgcQlgcXbTqd fbgcQlgcXbTqd = new JjgFbgcQlgcXbTqd();
                                        BeanUtils.copyProperties(xbTqdVo,fbgcQlgcXbTqd);
                                        fbgcQlgcXbTqd.setCreatetime(new Date());
                                        fbgcQlgcXbTqd.setProname(commonInfoVo.getProname());
                                        fbgcQlgcXbTqd.setHtd(commonInfoVo.getHtd());
                                        fbgcQlgcXbTqd.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcQlgcXbTqdMapper.insert(fbgcQlgcXbTqd);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }


    }

    @Override
    public void exportqlxbtqd(HttpServletResponse response) {
        String fileName = "01桥梁下部墩台砼强度实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcQlgcXbTqdVo()).finish();

    }

    @Override
    public List<Map<String, Object>> lookjg(CommonInfoVo commonInfoVo) {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String sheetname = "原始数据";
        DecimalFormat decf = new DecimalFormat("0.##");
        //获取鉴定表文件
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "25桥梁下部墩台砼强度.xlsx");
        if (!f.exists()) {
            return null;
        }else {
            List<Map<String, Object>> data = new ArrayList<>();

            try (FileInputStream fis = new FileInputStream(f)) {
                Workbook workbook = WorkbookFactory.create(fis);
                Sheet sheet = workbook.getSheet(sheetname); // 假设数据在第一个工作表中

                Iterator<Row> rowIterator = sheet.iterator();
                int rowNum = 0;
                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    rowNum++;
                    if (rowNum > 4 && row !=null) {
                        Cell cell0 = row.getCell(0);
                        Cell cell34 = row.getCell(34);
                        Cell cell35 = row.getCell(35);
                        Cell cell31 = row.getCell(31);

                        if (cell0 != null && cell34 != null) {
                            String cellValue0 = cell0.getStringCellValue();
                            cell31.setCellType(CellType.STRING);
                            String cellValue31 = cell31.getStringCellValue();
                            String cellValue34 = String.valueOf(decf.format(cell34.getNumericCellValue()));
                            String cellValue35 = String.valueOf(decf.format(cell35.getNumericCellValue()));

                            if (!cellValue0.isEmpty() && !cellValue34.isEmpty()) {
                                Map map = new HashMap();
                                map.put("qlmc",cellValue0);
                                map.put("zds",cellValue34);
                                map.put("hgds",cellValue35);
                                map.put("sjqd",cellValue31);

                                data.add(map);
                            }


                        }
                    }

                }
            } catch (Exception e) {
                e.printStackTrace();
            }
            return data;

        }
    }

}
