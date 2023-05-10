package glgc.jjgys.system.service.impl;

import cn.hutool.core.io.resource.ClassPathResource;
import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLjgcHdgqd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.ljgc.JjgFbgcLjgcHdgqdVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcLjgcHdgqdMapper;
import glgc.jjgys.system.service.JjgFbgcLjgcHdgqdService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * <p>
 *  涵洞砼强度服务类
 * </p>
 *
 * @author wq
 * @since 2023-01-14
 */
@Service
public class JjgFbgcLjgcHdgqdServiceImpl extends ServiceImpl<JjgFbgcLjgcHdgqdMapper, JjgFbgcLjgcHdgqd> implements JjgFbgcLjgcHdgqdService {

    @Autowired
    private JjgFbgcLjgcHdgqdMapper jjgFbgcLjgcHdgqdMapper;

    private static XSSFWorkbook wb = null;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    /**
     * 查看涵洞砼强度鉴定结果
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    @Override
    public List<Map<String,Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        String title = "混凝土强度质量鉴定表（回弹法）";
        String sheetname = "原始数据";
        //获取鉴定表文件
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"08路基涵洞砼强度.xlsx");
        if(!f.exists()){
            return null;
        }else {
            Map<String,Object> map = new HashMap<>();
            map.put("proname",proname);
            map.put("title",title);
            map.put("htd",htd);
            map.put("fbgc",fbgc);
            map.put("f",f);
            map.put("sheetname",sheetname);
            List<Map<String, Object>> mapList = JjgFbgcCommonUtils.gettqdjcjg(map);
            return mapList;
        }
    }

    /**
     * 导出涵洞砼强度数据模板文件
     * @param response
     */
    @Override
    public void exporthdgqd(HttpServletResponse response) {
        String fileName = "08涵洞砼强度实测数据";
        String sheetName = "原始数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcLjgcHdgqdVo()).finish();

    }

    /**
     * 导入涵洞砼强度数据文件
     * @param file
     * @param commonInfoVo
     */
    @Override
    public void importhdgqd(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcLjgcHdgqdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcLjgcHdgqdVo>(JjgFbgcLjgcHdgqdVo.class) {
                                @Override
                                public void handle(List<JjgFbgcLjgcHdgqdVo> dataList) {
                                    for(JjgFbgcLjgcHdgqdVo hdgqdVo: dataList)
                                    {
                                        JjgFbgcLjgcHdgqd fbgcLjgcHdgqd = new JjgFbgcLjgcHdgqd();
                                        BeanUtils.copyProperties(hdgqdVo,fbgcLjgcHdgqd);
                                        fbgcLjgcHdgqd.setCreatetime(new Date());
                                        fbgcLjgcHdgqd.setProname(commonInfoVo.getProname());
                                        fbgcLjgcHdgqd.setHtd(commonInfoVo.getHtd());
                                        fbgcLjgcHdgqd.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcLjgcHdgqdMapper.insert(fbgcLjgcHdgqd);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }

    /**
     * 平均值：将测定值去掉3个最大的，去掉3个最小的，然后求平均值
     * 每个结构名称有10个数据点，每页有20个
     * 鉴定表的设计强度值前有C，固定不变的
     */
    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws Exception {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        System.out.println(proname+htd+fbgc);
        //获取数据
        QueryWrapper<JjgFbgcLjgcHdgqd> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("zh","bw1");
        List<JjgFbgcLjgcHdgqd> data = jjgFbgcLjgcHdgqdMapper.selectList(wrapper);
        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"08路基涵洞砼强度.xlsx");
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
            File directory = new File("service-system/src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String name = "涵洞砼强度.xlsx";
            String path =reportPath + File.separator+name;
            System.out.println(path);
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            createTable(gettableNum(data.size()));
            if(DBtoExcel(data,proname,htd,fbgc)){
                //设置公式,计算合格点数
                calculateSheet(wb.getSheet("原始数据"));
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

    //生成表格，并设置要打印的区域
    public void createTable(int tableNum) throws IOException {
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

    //根据数据库中的数据桩号，判断要生成几个表格
    public int gettableNum(int size){
        return size%20 ==0 ? size/20 : size/20+1;
    }

    public boolean DBtoExcel(List<JjgFbgcLjgcHdgqd> data,String proname,String htd,String fbgc) throws ParseException {
        XSSFCellStyle cellstyle = JjgFbgcCommonUtils.dBtoExcelUtils(wb);

        //表格数据填写
        XSSFSheet sheet = wb.getSheet("原始数据");
        Date jcsj = data.get(0).getJcsj();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        String testtime = simpleDateFormat.format(jcsj);
        int index = 0;
        int tableNum = 0;
        //填写表头
        fillTitleCellData(sheet,tableNum,proname,htd,fbgc);
        //给每个table填写表头
        for(JjgFbgcLjgcHdgqd row:data){
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

    //填写表头
    public void fillTitleCellData(XSSFSheet sheet, int tableNum,String proname,String htd,String fbgc) {
        if(sheet.getRow(tableNum*25+1) == null || sheet.getRow(tableNum*25+1).getCell(2) == null){
            return;
        }
        sheet.getRow(tableNum*25+1).getCell(2).setCellValue(proname);
        sheet.getRow(tableNum*25+1).getCell(29).setCellValue(htd);
        sheet.getRow(tableNum*25+2).getCell(2).setCellValue(fbgc);
    }


    //填写中间下方的普通单元格
    public void fillCommonCellData(XSSFSheet sheet, int tableNum, int index,JjgFbgcLjgcHdgqd row, XSSFCellStyle cellstyle) {
        sheet.getRow(tableNum*25+index).getCell(0).setCellValue(row.getZh());
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


    public void calculateSheet(XSSFSheet sheet) {
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
            if (flag && !"".equals(row.getCell(0).toString()) && row.getCell(0).toString().contains("质量鉴定表")) {
                /*
                 * 计算每张表的总点数，合格点数，不合格点数
                 */
                rowstart.createCell(34).setCellFormula("COUNT("   //COUNTIF(B4:L4,"<>""*")-COUNTBLANK(B4:L4)  sumproduct(n(len(A1:A100)>0))
                        +rowstart.getCell(26).getReference()+":"
                        +rowend.getCell(26).getReference()+")");
                rowstart.createCell(35).setCellFormula("COUNTIF("
                        +rowstart.getCell(32).getReference()+":"
                        +rowend.getCell(32).getReference()+",\"√\")");//=COUNTIF(AG6:AG50,"√")

                rowstart.createCell(36).setCellFormula("COUNTIF("
                        +rowstart.getCell(33).getReference()+":"
                        +rowend.getCell(33).getReference()+",\"×\")");//=COUNTIF(AH6:AH50,"×")
                flag = false;
            }
            // 可以计算
            if (row.getCell(3).getCellType() == Cell.CELL_TYPE_NUMERIC && flag) {
                int [] value = new int[16];
                for(int index = 0 ; index < 16; index ++){
                    value[index] = (int) row.getCell(3+index).getNumericCellValue();
                }
                System.out.println(value);
                Arrays.sort(value);
                String array = "";
                for(int index = 0 ; index < 10; index ++){
                    value[index] = (int) row.getCell(3+index).getNumericCellValue();
                    array += value[index+3]+",";
                }
                row.getCell(19).setCellFormula(
                        "IF(" + row.getCell(3).getReference()
                                + ">0,ROUND(AVERAGE("
                                + array.substring(0, array.lastIndexOf(","))
                                + "),1),\" \")");
                row.getCell(21).setCellFormula(
                        "IF("
                                + row.getCell(3).getReference()
                                + "=\"\",\"\",INDEX(修正数据!$O$2:$X$403,MATCH(原始数据!"
                                + row.getCell(19).getReference()
                                + ",修正数据!$O$2:$O$403,0),MATCH(原始数据!"
                                + row.getCell(20).getReference()
                                + ",修正数据!$O$2:$X$2,0)))");
                row.getCell(21).getCellStyle().setLocked(true);
                row.getCell(22).setCellFormula(
                        "IF(" + row.getCell(3).getReference()
                                + "=\"\",\"\",ROUND("
                                + row.getCell(19).getReference() + "+"
                                + row.getCell(21).getReference() + ",1))");// W=IF(D6="","",ROUND(T6+V6,1))
                row.getCell(22).getCellStyle().setLocked(true);
                row.getCell(24).setCellFormula(
                        "IF("
                                + row.getCell(3).getReference()
                                + "=\"\",\"\",INDEX(修正数据!$Y$2:$AB$403,MATCH(原始数据!"
                                + row.getCell(22).getReference()
                                + ",修正数据!$Y$2:$Y$403,0),MATCH(原始数据!"
                                + row.getCell(23).getReference()
                                + ",修正数据!$Y$2:$AB$2,0)))");
                row.getCell(24).getCellStyle().setLocked(true);
                row.getCell(25).setCellFormula(
                        "IF(" + row.getCell(22).getReference()
                                + "=\"\",\"\",ROUND("
                                + row.getCell(22).getReference() + "+"
                                + row.getCell(24).getReference() + ",1))");
                row.getCell(25).getCellStyle().setLocked(true);
                row.getCell(27).setCellFormula(
                        "IF("
                                + row.getCell(25).getReference()
                                + "=\"\",\"\",INDEX(修正数据!$A$4:$N$405,MATCH(原始数据!"
                                + row.getCell(25).getReference()
                                + ",修正数据!$A$4:$A$405,0),MATCH(原始数据!"
                                + row.getCell(26).getReference()
                                + ",修正数据!$A$4:$N$4,0)))");
                row.getCell(27).getCellStyle().setLocked(true);
                row.getCell(29).setCellFormula(
                        "IF(" + row.getCell(3).getReference()
                                + "=\"\",\"\",IF("
                                + row.getCell(28).getReference()
                                + "=\"是\",IF(0.034488*"
                                + row.getCell(25).getReference()
                                + "^1.94*10^(-0.0176*"
                                + row.getCell(26).getReference()
                                + ")>60,60.0,0.034488*"
                                + row.getCell(25).getReference()
                                + "^1.94*10^(-0.0176*"
                                + row.getCell(26).getReference() + ")),"
                                + row.getCell(27).getReference() + "))");
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
            if ("桩号/\n结构名称".equals(row.getCell(0).toString())) {
                flag = true;
                i++;
                rowstart = sheet.getRow(i+1);
                rowend = sheet.getRow(i+20);
            }

        }
        /*
         * 计算每张表的总点数，合格点数，不合格点数
         */
        rowstart.createCell(34).setCellFormula("COUNT("   //COUNTIF(B4:L4,"<>""*")-COUNTBLANK(B4:L4)  sumproduct(n(len(A1:A100)>0))
                +rowstart.getCell(26).getReference()+":"
                +rowend.getCell(26).getReference()+")");

        rowstart.createCell(35).setCellFormula("COUNTIF("
                +rowstart.getCell(32).getReference()+":"
                +rowend.getCell(32).getReference()+",\"√\")");//=COUNTIF(AG6:AG50,"√")
        rowstart.createCell(36).setCellFormula("COUNTIF("
                +rowstart.getCell(33).getReference()+":"
                +rowend.getCell(33).getReference()+",\"×\")");//=COUNTIF(AH6:AH50,"×")
        /*
         * 计算总点数，合格点数，合格率
         */
        sheet.getRow(2).createCell(34).setCellFormula("SUM("
                +sheet.getRow(5).getCell(34).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(34).getReference()+")");//总点数=SUM(AI5:AI50)
        sheet.getRow(2).createCell(35).setCellFormula("SUM("
                +sheet.getRow(5).getCell(35).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(35).getReference()+")");//合格点数=SUM(AJ5:AJ50)
        sheet.getRow(2).createCell(36).setCellFormula(sheet.getRow(2).getCell(35).getReference()+"*100/"
                +sheet.getRow(2).getCell(34).getReference());//合格率

    }
}
