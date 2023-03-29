package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLjgcLjwc;
import glgc.jjgys.model.project.JjgFbgcQlgcQmhp;
import glgc.jjgys.model.project.JjgFbgcQlgcSbTqd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.qlgc.JjgFbgcQlgcQmhpVo;
import glgc.jjgys.model.projectvo.qlgc.JjgFbgcQlgcSbTqdVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcQlgcQmhpMapper;
import glgc.jjgys.system.service.JjgFbgcQlgcQmhpService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-03-20
 */
@Service
public class JjgFbgcQlgcQmhpServiceImpl extends ServiceImpl<JjgFbgcQlgcQmhpMapper, JjgFbgcQlgcQmhp> implements JjgFbgcQlgcQmhpService {

    @Autowired
    private JjgFbgcQlgcQmhpMapper jjgFbgcQlgcQmhpMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        List<Map<String,Object>> qlmclist = jjgFbgcQlgcQmhpMapper.selectqlmc(proname,htd,fbgc);
        if (qlmclist.size()>0){
            for (Map<String, Object> m : qlmclist)
            {
                for (String k : m.keySet()){
                    String qlmc = m.get(k).toString();
                    List<Map<String,Object>> zhlist = jjgFbgcQlgcQmhpMapper.selectzh(proname,htd,fbgc,qlmc);
                    String zh = zhlist.get(0).get("zh").toString();
                    String substring = zh.substring(0, 1);
                    if (substring.equals("Z") || substring.equals("Y")){
                        DBtoExcelZYql(proname,htd,fbgc,qlmc);
                    }else {
                        DBtoExcelql(proname,htd,fbgc,qlmc);
                    }

                }
            }
        }
    }

    /**
     * 桩号是不分左右幅的
     * @param proname
     * @param htd
     * @param fbgc
     * @param qlmc
     */
    private void DBtoExcelql(String proname, String htd, String fbgc, String qlmc) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        QueryWrapper<JjgFbgcQlgcQmhp> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.like("qlmc",qlmc);
        wrapper.orderByAsc("zh");
        List<JjgFbgcQlgcQmhp> data = jjgFbgcQlgcQmhpMapper.selectList(wrapper);
        String lmlx = data.get(0).getLmlx();
        String lx = lmlx.substring(0, 2);
        String sheetname;
        if (lx.equals("水泥")){
            sheetname="混凝土桥面左幅";
        }else {
            sheetname="沥青桥面左幅";

        }
        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"35桥面横坡-"+qlmc+".xlsx");
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
            File directory = new File("src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String name = "横坡.xlsx";
            String path = reportPath + File.separator + name;
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            createTable(gettableNumAll(data.size()),wb,sheetname);
            if(DBtoExcel(data,wb,sheetname,qlmc)){
                calculateSheet(wb,sheetname);
                for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                    JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
                }
                JjgFbgcCommonUtils.deleteEmptySheets(wb);
                FileOutputStream fileOut = new FileOutputStream(f);
                wb.write(fileOut);
                fileOut.flush();
                fileOut.close();
            }
            out.close();
            wb.close();
        }


    }

    private void calculateSheet(XSSFWorkbook wb, String sheetname) {
        XSSFSheet sheet = wb.getSheet(sheetname);
        XSSFRow row = null;
        boolean flag = false;
        String name = "";
        ArrayList<XSSFRow> startRowList = new ArrayList<XSSFRow>();
        ArrayList<XSSFRow> endRowList = new ArrayList<XSSFRow>();
        sheet.getRow(0).getCell(12).setCellValue("合计");
        sheet.getRow(0).getCell(13).setCellValue("总点数");
        sheet.getRow(0).getCell(14).setCellFormula("SUM("
                +sheet.getRow(5).getCell(14).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).getCell(14).getReference()+")");
        sheet.getRow(0).getCell(15).setCellValue("合格点数");
        sheet.getRow(0).getCell(16).setCellFormula("SUM("
                +sheet.getRow(5).getCell(16).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).getCell(16).getReference()+")");
        sheet.getRow(0).getCell(17).setCellValue("合格率");
        sheet.getRow(0).getCell(18).setCellFormula(
                sheet.getRow(0).getCell(16).getReference()+"/"
                        +sheet.getRow(0).getCell(14).getReference()+"*100");
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if (flag && !"".equals(row.getCell(0).toString()) && row.getCell(0).toString().contains("质量鉴定表")) {
                flag = false;
            }
            if(flag && !"".equals(row.getCell(2).toString())){
                row.getCell(5).setCellFormula(row.getCell(4).getReference()+"-"
                        +row.getCell(3).getReference());//F=E7-D7
                row.getCell(7).setCellFormula(row.getCell(5).getReference()+"/"
                        +row.getCell(6).getReference()+"*100");//H=F7/G7*100
                row.getCell(10).setCellFormula("IF((("
                        +row.getCell(7).getReference()+"+"
                        +row.getCell(9).getReference()+")>="
                        +row.getCell(8).getReference()+")*AND("
                        +row.getCell(8).getReference()+">=("
                        +row.getCell(7).getReference()+"-"
                        +row.getCell(9).getReference()+")),\"√\",\"\")");//K=IF(((H7+0.3)>=I7)*AND(I7>=(H7-0.3)),"√","")
                row.getCell(11).setCellFormula("IF((("
                        +row.getCell(7).getReference()+"+"
                        +row.getCell(9).getReference()+")>="
                        +row.getCell(8).getReference()+")*AND("
                        +row.getCell(8).getReference()+">=("
                        +row.getCell(7).getReference()+"-"
                        +row.getCell(9).getReference()+")),\"\",\"×\")");//L=IF(((H7+0.3)>=I7)*AND(I7>=(H7-0.3)),"","×")
                row.getCell(12).setCellFormula(row.getCell(7).getReference()+"-"
                        +row.getCell(8).getReference());
            }
            if ("桩  号".equals(row.getCell(0).toString())) {
                i += 1;
                flag = true;
                if(!name.equals(sheet.getRow(i-3).getCell(8).toString()) && !"".equals(name)){
                    String column14 = "";
                    String column16 = "";
                    for(int r = 0;r < startRowList.size(); r++){
                        column14 += startRowList.get(r).getCell(7).getReference()+":"+endRowList.get(r).getCell(7).getReference()+",";
                        column16 += "COUNTIF("+startRowList.get(r).getCell(10).getReference()+":"+endRowList.get(r).getCell(10).getReference()+",\"√\")+";
                    }
                    startRowList.get(0).createCell(14).setCellFormula("COUNT("+column14.substring(0, column14.lastIndexOf(','))+")");
                    startRowList.get(0).createCell(16).setCellFormula(column16.substring(0, column16.lastIndexOf('+')));//=COUNTIF(AG6:AG50,"√")
                    startRowList.get(0).createCell(18).setCellFormula(startRowList.get(0).getCell(16).getReference()+"*100/"
                            +startRowList.get(0).getCell(14).getReference());//=COUNTIF(AH6:AH50,"×")

                    startRowList.clear();
                    endRowList.clear();
                }
                startRowList.add(sheet.getRow(i+1));
                endRowList.add(sheet.getRow(i+29));
                name = sheet.getRow(i-3).getCell(8).toString();
            }
        }
        String column14 = "";
        String column16 = "";
        for(int r = 0;r < startRowList.size(); r++){
            column14 += startRowList.get(r).getCell(7).getReference()+":"+endRowList.get(r).getCell(7).getReference()+",";
            column16 += "COUNTIF("+startRowList.get(r).getCell(10).getReference()+":"+endRowList.get(r).getCell(10).getReference()+",\"√\")+";
        }
        startRowList.get(0).createCell(14).setCellFormula("COUNT("+column14.substring(0, column14.lastIndexOf(','))+")");
        startRowList.get(0).createCell(16).setCellFormula(column16.substring(0, column16.lastIndexOf('+')));//=COUNTIF(AG6:AG50,"√")
        startRowList.get(0).createCell(18).setCellFormula(startRowList.get(0).getCell(16).getReference()+"*100/"
                +startRowList.get(0).getCell(14).getReference());//=COUNTIF(AH6:AH50,"×")

        startRowList.clear();
        endRowList.clear();
    }

    private boolean DBtoExcel(List<JjgFbgcQlgcQmhp> data, XSSFWorkbook wb,String sheetname,String qlmc) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        XSSFCellStyle cellstyle = JjgFbgcCommonUtils.dBtoExcelUtils(wb);
        String proname = data.get(0).getProname();
        String htd = data.get(0).getHtd();
        String fbgc = data.get(0).getFbgc();
        XSSFSheet sheet = wb.getSheet(sheetname);
        Date jcsj = data.get(0).getJcsj();
        String testtime = simpleDateFormat.format(jcsj);
        int index = 0;
        int tableNum = 0;
        fillTitleCellData(sheet,tableNum,proname,htd,fbgc,data.get(0).getLmlx(),qlmc);
        for(JjgFbgcQlgcQmhp row:data){
            //比较检测时间，拿到最新的检测时间
            testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
            if(index/29 == 1){
                sheet.getRow(tableNum*35+3).getCell(8).setCellValue(testtime);
                testtime = simpleDateFormat.format(row.getJcsj());
                tableNum ++;
                fillTitleCellData(sheet,tableNum,proname,htd,fbgc,row.getLmlx(),qlmc);
                index = 0;
            }
            //填写中间下方的普通单元格
            fillCommonCellData(sheet, tableNum, index+6, row,cellstyle);
            index++;
        }
        sheet.getRow(tableNum*35+3).getCell(8).setCellValue(testtime);
        return true;


    }

    private void fillCommonCellData(XSSFSheet sheet, int tableNum, int index, JjgFbgcQlgcQmhp row, XSSFCellStyle cellstyle) {
        sheet.getRow(tableNum*35+index).getCell(0).setCellValue(row.getZh());
        sheet.getRow(tableNum*35+index).getCell(2).setCellValue(row.getWz());
        sheet.getRow(tableNum*35+index).getCell(3).setCellValue(Double.parseDouble(row.getQsds()));
        sheet.getRow(tableNum*35+index).getCell(4).setCellValue(Double.parseDouble(row.getHsds()));
        sheet.getRow(tableNum*35+index).getCell(6).setCellValue(Double.parseDouble(row.getLength()));
        sheet.getRow(tableNum*35+index).getCell(8).setCellValue(Double.parseDouble(row.getSjz()));
        sheet.getRow(tableNum*35+index).getCell(9).setCellValue(Double.parseDouble(row.getYxps()));
    }

    private void fillTitleCellData(XSSFSheet sheet, int tableNum, String proname, String htd, String fbgc,String lmlx,String qlmc) {
        sheet.getRow(tableNum*35+1).getCell(2).setCellValue(proname);
        sheet.getRow(tableNum*35+1).getCell(8).setCellValue(htd);
        sheet.getRow(tableNum*35+2).getCell(2).setCellValue("桥梁工程");
        sheet.getRow(tableNum*35+2).getCell(8).setCellValue(fbgc+"("+qlmc+")");
        sheet.getRow(tableNum*35+3).getCell(2).setCellValue(lmlx);

    }

    private void createTable(int tableNum, XSSFWorkbook wb,String sheetname) {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            RowCopy.copyRows(wb, sheetname, sheetname, 0, 34, i* 35);
        }
        if(record >= 1){
            wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 11, 0,(record) * 35-1);
        }
    }

    private int gettableNumAll(int size) {
        return size%29 ==0 ? size/29 : size/29+1;
    }

    /**
     * 桩号是分左右幅的
     * @param proname
     * @param htd
     * @param fbgc
     * @param qlmc
     * @throws IOException
     */
    private void DBtoExcelZYql(String proname, String htd, String fbgc, String qlmc) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        QueryWrapper<JjgFbgcQlgcQmhp> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.like("qlmc",qlmc);
        wrapper.like("wz","左车道");
        wrapper.orderByAsc("zh");
        List<JjgFbgcQlgcQmhp> zdata = jjgFbgcQlgcQmhpMapper.selectList(wrapper);
        QueryWrapper<JjgFbgcQlgcQmhp> wrapper2=new QueryWrapper<>();
        wrapper2.like("proname",proname);
        wrapper2.like("htd",htd);
        wrapper2.like("fbgc",fbgc);
        wrapper2.like("qlmc",qlmc);
        wrapper2.like("wz","右车道");
        wrapper2.orderByAsc("zh");
        List<JjgFbgcQlgcQmhp> ydata = jjgFbgcQlgcQmhpMapper.selectList(wrapper2);
        String lmlx = zdata.get(0).getLmlx();
        String lx = lmlx.substring(0, 2);
        String shname;
        if (lx.equals("水泥")){
            shname="混凝土桥面";
        }else {
            shname="沥青桥面";

        }
        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"35桥面横坡-"+qlmc+".xlsx");
        //存放鉴定表的目录
        File fdir = new File(filepath + File.separator + proname + File.separator + htd);
        if (!fdir.exists()) {
            //创建文件根目录
            fdir.mkdirs();
        }
        File directory = new File("src/main/resources/static");
        String reportPath = directory.getCanonicalPath();
        String name = "横坡.xlsx";
        String path = reportPath + File.separator + name;
        Files.copy(Paths.get(path), new FileOutputStream(f));
        FileInputStream out = new FileInputStream(f);
        wb = new XSSFWorkbook(out);
        createTable(gettableNumAll(zdata.size()),wb,shname+"左幅");
        createTable(gettableNumAll(ydata.size()),wb,shname+"右幅");
        DBtoExcel(zdata,wb,shname+"左幅",qlmc);
        DBtoExcel(ydata,wb,shname+"右幅",qlmc);
        calculateSheet(wb,shname+"左幅");
        calculateSheet(wb,shname+"右幅");
        for (int j = 0; j < wb.getNumberOfSheets(); j++) {
            JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
        }
        JjgFbgcCommonUtils.deleteEmptySheets(wb);
        FileOutputStream fileOut = new FileOutputStream(f);
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
        out.close();
        wb.close();
    }


    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) {
        return null;
    }

    @Override
    public void export(HttpServletResponse response) {
        String fileName = "桥面横坡实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcQlgcQmhpVo()).finish();
    }

    @Override
    public void importqmhp(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcQlgcQmhpVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcQlgcQmhpVo>(JjgFbgcQlgcQmhpVo.class) {
                                @Override
                                public void handle(List<JjgFbgcQlgcQmhpVo> dataList) {
                                    for(JjgFbgcQlgcQmhpVo fbgcQlgcQmhpVo: dataList)
                                    {
                                        JjgFbgcQlgcQmhp fbgcQlgcQmhp = new JjgFbgcQlgcQmhp();
                                        BeanUtils.copyProperties(fbgcQlgcQmhpVo,fbgcQlgcQmhp);
                                        fbgcQlgcQmhp.setCreatetime(new Date());
                                        fbgcQlgcQmhp.setProname(commonInfoVo.getProname());
                                        fbgcQlgcQmhp.setHtd(commonInfoVo.getHtd());
                                        fbgcQlgcQmhp.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcQlgcQmhpMapper.insert(fbgcQlgcQmhp);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }
}
