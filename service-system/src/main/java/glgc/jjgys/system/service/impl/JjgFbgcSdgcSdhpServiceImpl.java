package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcSdgcSdhp;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.sdgc.JjgFbgcSdgcSdhpVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcSdgcSdhpMapper;
import glgc.jjgys.system.service.JjgFbgcSdgcSdhpService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.poi.ss.usermodel.CellType;
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
 * @since 2023-03-27
 */
@Service
public class JjgFbgcSdgcSdhpServiceImpl extends ServiceImpl<JjgFbgcSdgcSdhpMapper, JjgFbgcSdgcSdhp> implements JjgFbgcSdgcSdhpService {

    @Autowired
    private JjgFbgcSdgcSdhpMapper jjgFbgcSdgcSdhpMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        List<Map<String,Object>> sdmclist = jjgFbgcSdgcSdhpMapper.selectsdmc(proname,htd,fbgc);
        if (sdmclist.size()>0){
            for (Map<String, Object> m : sdmclist)
            {
                for (String k : m.keySet()){
                    String sdmc = m.get(k).toString();
                    List<Map<String,Object>> zhlist = jjgFbgcSdgcSdhpMapper.selectzh(proname,htd,fbgc,sdmc);
                    String zh = zhlist.get(0).get("zh").toString();
                    String substring = zh.substring(0, 1);
                    if (substring.equals("Z") || substring.equals("Y")){
                        DBtoExcelZYsd(proname,htd,fbgc,sdmc);
                    }else {
                        DBtoExcelsd(proname,htd,fbgc,sdmc);
                    }

                }
            }
        }

    }

    private void DBtoExcelsd(String proname, String htd, String fbgc, String sdmc) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        QueryWrapper<JjgFbgcSdgcSdhp> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.like("sdmc",sdmc);
        wrapper.orderByAsc("zh");
        List<JjgFbgcSdgcSdhp> data = jjgFbgcSdgcSdhpMapper.selectList(wrapper);
        String lmlx = data.get(0).getLmlx();
        String lx = lmlx.substring(0, 2);
        String sheetname;
        if (lx.equals("水泥")){
            sheetname="混凝土隧道左幅";
        }else {
            sheetname="沥青隧道左幅";

        }
        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"55隧道横坡-"+sdmc+".xlsx");
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
            String name = "横坡.xlsx";
            String path = reportPath + File.separator + name;
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            createTable(gettableNumAll(data.size()),wb,sheetname);
            if(DBtoExcel(data,wb,sheetname,sdmc)){
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

    private void DBtoExcelZYsd(String proname, String htd, String fbgc, String sdmc) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        QueryWrapper<JjgFbgcSdgcSdhp> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.like("sdmc",sdmc);
        wrapper.like("wz","左车道");
        wrapper.orderByAsc("zh");
        List<JjgFbgcSdgcSdhp> zdata = jjgFbgcSdgcSdhpMapper.selectList(wrapper);
        QueryWrapper<JjgFbgcSdgcSdhp> wrapper2=new QueryWrapper<>();
        wrapper2.like("proname",proname);
        wrapper2.like("htd",htd);
        wrapper2.like("fbgc",fbgc);
        wrapper2.like("sdmc",sdmc);
        wrapper2.like("wz","右车道");
        wrapper2.orderByAsc("zh");
        List<JjgFbgcSdgcSdhp> ydata = jjgFbgcSdgcSdhpMapper.selectList(wrapper2);
        String lmlx = zdata.get(0).getLmlx();
        String lx = lmlx.substring(0, 2);
        String shname;
        if (lx.equals("水泥")){
            shname="混凝土隧道";
        }else {
            shname="沥青隧道";

        }
        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"55隧道横坡-"+sdmc+".xlsx");
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
        DBtoExcel(zdata,wb,shname+"左幅",sdmc);
        DBtoExcel(ydata,wb,shname+"右幅",sdmc);
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

    private boolean DBtoExcel(List<JjgFbgcSdgcSdhp> data, XSSFWorkbook wb,String sheetname,String sdmc) throws ParseException {
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
        fillTitleCellData(sheet,tableNum,proname,htd,fbgc,data.get(0).getLmlx(),sdmc);
        for(JjgFbgcSdgcSdhp row:data){
            //比较检测时间，拿到最新的检测时间
            testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
            if(index/29 == 1){
                sheet.getRow(tableNum*35+3).getCell(8).setCellValue(testtime);
                testtime = simpleDateFormat.format(row.getJcsj());
                tableNum ++;
                fillTitleCellData(sheet,tableNum,proname,htd,fbgc,row.getLmlx(),sdmc);
                index = 0;
            }
            //填写中间下方的普通单元格
            fillCommonCellData(sheet, tableNum, index+6, row,cellstyle);
            index++;
        }
        sheet.getRow(tableNum*35+3).getCell(8).setCellValue(testtime);
        return true;


    }

    private void fillCommonCellData(XSSFSheet sheet, int tableNum, int index, JjgFbgcSdgcSdhp row, XSSFCellStyle cellstyle) {
        sheet.getRow(tableNum*35+index).getCell(0).setCellValue(row.getZh());
        sheet.getRow(tableNum*35+index).getCell(2).setCellValue(row.getWz());
        sheet.getRow(tableNum*35+index).getCell(3).setCellValue(Double.parseDouble(row.getQsds()));
        sheet.getRow(tableNum*35+index).getCell(4).setCellValue(Double.parseDouble(row.getHsds()));
        sheet.getRow(tableNum*35+index).getCell(6).setCellValue(Double.parseDouble(row.getLength()));
        sheet.getRow(tableNum*35+index).getCell(8).setCellValue(Double.parseDouble(row.getSjz()));
        sheet.getRow(tableNum*35+index).getCell(9).setCellValue(Double.parseDouble(row.getYxps()));
    }

    private void fillTitleCellData(XSSFSheet sheet, int tableNum, String proname, String htd, String fbgc,String lmlx,String sdmc) {
        sheet.getRow(tableNum*35+1).getCell(2).setCellValue(proname);
        sheet.getRow(tableNum*35+1).getCell(8).setCellValue(htd);
        sheet.getRow(tableNum*35+2).getCell(2).setCellValue("隧道工程");
        sheet.getRow(tableNum*35+2).getCell(8).setCellValue(fbgc+"("+sdmc+")");
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

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat(".00");
        DecimalFormat decf = new DecimalFormat("0.##");
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        List<Map<String, Object>> selectsdmc = selectsdmc(proname, htd, fbgc);

        List<Map<String, Object>> mapList = new ArrayList<>();

        for (int i=0;i<selectsdmc.size();i++){
            String sdmc = selectsdmc.get(i).get("sdmc").toString();
            File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "55隧道横坡-"+sdmc+".xlsx");
            if (!f.exists()) {
                return null;
            } else {
                List<Map<String,Object>> zhlist = jjgFbgcSdgcSdhpMapper.selectzh(proname,htd,fbgc,sdmc);
                String zh = zhlist.get(0).get("zh").toString();
                String lx = zhlist.get(0).get("lmlx").toString().substring(0,2);
                String lxname;
                if (lx.equals("水泥")){
                    lxname="混凝土隧道";
                }else {
                    lxname="沥青隧道";

                }
                String substring = zh.substring(0, 1);
                XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));

                if (substring.equals("Z") || substring.equals("Y")){
                    String sheetnameleft = lxname+"左幅";
                    String sheetnameright = lxname+"右幅";
                    XSSFSheet sheetleft = xwb.getSheet(sheetnameleft);
                    XSSFSheet sheetright = xwb.getSheet(sheetnameright);
                    XSSFCell xmnameleft = sheetleft.getRow(1).getCell(2);//项目名
                    XSSFCell htdnameleft = sheetleft.getRow(1).getCell(8);//合同段名
                    if (proname.equals(xmnameleft.toString()) && htd.equals(htdnameleft.toString())) {
                        sheetleft.getRow(0).getCell(14).setCellType(CellType.STRING);//总点数
                        sheetleft.getRow(0).getCell(16).setCellType(CellType.STRING);//合格点数
                        sheetleft.getRow(0).getCell(18).setCellType(CellType.STRING);//合格率

                        sheetright.getRow(0).getCell(14).setCellType(CellType.STRING);//总点数
                        sheetright.getRow(0).getCell(16).setCellType(CellType.STRING);//合格点数
                        sheetright.getRow(0).getCell(18).setCellType(CellType.STRING);//合格率

                        Map<String, Object> jgmap = new HashMap<>();
                        jgmap.put("检测项目",sdmc+sheetnameleft);
                        jgmap.put("总点数", decf.format(Double.valueOf(sheetleft.getRow(0).getCell(14).getStringCellValue())));
                        jgmap.put("合格点数", decf.format(Double.valueOf(sheetleft.getRow(0).getCell(16).getStringCellValue())));
                        jgmap.put("合格率", df.format(Double.valueOf(sheetleft.getRow(0).getCell(18).getStringCellValue())));

                        Map<String, Object> jgmapright = new HashMap<>();
                        jgmapright.put("检测项目",sdmc+sheetnameright);
                        jgmapright.put("总点数", decf.format(Double.valueOf(sheetright.getRow(0).getCell(14).getStringCellValue())));
                        jgmapright.put("合格点数", decf.format(Double.valueOf(sheetright.getRow(0).getCell(16).getStringCellValue())));
                        jgmapright.put("合格率", df.format(Double.valueOf(sheetright.getRow(0).getCell(18).getStringCellValue())));
                        mapList.add(jgmap);
                        mapList.add(jgmapright);
                    }
                }else {
                    String sheetname = lxname+"左幅";
                    XSSFSheet sheetl = xwb.getSheet(sheetname);
                    XSSFCell xmnameleft = sheetl.getRow(1).getCell(2);//项目名
                    XSSFCell htdnameleft = sheetl.getRow(1).getCell(8);//合同段名
                    if (proname.equals(xmnameleft.toString()) && htd.equals(htdnameleft.toString())) {
                        sheetl.getRow(0).getCell(14).setCellType(CellType.STRING);//总点数
                        sheetl.getRow(0).getCell(16).setCellType(CellType.STRING);//合格点数
                        sheetl.getRow(0).getCell(18).setCellType(CellType.STRING);//合格率
                        Map<String, Object> jgmap = new HashMap<>();
                        jgmap.put("检测项目",sdmc+lxname);
                        jgmap.put("总点数", decf.format(Double.valueOf(sheetl.getRow(0).getCell(14).getStringCellValue())));
                        jgmap.put("合格点数", decf.format(Double.valueOf(sheetl.getRow(0).getCell(16).getStringCellValue())));
                        jgmap.put("合格率", df.format(Double.valueOf(sheetl.getRow(0).getCell(18).getStringCellValue())));
                        mapList.add(jgmap);
                    }
                }
            }

        }
        return mapList;
    }

    @Override
    public void export(HttpServletResponse response) {
        String fileName = "15隧道横坡实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcSdgcSdhpVo()).finish();

    }

    @Override
    public void importsdhp(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcSdgcSdhpVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcSdgcSdhpVo>(JjgFbgcSdgcSdhpVo.class) {
                                @Override
                                public void handle(List<JjgFbgcSdgcSdhpVo> dataList) {
                                    for(JjgFbgcSdgcSdhpVo fbgcSdgcSdhpVo: dataList)
                                    {
                                        JjgFbgcSdgcSdhp fbgcSdgcSdhp = new JjgFbgcSdgcSdhp();
                                        BeanUtils.copyProperties(fbgcSdgcSdhpVo,fbgcSdgcSdhp);
                                        fbgcSdgcSdhp.setCreatetime(new Date());
                                        fbgcSdgcSdhp.setProname(commonInfoVo.getProname());
                                        fbgcSdgcSdhp.setHtd(commonInfoVo.getHtd());
                                        fbgcSdgcSdhp.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcSdgcSdhpMapper.insert(fbgcSdgcSdhp);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }

    @Override
    public List<Map<String, Object>> selectsdmc(String proname, String htd, String fbgc) {
        List<Map<String,Object>> sdmclist = jjgFbgcSdgcSdhpMapper.selectsdmc(proname,htd,fbgc);
        return sdmclist;
    }
}
