package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcSdgcSdhntlmhdzxf;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.sdgc.JjgFbgcSdgcSdhntlmhdzxfVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcSdgcSdhntlmhdzxfMapper;
import glgc.jjgys.system.service.JjgFbgcSdgcSdhntlmhdzxfService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.service.ProjectService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.commons.lang3.StringUtils;
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
 * @since 2023-05-20
 */
@Service
public class JjgFbgcSdgcSdhntlmhdzxfServiceImpl extends ServiceImpl<JjgFbgcSdgcSdhntlmhdzxfMapper, JjgFbgcSdgcSdhntlmhdzxf> implements JjgFbgcSdgcSdhntlmhdzxfService {

    @Autowired
    private JjgFbgcSdgcSdhntlmhdzxfMapper jjgFbgcSdgcSdhntlmhdzxfMapper;
    @Autowired
    private ProjectService projectService;



    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        List<Map<String,Object>> sdmclist = jjgFbgcSdgcSdhntlmhdzxfMapper.selectsdmc(proname,htd);
        if (sdmclist.size()>0){
            for (Map<String, Object> m : sdmclist)
            {
                for (String k : m.keySet()){
                    String sdmc = m.get(k).toString();
                    DBtoExcelsd(proname,htd,fbgc,sdmc);
                }
            }
        }

    }

    /**
     *
     * @param proname
     * @param htd
     * @param fbgc
     * @param sdmc
     */
    private void DBtoExcelsd(String proname, String htd, String fbgc, String sdmc) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        QueryWrapper<JjgFbgcSdgcSdhntlmhdzxf> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.like("sdmc",sdmc);
        wrapper.orderByAsc("zh");
        List<JjgFbgcSdgcSdhntlmhdzxf> data = jjgFbgcSdgcSdhntlmhdzxfMapper.selectList(wrapper);
        Integer level = projectService.getlevel(proname);
        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"54隧道混凝土路面厚度-钻芯法-"+sdmc+".xlsx");
        if (data == null || data.size()==0 ){
            return;
        }else {
            File fdir = new File(filepath + File.separator + proname + File.separator + htd);
            if (!fdir.exists()) {
                //创建文件根目录
                fdir.mkdirs();
            }

            File directory = new File("service-system/src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String path = reportPath + File.separator + "混凝土路面厚度-钻芯法.xlsx";
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            String sheetname = "混凝土路面";
            createTable(gettableNum(data.size()),wb,sheetname);
            if(DBtoExcel(data,wb)){

                for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                    if (shouldBeCalculate(wb.getSheetAt(j))) {
                        calculateThicknessSheet(wb.getSheetAt(j),level);
                        getTunnelTotal(wb.getSheetAt(j));
                    }
                }
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
                //System.out.println("下一张表！！！");
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

                //原先的计算方式
//				sheet.getRow(i+1).getCell(4).setCellValue(((sheet.getRow(3).getCell(2).getNumericCellValue()) <=60
//						? -5 : (sheet.getRow(3).getCell(2).getNumericCellValue())*-0.08));
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
     * @param data
     * @param wb
     * @return
     * @throws ParseException
     */
    private boolean DBtoExcel(List<JjgFbgcSdgcSdhntlmhdzxf> data, XSSFWorkbook wb) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        if(data.size() > 0){
            XSSFSheet sheet = wb.getSheet("混凝土路面");
            int index = 6;
            sheet.getRow(1).getCell(2).setCellValue(data.get(0).getProname());
            sheet.getRow(1).getCell(7).setCellValue(data.get(0).getHtd());
            sheet.getRow(2).getCell(2).setCellValue(data.get(0).getFbgc());//
            sheet.getRow(2).getCell(7).setCellValue("路面面层（混凝土路面）");
            String date = simpleDateFormat.format(data.get(0).getJcsj());
            for(int i =1; i < data.size(); i++){
                date = JjgFbgcCommonUtils.getLastDate(date, simpleDateFormat.format(data.get(i).getJcsj()));
            }
            sheet.getRow(3).getCell(2).setCellValue(Double.parseDouble(data.get(0).getSjz()));
            sheet.getRow(3).getCell(7).setCellValue(date);
            for(int i =0; i < data.size(); i++){
                sheet.getRow(index+i).getCell(0).setCellValue(i+1);
                sheet.addMergedRegion(new CellRangeAddress(i+index, i+index, 1, 2));
                sheet.getRow(index+i).getCell(1).setCellValue(data.get(i).getLqszd()+data.get(i).getZh());
                sheet.getRow(index+i).getCell(3).setCellValue(data.get(i).getBw());
                sheet.getRow(index+i).getCell(4).setCellValue(Double.parseDouble(data.get(i).getCz1()));
                sheet.getRow(index+i).getCell(5).setCellValue(Double.parseDouble(data.get(i).getCz2()));
                sheet.getRow(index+i).getCell(6).setCellValue(Double.parseDouble(data.get(i).getCz3()));
                sheet.getRow(index+i).getCell(7).setCellValue(Double.parseDouble(data.get(i).getCz4()));
                sheet.getRow(index+i).getCell(9).setCellValue(Double.parseDouble(data.get(i).getYxps())<=60?-10:Double.parseDouble(data.get(i).getYxps())*-0.15);
            }
            return true;
        }
        else{

            return false;
        }
    }

    /**
     *
     * @param tableNum
     * @param wb
     */
    private void createTable(int tableNum, XSSFWorkbook wb,String sheetname) {
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

    private int gettableNum(int size) {
        return size%26 <= 24 ? size/26+1 : size/26+2;
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        //String fbgc = commonInfoVo.getFbgc();
        List<Map<String, Object>> mapList = new ArrayList<>();

        List<Map<String,Object>> sdmclist = jjgFbgcSdgcSdhntlmhdzxfMapper.selectsdmc(proname,htd);
        if (sdmclist.size()>0){
            for (Map<String, Object> m : sdmclist) {
                for (String k : m.keySet()){
                    String sdmc = m.get(k).toString();
                    List<Map<String, Object>>  looksdjdb = looksdjdb(proname, htd, sdmc);
                    mapList.addAll(looksdjdb);
                }
            }
            return mapList;
        }else {
            return null;
        }
    }

    /**
     *
     * @param proname
     * @param htd
     * @param sdmc
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>>  looksdjdb(String proname, String htd, String sdmc) throws IOException {
        //String sheetname = "混凝土路面";

        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        //获取鉴定表文件
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"54隧道混凝土路面厚度-钻芯法-"+sdmc+".xlsx");
        if(!f.exists()){
            return null;
        }else {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
            List<Map<String,Object>> mapList = new ArrayList<>();
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                if (!wb.isSheetHidden(wb.getSheetIndex(wb.getSheetAt(j)))) {
                    XSSFSheet slSheet = wb.getSheetAt(j);
                    XSSFCell xmname = slSheet.getRow(1).getCell(2);//项目名
                    XSSFCell htdname = slSheet.getRow(1).getCell(7);//合同段名
                    Map<String, Object> jgmap = new HashMap<>();

                    if (proname.equals(xmname.toString()) && htd.equals(htdname.toString())) {
                        //获取到最后一行
                        int lastRowNum = slSheet.getLastRowNum();
                        slSheet.getRow(lastRowNum).getCell(3).setCellType(CellType.STRING);//总点数
                        slSheet.getRow(lastRowNum).getCell(6).setCellType(CellType.STRING);//合格点数
                        slSheet.getRow(lastRowNum).getCell(10).setCellType(CellType.STRING);//合格率
                        slSheet.getRow(lastRowNum - 2).getCell(15).setCellType(CellType.STRING);
                        slSheet.getRow(lastRowNum - 2).getCell(16).setCellType(CellType.STRING);
                        slSheet.getRow(lastRowNum - 2).getCell(10).setCellType(CellType.STRING);
                        slSheet.getRow(3).getCell(2).setCellType(CellType.STRING);
                        double zds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(3).getStringCellValue());
                        double hgds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(6).getStringCellValue());
                        double hgl = Double.valueOf(slSheet.getRow(lastRowNum).getCell(10).getStringCellValue());
                        String zdsz1 = decf.format(zds);
                        String hgdsz1 = decf.format(hgds);
                        String hglz1 = df.format(hgl);
                        jgmap.put("路面类型", wb.getSheetName(j));
                        jgmap.put("检测项目", sdmc);
                        jgmap.put("检测点数", zdsz1);
                        jgmap.put("合格点数", hgdsz1);
                        jgmap.put("合格率", hglz1);
                        jgmap.put("最大值", slSheet.getRow(lastRowNum - 2).getCell(15).getStringCellValue());
                        jgmap.put("最小值", slSheet.getRow(lastRowNum - 2).getCell(16).getStringCellValue());
                        jgmap.put("代表值", slSheet.getRow(lastRowNum - 2).getCell(10).getStringCellValue());
                        jgmap.put("设计值", slSheet.getRow(3).getCell(2).getStringCellValue());
                        mapList.add(jgmap);

                    }

                }
            }
            return mapList;

        }
    }

    @Override
    public void exportsdhntzxf(HttpServletResponse response) {
        String fileName = "14隧道混凝土路面厚度钻芯法实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcSdgcSdhntlmhdzxfVo()).finish();

    }

    @Override
    public void importsdhntzxf(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcSdgcSdhntlmhdzxfVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcSdgcSdhntlmhdzxfVo>(JjgFbgcSdgcSdhntlmhdzxfVo.class) {
                                @Override
                                public void handle(List<JjgFbgcSdgcSdhntlmhdzxfVo> dataList) {
                                    for(JjgFbgcSdgcSdhntlmhdzxfVo zxfvo: dataList)
                                    {
                                        JjgFbgcSdgcSdhntlmhdzxf sdlmzxf = new JjgFbgcSdgcSdhntlmhdzxf();
                                        BeanUtils.copyProperties(zxfvo,sdlmzxf);
                                        sdlmzxf.setCreatetime(new Date());
                                        sdlmzxf.setProname(commonInfoVo.getProname());
                                        sdlmzxf.setHtd(commonInfoVo.getHtd());
                                        sdlmzxf.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcSdgcSdhntlmhdzxfMapper.insert(sdlmzxf);
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
        List<Map<String,Object>> sdmclist = jjgFbgcSdgcSdhntlmhdzxfMapper.selectsdmc(proname,htd);
        return sdmclist;
    }

    @Override
    public List<Map<String, Object>> lookjg(CommonInfoVo commonInfoVo, String value) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        //获取鉴定表文件
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+value);
        if(!f.exists()){
            return null;
        }else {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
            List<Map<String, Object> > jgmap = new ArrayList<>();
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                if (!wb.isSheetHidden(wb.getSheetIndex(wb.getSheetAt(j)))) {
                    XSSFSheet slSheet = wb.getSheetAt(j);
                    XSSFCell xmname = slSheet.getRow(1).getCell(2);//项目名
                    XSSFCell htdname = slSheet.getRow(1).getCell(7);//合同段名
                    Map map = new HashMap();

                    if (proname.equals(xmname.toString()) && htd.equals(htdname.toString())) {
                        //获取到最后一行
                        int lastRowNum = slSheet.getLastRowNum();
                        slSheet.getRow(lastRowNum).getCell(3).setCellType(CellType.STRING);//总点数
                        slSheet.getRow(lastRowNum).getCell(6).setCellType(CellType.STRING);//合格点数
                        slSheet.getRow(lastRowNum).getCell(10).setCellType(CellType.STRING);//合格率
                        slSheet.getRow(3).getCell(2).setCellType(CellType.STRING);//合格率
                        double zds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(3).getStringCellValue());
                        double hgds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(6).getStringCellValue());
                        double hgl = Double.valueOf(slSheet.getRow(lastRowNum).getCell(10).getStringCellValue());
                        String zdsz1 = decf.format(zds);
                        String hgdsz1 = decf.format(hgds);
                        String hglz1 = df.format(hgl);
                        map.put("检测项目", StringUtils.substringBetween(value, "-", "."));
                        map.put("检测点数", zdsz1);
                        map.put("路面类型", wb.getSheetName(j));
                        map.put("合格点数", hgdsz1);
                        map.put("合格率", hglz1);
                        map.put("设计值", slSheet.getRow(3).getCell(2).getStringCellValue());
                        jgmap.add(map);
                    }

                }
            }
            return jgmap;
        }

    }
}
