package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLmgcTlmxlbgc;
import glgc.jjgys.model.project.JjgFbgcSdgcLmssxs;
import glgc.jjgys.model.project.JjgFbgcSdgcTlmxlbgc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.lmgc.JjgFbgcLmgcTlmxlbgcVo;
import glgc.jjgys.model.projectvo.sdgc.JjgFbgcSdgcTlmxlbgcVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcSdgcTlmxlbgcMapper;
import glgc.jjgys.system.service.JjgFbgcSdgcTlmxlbgcService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
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
 * @since 2023-05-04
 */
@Service
public class JjgFbgcSdgcTlmxlbgcServiceImpl extends ServiceImpl<JjgFbgcSdgcTlmxlbgcMapper, JjgFbgcSdgcTlmxlbgc> implements JjgFbgcSdgcTlmxlbgcService {

    @Autowired
    private JjgFbgcSdgcTlmxlbgcMapper jjgFbgcSdgcTlmxlbgcMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        List<Map<String,Object>> sdmclist = jjgFbgcSdgcTlmxlbgcMapper.selectsdmc(proname,htd);
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
     * @throws IOException
     * @throws ParseException
     */
    private void DBtoExcelsd(String proname, String htd, String fbgc, String sdmc) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        QueryWrapper<JjgFbgcSdgcTlmxlbgc> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.like("sdmc",sdmc);
        wrapper.orderByAsc("zh");
        List<JjgFbgcSdgcTlmxlbgc> data = jjgFbgcSdgcTlmxlbgcMapper.selectList(wrapper);

        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"48隧道混凝土路面相邻板高差-"+sdmc+".xlsx");
        if (data == null || data.size()==0 ){
            return;
        }else {
            File fdir = new File(filepath + File.separator + proname + File.separator + htd);
            if (!fdir.exists()) {
                //创建文件根目录
                fdir.mkdirs();
            }
            File directory = new File("src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String path = reportPath + File.separator + "混凝土路面相邻板高差.xlsx";
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            createTable(gettableNum(data.size()),wb);
            if(DBtoExcel(data,wb)){
                for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                    if (shouldBeCalculate(wb.getSheetAt(j))) {
                        calculateSheet(wb.getSheetAt(j));
                        getTunnelTotal(wb.getSheetAt(j));
                    }
                }
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
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if ("合计".equals(row.getCell(0).toString())) {
                endrow = sheet.getRow(i-1);
                startrow.createCell(8).setCellFormula("COUNT("
                        +startrow.getCell(3).getReference()+":"
                        +endrow.getCell(3).getReference()+")");//=COUNT(D7:D36)
                startrow.createCell(9).setCellFormula("COUNTIF("
                        +startrow.getCell(4).getReference()+":"
                        +endrow.getCell(4).getReference()+",\"√\")");//=COUNTIF(E7:E36,"√")
                startrow.createCell(10).setCellFormula(startrow.getCell(9).getReference()+"*100/"
                        +startrow.getCell(8).getReference());
                break;
            }
            if(flag && row.getCell(1) != null && !"".equals(row.getCell(1).toString())){
                if(!name.equals(row.getCell(1).toString().substring(0, row.getCell(1).toString().indexOf("K"))) && !"".equals(name)){
                    endrow = sheet.getRow(i-1);
                    startrow.createCell(8).setCellFormula("COUNT("
                            +startrow.getCell(3).getReference()+":"
                            +endrow.getCell(3).getReference()+")");//=COUNT(D7:D36)
                    startrow.createCell(9).setCellFormula("COUNTIF("
                            +startrow.getCell(4).getReference()+":"
                            +endrow.getCell(4).getReference()+",\"√\")");//=COUNTIF(E7:E36,"√")
                    startrow.createCell(10).setCellFormula(startrow.getCell(9).getReference()+"*100/"
                            +startrow.getCell(8).getReference());
                    name = row.getCell(1).toString().substring(0, row.getCell(1).toString().indexOf("K"));
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

        }

    }

    /**
     *
     * @param sheet
     */
    private void calculateSheet(XSSFSheet sheet) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        boolean flag = false;
        //System.out.println(sheet.getPhysicalNumberOfRows());
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            // 可以计算
            if (row.getCell(3).getCellType() == Cell.CELL_TYPE_NUMERIC
                    && flag) {
                row.getCell(4).setCellFormula(
                        "IF(" + row.getCell(3).getReference()
                                + "=\"\",\"\",IF("
                                + row.getCell(3).getReference()
                                + "<=$C$4,\"√\",\"\"))");// E=IF(D7="","",IF(D7<=$C$4,"√",""))
                row.getCell(6).setCellFormula(
                        "IF(" + row.getCell(3).getReference()
                                + ">$C$4,\"×\",\"\")");// G=IF(D7>$C$4,"×","")
                rowend = row;
            }
            // 可以计算啦
            if ("序号".equals(row.getCell(0).toString())) {
                flag = true;
                i++;
                rowstart = sheet.getRow(i + 1);
                rowend = rowstart;
            }
            if ("合计".equals(row.getCell(0).toString())) {

                row.getCell(3).setCellFormula(
                        "COUNT(" + rowstart.getCell(3).getReference() + ":"
                                + rowend.getCell(3).getReference() + ")");// =COUNT(D7:D30)
                row.getCell(5).setCellFormula(
                        "COUNTIF(" + rowstart.getCell(4).getReference() + ":"
                                + rowend.getCell(4).getReference() + ",\"√\")");// =COUNTIF(E7:F30,"√")
                row.getCell(7).setCellFormula(
                        row.getCell(5).getReference() + "/"
                                + row.getCell(3).getReference() + "*100");// =F36/D36*100
                row.createCell(8).setCellFormula(
                        "MAX(" + rowstart.getCell(3).getReference() + ":"
                                + rowend.getCell(3).getReference() + ")");// =COUNT(D7:D30)
                row.createCell(9).setCellFormula(
                        "MIN(" + rowstart.getCell(3).getReference() + ":"
                                + rowend.getCell(3).getReference() + ")");// =COUNT(D7:D30)

                row.createCell(10).setCellFormula(
                        "AVERAGE(" + rowstart.getCell(3).getReference() + ":"
                                + rowend.getCell(3).getReference() + ")");// =COUNT(D7:D30)
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
        if (title.endsWith("鉴定表")) {
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
    private boolean DBtoExcel(List<JjgFbgcSdgcTlmxlbgc> data, XSSFWorkbook wb) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        if(data.size() > 0){
            XSSFSheet sheet = wb.getSheet("相邻板高差");
            int index = 6;
            sheet.getRow(1).getCell(2).setCellValue(data.get(0).getProname());
            sheet.getRow(1).getCell(6).setCellValue(data.get(0).getHtd());
            sheet.getRow(2).getCell(2).setCellValue(data.get(0).getFbgc());
            sheet.getRow(2).getCell(6).setCellValue("路面面层");
            //sheet.getRow(3).getCell(2).setCellValue(Double.valueOf(data.get(0).getBgcgdz()));
            String date = simpleDateFormat.format(data.get(0).getJcsj());
            for(int i =1; i < data.size(); i++){
                date = JjgFbgcCommonUtils.getLastDate(date, simpleDateFormat.format(data.get(i).getJcsj()));
            }
            sheet.getRow(3).getCell(6).setCellValue(date);
            for(int i =0; i < data.size(); i++){
                sheet.addMergedRegion(new CellRangeAddress(index+i, index+i+2, 0, 0));
                sheet.getRow(index+i).getCell(0).setCellValue(i/3+1);

                sheet.addMergedRegion(new CellRangeAddress(index+i, index+i+2, 1, 2));
                sheet.getRow(index+i).getCell(1).setCellValue(data.get(i).getSdmc()+ data.get(i).getZh());

                sheet.getRow(index+i).getCell(3).setCellValue(Double.parseDouble(data.get(i).getScz1()));
                sheet.addMergedRegion(new CellRangeAddress(index+i, index+i, 4, 5));
                sheet.addMergedRegion(new CellRangeAddress(index+i, index+i, 6, 7));
                sheet.getRow(index+i+1).getCell(3).setCellValue(Double.parseDouble(data.get(i).getScz2()));
                sheet.addMergedRegion(new CellRangeAddress(index+i+1, index+i+1, 4, 5));
                sheet.addMergedRegion(new CellRangeAddress(index+i+1, index+i+1, 6, 7));
                sheet.getRow(index+i+2).getCell(3).setCellValue(Double.parseDouble(data.get(i).getScz3()));
                sheet.addMergedRegion(new CellRangeAddress(index+i+2, index+i+2, 4, 5));
                sheet.addMergedRegion(new CellRangeAddress(index+i+2, index+i+2, 6, 7));
                i+=2;
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
    private void createTable(int tableNum, XSSFWorkbook wb) {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            if(i < record-1){
                RowCopy.copyRows(wb, "相邻板高差", "相邻板高差", 6, 35, (i - 1) * 30 + 36);
            }
            else{
                RowCopy.copyRows(wb, "相邻板高差", "相邻板高差", 6, 34, (i - 1) * 30 + 36);
            }
        }
        if(record == 1){
            wb.getSheet("相邻板高差").shiftRows(36, 36, -1);
        }
        RowCopy.copyRows(wb, "source", "相邻板高差", 0, 0,(record) * 30 + 5);
        wb.setPrintArea(wb.getSheetIndex("相邻板高差"), 0, 7, 0,(record) * 30 + 5);
    }

    /**
     *
     * @param size
     * @return
     */
    private int gettableNum(int size) {
        return size%30 <= 29 ? size/30+1 : size/30+2;
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        //String fbgc = commonInfoVo.getFbgc();
        String title = "混凝土路面相邻板高差质量鉴定表";
        String sheetname = "相邻板高差";

        List<Map<String, Object>> mapList = new ArrayList<>();

        List<Map<String,Object>> sdmclist = jjgFbgcSdgcTlmxlbgcMapper.selectsdmc(proname,htd);
        if (sdmclist.size()>0){
            for (Map<String, Object> m : sdmclist) {
                for (String k : m.keySet()){
                    String sdmc = m.get(k).toString();
                    Map<String, Object> looksdjdb = looksdjdb(proname, htd, sdmc, sheetname,title);
                    mapList.add(looksdjdb);
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
     * @param sheetname
     * @param title
     * @return
     */
    private Map<String, Object> looksdjdb(String proname, String htd, String sdmc, String sheetname, String title) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "48隧道混凝土路面相邻板高差-"+sdmc+".xlsx");
        if (!f.exists()) {
            return null;
        } else {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
            XSSFSheet slSheet = wb.getSheet(sheetname);
            XSSFCell bt = slSheet.getRow(0).getCell(0);//标题
            XSSFCell xmname = slSheet.getRow(1).getCell(2);//项目名
            XSSFCell htdname = slSheet.getRow(1).getCell(6);//合同段名
            Map<String, Object> jgmap = new HashMap<>();
            if (proname.equals(xmname.toString()) && title.equals(bt.toString()) && htd.equals(htdname.toString())) {
                //获取到最后一行
                int lastRowNum = slSheet.getLastRowNum();
                slSheet.getRow(lastRowNum).getCell(3).setCellType(CellType.STRING);//总点数
                slSheet.getRow(lastRowNum).getCell(5).setCellType(CellType.STRING);//合格点数
                slSheet.getRow(lastRowNum).getCell(7).setCellType(CellType.STRING);//合格率
                slSheet.getRow(3).getCell(2).setCellType(CellType.STRING);//合格率
                slSheet.getRow(lastRowNum).getCell(8).setCellType(CellType.STRING);
                slSheet.getRow(lastRowNum).getCell(9).setCellType(CellType.STRING);
                double zds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(3).getStringCellValue());
                double hgds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(5).getStringCellValue());
                double hgl = Double.valueOf(slSheet.getRow(lastRowNum).getCell(7).getStringCellValue());
                String zdsz1 = decf.format(zds);
                String hgdsz1 = decf.format(hgds);
                String hglz1 = df.format(hgl);
                jgmap.put("检测项目", sdmc);
                jgmap.put("检测点数", zdsz1);
                jgmap.put("合格点数", hgdsz1);
                jgmap.put("合格率", hglz1);
                jgmap.put("最大值", slSheet.getRow(lastRowNum).getCell(8).getStringCellValue());
                jgmap.put("最小值", slSheet.getRow(lastRowNum).getCell(9).getStringCellValue());
                jgmap.put("规定值", slSheet.getRow(3).getCell(2).getStringCellValue());
            }
            return jgmap;
        }

    }

    @Override
    public void exportsdxlbgs(HttpServletResponse response) {
        String fileName = "11隧道砼路面相邻板高差实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcSdgcTlmxlbgcVo()).finish();

    }

    @Override
    public void importsdxlbgs(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcSdgcTlmxlbgcVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcSdgcTlmxlbgcVo>(JjgFbgcSdgcTlmxlbgcVo.class) {
                                @Override
                                public void handle(List<JjgFbgcSdgcTlmxlbgcVo> dataList) {
                                    for(JjgFbgcSdgcTlmxlbgcVo tlmxlbgcVo: dataList)
                                    {
                                        JjgFbgcSdgcTlmxlbgc tlmxlbgc = new JjgFbgcSdgcTlmxlbgc();
                                        BeanUtils.copyProperties(tlmxlbgcVo,tlmxlbgc);
                                        tlmxlbgc.setCreatetime(new Date());
                                        tlmxlbgc.setProname(commonInfoVo.getProname());
                                        tlmxlbgc.setHtd(commonInfoVo.getHtd());
                                        tlmxlbgc.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcSdgcTlmxlbgcMapper.insert(tlmxlbgc);
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
        List<Map<String,Object>> sdmclist = jjgFbgcSdgcTlmxlbgcMapper.selectsdmc(proname,htd);
        return sdmclist;
    }

    @Override
    public List<Map<String, Object>> lookjg(CommonInfoVo commonInfoVo, String value) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + value);
        if (!f.exists()) {
            return null;
        } else {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
            List<Map<String, Object> > resultList = new ArrayList<>();
            XSSFSheet slSheet = wb.getSheet("相邻板高差");
            XSSFCell bt = slSheet.getRow(0).getCell(0);//标题
            XSSFCell xmname = slSheet.getRow(1).getCell(2);//项目名
            XSSFCell htdname = slSheet.getRow(1).getCell(6);//合同段名
            Map<String, Object> jgmap = new HashMap<>();
            if (proname.equals(xmname.toString()) &&  htd.equals(htdname.toString())) {
                //获取到最后一行
                int lastRowNum = slSheet.getLastRowNum();
                slSheet.getRow(lastRowNum).getCell(3).setCellType(CellType.STRING);//总点数
                slSheet.getRow(lastRowNum).getCell(5).setCellType(CellType.STRING);//合格点数
                slSheet.getRow(lastRowNum).getCell(7).setCellType(CellType.STRING);//合格率
                slSheet.getRow(3).getCell(2).setCellType(CellType.STRING);//合格率
                double zds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(3).getStringCellValue());
                double hgds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(5).getStringCellValue());
                double hgl = Double.valueOf(slSheet.getRow(lastRowNum).getCell(7).getStringCellValue());
                String zdsz1 = decf.format(zds);
                String hgdsz1 = decf.format(hgds);
                String hglz1 = df.format(hgl);
                jgmap.put("检测项目", StringUtils.substringBetween(value, "-", "."));
                jgmap.put("检测点数", zdsz1);
                jgmap.put("规定值", slSheet.getRow(3).getCell(2).getStringCellValue());
                jgmap.put("合格点数", hgdsz1);
                jgmap.put("合格率", hglz1);
                resultList.add(jgmap);
            }
            return resultList;
        }
    }
}
