package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLmgcLmssxs;
import glgc.jjgys.model.project.JjgFbgcSdgcLmgzsdsgpsf;
import glgc.jjgys.model.project.JjgFbgcSdgcLmssxs;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.lmgc.JjgFbgcLmgcLmssxsVo;
import glgc.jjgys.model.projectvo.sdgc.JjgFbgcSdgcLmssxsVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcSdgcLmssxsMapper;
import glgc.jjgys.system.service.JjgFbgcSdgcLmssxsService;
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
public class JjgFbgcSdgcLmssxsServiceImpl extends ServiceImpl<JjgFbgcSdgcLmssxsMapper, JjgFbgcSdgcLmssxs> implements JjgFbgcSdgcLmssxsService {


    @Autowired
    private JjgFbgcSdgcLmssxsMapper jjgFbgcSdgcLmssxsMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        List<Map<String,Object>> sdmclist = jjgFbgcSdgcLmssxsMapper.selectsdmc(proname,htd,fbgc);
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
        QueryWrapper<JjgFbgcSdgcLmssxs> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.like("sdmc",sdmc);
        wrapper.orderByAsc("zh");
        List<JjgFbgcSdgcLmssxs> data = jjgFbgcSdgcLmssxsMapper.selectList(wrapper);

        List<JjgFbgcSdgcLmssxs> zxzfdata = new ArrayList<>();
        List<JjgFbgcSdgcLmssxs> zxyfdata = new ArrayList<>();
        List<JjgFbgcSdgcLmssxs> sdzfdata = new ArrayList<>();
        List<JjgFbgcSdgcLmssxs> sdyfdata = new ArrayList<>();
        List<JjgFbgcSdgcLmssxs> zddata = new ArrayList<>();
        for (JjgFbgcSdgcLmssxs lmssxs : data){
            if (lmssxs.getZh().substring(0,1).equals("Z") && lmssxs.getLzdsd().equals("路")){
                zxzfdata.add(lmssxs);
            }else if (lmssxs.getZh().substring(0,1).equals("Y") && lmssxs.getLzdsd().equals("路")){
                zxyfdata.add(lmssxs);
            }else if (lmssxs.getZh().substring(0,1).equals("Z") && lmssxs.getLzdsd().contains("隧道")){
                sdzfdata.add(lmssxs);
            }else if (lmssxs.getZh().substring(0,1).equals("Y") && lmssxs.getLzdsd().contains("隧道")){
                sdyfdata.add(lmssxs);
            }else if(lmssxs.getLzdsd().contains("匝道") || lmssxs.getLzdsd().contains("互通") || lmssxs.getLzdsd().contains("服务区")){
                zddata.add(lmssxs);
            }
        }

        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"46隧道沥青路面渗水系数-"+sdmc+".xlsx");
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
            String path = reportPath + File.separator + "沥青路面渗水系数.xlsx";
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            createTable(gettableNum(zxzfdata.size()),wb,"沥青路面(左幅)");
            createTable(gettableNum(zxyfdata.size()),wb,"沥青路面(右幅)");
            createTable(gettableNum(sdzfdata.size()),wb,"隧道路面(左幅)");
            createTable(gettableNum(sdyfdata.size()),wb,"隧道路面(右幅)");
            createTable(gettableNum(zddata.size()),wb,"匝道路面");

            DBtoExcel(zxzfdata,wb,"沥青路面(左幅)");
            DBtoExcel(zxyfdata,wb,"沥青路面(右幅)");
            DBtoExcel(sdzfdata,wb,"隧道路面(左幅)");
            DBtoExcel(sdyfdata,wb,"隧道路面(右幅)");
            DBtoExcel(zddata,wb,"匝道路面");
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                if (shouldBeCalculate(wb.getSheetAt(j))) {
                    calculateOneSheet(wb.getSheetAt(j));
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
            out.close();
            wb.close();
        }

    }

    /**
     *
     * @param sheet
     */
    private void calculateOneSheet(XSSFSheet sheet) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        XSSFRow rowtol = null;
        boolean flag = false, cal = false;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if (flag) {
                if (row.getCell(1).getCellType() != Cell.CELL_TYPE_NUMERIC
                        && !"合计".equals(row.getCell(0).toString())) {
                    continue;
                }
                if ("".equals(row.getCell(0).toString())
                        && row.getCell(1).getCellType() == Cell.CELL_TYPE_NUMERIC) {// !"合计".equals(row.getCell(0).toString())){
                    cal = true;
                    rowend = row;
                } else {
                    if (cal) {
                        rowstart.getCell(8).setCellFormula(
                                "AVERAGE(" + rowstart.getCell(7).getReference()
                                        + ":"
                                        + rowend.getCell(7).getReference()
                                        + ")");// =AVERAGE(H7:H11)
                        rowstart.getCell(10).setCellFormula(
                                "IF(" + rowstart.getCell(8).getReference()
                                        + "<="
                                        + rowstart.getCell(9).getReference()
                                        + ",\"√\",\"\")");// =IF(I7<=J7,"√","")
                        //=IF(COUNTIF(C17:F21,"/"),"",IF(I17<=J17,"√",""))

                        rowstart.getCell(11).setCellFormula(
                                "IF(" + rowstart.getCell(10).getReference()
                                        + "=\"\",\"×\",\"\")");// =IF(K7="","×","")
                        //=IF(COUNTIF(C17:F21,"/"),"×",IF(I17<=J17,"","×"))
                    }
                    rowstart = row;
                    cal = false;

                }
                row.getCell(7).setCellFormula(
                        "IF(ISERROR(ROUND((" + row.getCell(5).getReference() + "-"
                                + row.getCell(2).getReference() + ")/"
                                + row.getCell(6).getReference() + "*+60,1)),"
                                + "\"/\","
                                +"ROUND((" + row.getCell(5).getReference() + "-"
                                + row.getCell(2).getReference() + ")/"
                                + row.getCell(6).getReference() + "*+60,1))");
                //IF(ISERROR(ROUND((F88-C88)/G88*60,1)),"/",ROUND((F88-C88)/G88*60,1))
            }
            if ("桩号".equals(row.getCell(0).toString())) {
                i++;
                rowstart = sheet.getRow(i + 1);
                rowtol = rowstart;
                rowend = rowstart;
                flag = true;
            }
            if ("合计".equals(row.getCell(0).toString())) {
                row.getCell(4).setCellFormula(
                        "COUNT(" + rowtol.getCell(8).getReference() + ":"
                                + rowend.getCell(8).getReference() + ")");// =COUNT(I7:I61)
                row.getCell(8)
                        .setCellFormula(
                                "COUNTIF(" + rowtol.getCell(10).getReference()
                                        + ":"
                                        + rowend.getCell(10).getReference()
                                        + ",\"√\")");// =COUNTIF(K7:K61,"√")
                row.getCell(10).setCellFormula(
                        row.getCell(8).getReference() + "/"
                                + row.getCell(4).getReference() + "*100");// =I66/E66*100

                row.createCell(12).setCellFormula(
                        "MAX(" + rowtol.getCell(8).getReference() + ":"
                                + rowend.getCell(8).getReference() + ")");// =COUNT(I7:I61)
                row.createCell(13).setCellFormula(
                        "MIN(" + rowtol.getCell(8).getReference() + ":"
                                + rowend.getCell(8).getReference() + ")");// =COUNT(I7:I61)
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
     */
    private boolean DBtoExcel(List<JjgFbgcSdgcLmssxs> data, XSSFWorkbook wb,String sheetname) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        if(data.size() > 0){
            XSSFSheet sheet = wb.getSheet(sheetname);
            String type = data.get(0).getSdmc();
            String zh = data.get(0).getZh();
            int index = 6;
            sheet.getRow(1).getCell(2).setCellValue(data.get(0).getProname());
            sheet.getRow(1).getCell(8).setCellValue(data.get(0).getHtd());
            sheet.getRow(2).getCell(2).setCellValue(data.get(0).getFbgc());
            sheet.getRow(2).getCell(8).setCellValue("路面面层");
            sheet.getRow(3).getCell(2).setCellValue("AC沥青混凝土");
            String date = simpleDateFormat.format(data.get(0).getJcsj());
            for(int i =1; i < data.size(); i++){
                date = JjgFbgcCommonUtils.getLastDate(date, simpleDateFormat.format(data.get(i).getJcsj()));
            }
            sheet.getRow(3).getCell(8).setCellValue(date);
            for(int i =0; i < data.size(); i++){
                if(type.equals(data.get(i).getSdmc()) && zh.equals(data.get(i).getZh())){
                    for(int k = 0; k < 5; k++){
                        if(i+k < data.size()){
                            fillCommonCellData(sheet, index, data.get(i+k));
                            index ++;

                        }
                        else{
                            break;
                        }
                    }
                    sheet.addMergedRegion(new CellRangeAddress(index-5, index-1, 0, 0));
                    sheet.getRow(index-5).getCell(0).setCellValue(type+zh);
                    sheet.addMergedRegion(new CellRangeAddress(index-5, index-1, 8, 8));
                    sheet.addMergedRegion(new CellRangeAddress(index-5, index-1, 9, 9));
                    sheet.getRow(index-5).getCell(9).setCellValue(Integer.parseInt(data.get(i).getSsxsgdz()));
                    sheet.addMergedRegion(new CellRangeAddress(index-5, index-1, 10, 10));
                    sheet.addMergedRegion(new CellRangeAddress(index-5, index-1, 11, 11));
                    i += 4;
                }
                else{
                    type = data.get(i).getSdmc();
                    zh = data.get(i).getZh();
                    i--;
                }
            }
            return true;
        }
        else{
            return false;
        }
    }

    /**
     *
     * @param sheet
     * @param index
     * @param row
     */
    private void fillCommonCellData(XSSFSheet sheet, int index, JjgFbgcSdgcLmssxs row) {
        sheet.getRow(index).getCell(1).setCellValue((index-1)%5+1);
        sheet.getRow(index).getCell(2).setCellValue(Integer.parseInt(row.getCds()));
        sheet.getRow(index).getCell(3).setCellValue(Integer.parseInt(row.getOfzds()));
        sheet.getRow(index).getCell(4).setCellValue(Integer.parseInt(row.getTfzds()));
        sheet.getRow(index).getCell(5).setCellValue(Integer.parseInt(row.getSl()));
        sheet.getRow(index).getCell(6).setCellValue(Integer.parseInt(row.getSj()));
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
                RowCopy.copyRows(wb, sheetname, sheetname, 6, 35, (i - 1) * 30 + 36);
            }
            else{
                RowCopy.copyRows(wb, sheetname, sheetname, 6, 34, (i - 1) * 30 + 36);
            }
        }
        if(record == 1){
            wb.getSheet(sheetname).shiftRows(36, 36, -1);
        }
        RowCopy.copyRows(wb, "source", sheetname, 0, 0,(record) * 30 + 5);
        wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 11, 0,(record) * 30 + 5);
    }

    /**
     *
     * @param size
     * @return
     */
    private int gettableNum(int size) {
        return size%36 <= 35 ? size/36+1 : size/36+2;
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        //String fbgc = commonInfoVo.getFbgc();
        String title = "路面渗水系数质量鉴定表";
        //String sheetname = "隧道路面";

        List<Map<String, Object>> mapList = new ArrayList<>();

        List<Map<String,Object>> sdmclist = jjgFbgcSdgcLmssxsMapper.selectsdmc1(proname,htd);
        if (sdmclist.size()>0){
            for (Map<String, Object> m : sdmclist) {
                for (String k : m.keySet()){
                    String sdmc = m.get(k).toString();
                    List<Map<String, Object>> looksdjdb = looksdjdb(proname, htd, sdmc,title);
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
     * @param title
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> looksdjdb(String proname, String htd, String sdmc, String title) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "46隧道沥青路面渗水系数-"+sdmc+".xlsx");
        if (!f.exists()) {
            return null;
        } else {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
            List<Map<String, Object> > jgmap = new ArrayList<>();
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                if (!wb.isSheetHidden(wb.getSheetIndex(wb.getSheetAt(j)))) {
                    XSSFSheet slSheet = wb.getSheetAt(j);
                    XSSFCell bt = slSheet.getRow(0).getCell(0);//标题
                    XSSFCell xmname = slSheet.getRow(1).getCell(2);//项目名
                    XSSFCell htdname = slSheet.getRow(1).getCell(8);//合同段名
                    Map map = new HashMap();
                    if (proname.equals(xmname.toString()) && title.equals(bt.toString()) && htd.equals(htdname.toString())) {
                        //获取到最后一行
                        int lastRowNum = slSheet.getLastRowNum();
                        slSheet.getRow(lastRowNum).getCell(4).setCellType(CellType.STRING);//总点数
                        slSheet.getRow(lastRowNum).getCell(8).setCellType(CellType.STRING);//合格点数
                        slSheet.getRow(lastRowNum).getCell(10).setCellType(CellType.STRING);//合格率
                        double zds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(4).getStringCellValue());
                        double hgds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(8).getStringCellValue());
                        double hgl = Double.valueOf(slSheet.getRow(lastRowNum).getCell(10).getStringCellValue());
                        String zdsz1 = decf.format(zds);
                        String hgdsz1 = decf.format(hgds);
                        String hglz1 = df.format(hgl);
                        map.put("检测项目", sdmc);
                        map.put("路面类型", wb.getSheetName(j));
                        map.put("检测点数", zdsz1);
                        map.put("合格点数", hgdsz1);
                        map.put("合格率", hglz1);
                    }
                    jgmap.add(map);
                }
            }
            return jgmap;
        }


    }

    @Override
    public void exportSdlmssxs(HttpServletResponse response) {
        String fileName = "06隧道沥青路面压实度实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcSdgcLmssxsVo()).finish();

    }

    @Override
    public void importSdlmssxs(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcSdgcLmssxsVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcSdgcLmssxsVo>(JjgFbgcSdgcLmssxsVo.class) {
                                @Override
                                public void handle(List<JjgFbgcSdgcLmssxsVo> dataList) {
                                    for(JjgFbgcSdgcLmssxsVo lmssxsVo: dataList)
                                    {
                                        JjgFbgcSdgcLmssxs lmssxs = new JjgFbgcSdgcLmssxs();
                                        BeanUtils.copyProperties(lmssxsVo,lmssxs);
                                        lmssxs.setCreatetime(new Date());
                                        lmssxs.setProname(commonInfoVo.getProname());
                                        lmssxs.setHtd(commonInfoVo.getHtd());
                                        lmssxs.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcSdgcLmssxsMapper.insert(lmssxs);
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
        List<Map<String,Object>> sdmclist = jjgFbgcSdgcLmssxsMapper.selectsdmc(proname,htd,fbgc);
        return sdmclist;
    }

    @Override
    public List<Map<String, Object>> lookjg(CommonInfoVo commonInfoVo,String value) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + value);
        if (!f.exists()) {
            return null;
        } else {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
            List<Map<String, Object> > jgmap = new ArrayList<>();
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                if (!wb.isSheetHidden(wb.getSheetIndex(wb.getSheetAt(j)))) {
                    XSSFSheet slSheet = wb.getSheetAt(j);
                    XSSFCell bt = slSheet.getRow(0).getCell(0);//标题
                    XSSFCell xmname = slSheet.getRow(1).getCell(2);//项目名
                    XSSFCell htdname = slSheet.getRow(1).getCell(8);//合同段名
                    Map map = new HashMap();
                    if (proname.equals(xmname.toString()) && htd.equals(htdname.toString())) {
                        //获取到最后一行
                        int lastRowNum = slSheet.getLastRowNum();
                        slSheet.getRow(lastRowNum).getCell(4).setCellType(CellType.STRING);//总点数
                        slSheet.getRow(lastRowNum).getCell(8).setCellType(CellType.STRING);//合格点数
                        slSheet.getRow(lastRowNum).getCell(10).setCellType(CellType.STRING);//合格率
                        slSheet.getRow(6).getCell(9).setCellType(CellType.STRING);//合格率
                        double zds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(4).getStringCellValue());
                        double hgds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(8).getStringCellValue());
                        double hgl = Double.valueOf(slSheet.getRow(lastRowNum).getCell(10).getStringCellValue());
                        String zdsz1 = decf.format(zds);
                        String hgdsz1 = decf.format(hgds);
                        String hglz1 = df.format(hgl);
                        map.put("检测项目", StringUtils.substringBetween(value, "-", "."));
                        map.put("路面类型", wb.getSheetName(j));
                        map.put("检测点数", zdsz1);
                        map.put("合格点数", hgdsz1);
                        map.put("合格率", hglz1);
                        map.put("规定值", slSheet.getRow(6).getCell(9).getStringCellValue());
                    }
                    jgmap.add(map);
                }
            }
            return jgmap;
        }    }
}
