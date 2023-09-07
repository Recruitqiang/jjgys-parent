package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLmgcLmssxs;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.lmgc.JjgFbgcLmgcLmssxsVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcLmgcLmssxsMapper;
import glgc.jjgys.system.service.JjgFbgcLmgcLmssxsService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
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
 * @since 2023-04-19
 */
@Service
public class JjgFbgcLmgcLmssxsServiceImpl extends ServiceImpl<JjgFbgcLmgcLmssxsMapper, JjgFbgcLmgcLmssxs> implements JjgFbgcLmgcLmssxsService {

    @Autowired
    private JjgFbgcLmgcLmssxsMapper jjgFbgcLmgcLmssxsMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获取数据
        QueryWrapper<JjgFbgcLmgcLmssxs> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("zh");
        List<JjgFbgcLmgcLmssxs> data = jjgFbgcLmgcLmssxsMapper.selectList(wrapper);

        List<JjgFbgcLmgcLmssxs> zxzfdata = new ArrayList<>();
        List<JjgFbgcLmgcLmssxs> zxyfdata = new ArrayList<>();
        List<JjgFbgcLmgcLmssxs> sdzfdata = new ArrayList<>();
        List<JjgFbgcLmgcLmssxs> sdyfdata = new ArrayList<>();
        List<JjgFbgcLmgcLmssxs> zddata = new ArrayList<>();
        for (JjgFbgcLmgcLmssxs lmssxs : data){
            if (lmssxs.getZh().substring(0,1).equals("Z") && lmssxs.getLxlx().equals("主线")){
                zxzfdata.add(lmssxs);
            }else if (lmssxs.getZh().substring(0,1).equals("Y") && lmssxs.getLxlx().equals("主线")){
                zxyfdata.add(lmssxs);
            }else if (lmssxs.getZh().substring(0,1).equals("Z") && lmssxs.getLxlx().contains("隧道")){
                sdzfdata.add(lmssxs);
            }else if (lmssxs.getZh().substring(0,1).equals("Y") && lmssxs.getLxlx().contains("隧道")){
                sdyfdata.add(lmssxs);
            }else if(lmssxs.getLxlx().contains("匝道") || lmssxs.getLxlx().contains("互通") || lmssxs.getLxlx().contains("服务区")){
                zddata.add(lmssxs);
            }
        }
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"15沥青路面渗水系数.xlsx");
        if (data == null || data.size()==0){
            return;
        }else {
            File fdir = new File(filepath+File.separator+proname+File.separator+htd);
            if(!fdir.exists()){
                //创建文件根目录
                fdir.mkdirs();
            }
            File directory = new File("service-system/src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String path =reportPath +File.separator+ "沥青路面渗水系数.xlsx";
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            createTable(getzxzftableNum(zxzfdata.size()),wb,"沥青路面(左幅)");
            createTable(getzxzftableNum(zxyfdata.size()),wb,"沥青路面(右幅)");
            createTable(getzxzftableNum(sdzfdata.size()),wb,"隧道路面(左幅)");
            createTable(getzxzftableNum(sdyfdata.size()),wb,"隧道路面(右幅)");
            createTable(getzxzftableNum(zddata.size()),wb,"匝道路面");
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
            if(wb.getSheet("隧道路面(左幅)") != null){
                getTunnelTotal(wb.getSheet("隧道路面(左幅)"));
            }
            if(wb.getSheet("隧道路面(右幅)") != null){
                getTunnelTotal(wb.getSheet("隧道路面(右幅)"));
            }
            if(wb.getSheet("匝道路面") != null){
                getTunnelTotal(wb.getSheet("匝道路面"));
            }

            for (int j = 0; j < wb.getNumberOfSheets(); j++) {   //表内公式  计算 显示结果
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
            if(flag && row.getCell(0) != null && !"".equals(row.getCell(0).toString())){
                if(!name.equals(
                        row.getCell(0).toString().contains("Z")?
                                row.getCell(0).toString().replaceAll("[^\u4E00-\u9FA5]", "")+"Z" :
                                row.getCell(0).toString().replaceAll("[^\u4E00-\u9FA5]", "")+"Y"
                ) && !"".equals(name)){
                    endrow = sheet.getRow(i-1);
                    startrow.createCell(12).setCellFormula("COUNT("
                            +startrow.getCell(8).getReference()+":"
                            +endrow.getCell(8).getReference()+")");//=COUNT(I7:I36)
                    startrow.createCell(13).setCellFormula("COUNTIF("
                            +startrow.getCell(10).getReference()+":"
                            +endrow.getCell(10).getReference()+",\"√\")");//=COUNTIF(K7:K36,"√")
                    startrow.createCell(14).setCellFormula(startrow.getCell(13).getReference()+"*100/"
                            +startrow.getCell(12).getReference());
                    name = row.getCell(0).toString().contains("Z")?
                            row.getCell(0).toString().replaceAll("[^\u4E00-\u9FA5]", "")+"Z" :
                            row.getCell(0).toString().replaceAll("[^\u4E00-\u9FA5]", "")+"Y";
                    startrow = row;
                }
                if("".equals(name)){
                    /*
                     * 隧道要分左右幅统计，但渗水系数没有分开统计，所以要根据桩号的z/y来判断
                     */
                    name = row.getCell(0).toString().contains("Z")?
                            row.getCell(0).toString().replaceAll("[^\u4E00-\u9FA5]", "")+"Z" :
                            row.getCell(0).toString().replaceAll("[^\u4E00-\u9FA5]", "")+"Y";
                    startrow = row;
                }

            }
            if ("桩号".equals(row.getCell(0).toString())) {
                flag = true;
                i++;
            }
            if ("评定".equals(row.getCell(0).toString())) {
                break;
            }
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
        for (int i = sheet.getFirstRowNum(); i <= sheet
                .getPhysicalNumberOfRows(); i++) {
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
                        rowstart.getCell(11).setCellFormula(
                                "IF(" + rowstart.getCell(10).getReference()
                                        + "=\"\",\"×\",\"\")");// =IF(K7="","×","")
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
                /*
                 *
                 */
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
     * 判断此sheet是否需要进行计算
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
     * @throws ParseException
     */
    private void DBtoExcel(List<JjgFbgcLmgcLmssxs> data, XSSFWorkbook wb,String sheetname) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        if (data.size()>0){
            XSSFSheet sheet = wb.getSheet(sheetname);
            String type = data.get(0).getLxlx();
            String zh = data.get(0).getZh();
            //String startpileNO = data.get(0)[6];
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
                if(type.equals(data.get(i).getLxlx()) && zh.equals(data.get(i).getZh())){
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
                    sheet.getRow(index-5).getCell(9).setCellValue(Double.valueOf(data.get(i).getSsxsgdz()));
                    sheet.addMergedRegion(new CellRangeAddress(index-5, index-1, 10, 10));
                    sheet.addMergedRegion(new CellRangeAddress(index-5, index-1, 11, 11));
                    i += 4;
                }
                else{
                    type = data.get(i).getLxlx();
                    zh = data.get(i).getZh();
                    i--;
                }
            }
        }

    }

    /**
     *
     * @param sheet
     * @param index
     * @param row
     */
    private void fillCommonCellData(XSSFSheet sheet, int index, JjgFbgcLmgcLmssxs row) {
        sheet.getRow(index).getCell(1).setCellValue((index-1)%5+1);
        if("/".equals(row.getCds())){
            sheet.getRow(index).getCell(2).setCellValue("/");
        }
        else{
            sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(row.getCds()).intValue());
        }
        if("/".equals(row.getOfzds())){
            sheet.getRow(index).getCell(3).setCellValue("/");
        }
        else{
            sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(row.getOfzds()).intValue());
        }
        if("/".equals(row.getTfzds())){
            sheet.getRow(index).getCell(4).setCellValue("/");
        }
        else{
            sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(row.getTfzds()).intValue());
        }
        if("/".equals(row.getSl())){
            sheet.getRow(index).getCell(5).setCellValue("/");
        }
        else{
            sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(row.getSl()).intValue());
        }
        if("/".equals(row.getSj())){
            sheet.getRow(index).getCell(6).setCellValue("/");
        }
        else{
            sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(row.getSj()).intValue());
        }
    }

    /**
     *
     * @param tableNum
     * @param wb
     * @param sheetname
     */
    private void createTable(int tableNum, XSSFWorkbook wb, String sheetname) {
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
    private int getzxzftableNum(int size) {
        return size%36 <= 35 ? size/36+1 : size/36+2;
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();

        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "15沥青路面渗水系数.xlsx");
        if (!f.exists()) {
            return null;
        } else {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
            boolean lmzfsheet = wb.isSheetHidden(0);
            boolean lmyfsheet = wb.isSheetHidden(1);
            boolean sdzfsheet = wb.isSheetHidden(2);
            boolean sdyfsheet = wb.isSheetHidden(3);
            boolean zdsheet = wb.isSheetHidden(4);
            List<Map<String,Object>> mapList = new ArrayList<>();
            Map<String,Object> map1 = new HashMap<>();
            Map<String,Object> map2 = new HashMap<>();
            Map<String,Object> map3 = new HashMap<>();
            Map<String,Object> map4 = new HashMap<>();
            Map<String,Object> map5 = new HashMap<>();

            XSSFSheet s1 = wb.getSheet("沥青路面(左幅)");
            XSSFSheet s2 = wb.getSheet("沥青路面(右幅)");
            XSSFSheet s3 = wb.getSheet("隧道路面(左幅)");
            XSSFSheet s4 = wb.getSheet("隧道路面(右幅)");
            XSSFSheet s5 = wb.getSheet("匝道路面");

            if (!lmzfsheet && proname.equals(s1.getRow(1).getCell(2).toString()) && htd.equals(s1.getRow(1).getCell(8).toString())){
                int sllastRowNum = s1.getLastRowNum();
                s1.getRow(sllastRowNum).getCell(4).setCellType(CellType.STRING);//检测点数
                s1.getRow(sllastRowNum).getCell(8).setCellType(CellType.STRING);//合格点数
                s1.getRow(sllastRowNum).getCell(10).setCellType(CellType.STRING);//合格率（%）
                s1.getRow(6).getCell(9).setCellType(CellType.STRING);//合格率（%）

                s1.getRow(sllastRowNum).getCell(12).setCellType(CellType.STRING);
                s1.getRow(sllastRowNum).getCell(13).setCellType(CellType.STRING);

                map1.put("检测项目","沥青路面(左幅)");
                map1.put("规定值",s1.getRow(6).getCell(9).getStringCellValue());
                map1.put("检测点数",decf.format(Double.valueOf(s1.getRow(sllastRowNum).getCell(4).getStringCellValue())));
                map1.put("合格点数",decf.format(Double.valueOf(s1.getRow(sllastRowNum).getCell(8).getStringCellValue())));
                map1.put("合格率",df.format(Double.valueOf(s1.getRow(sllastRowNum).getCell(10).getStringCellValue())));

                map1.put("最大值",s1.getRow(sllastRowNum).getCell(12).getStringCellValue());
                map1.put("最小值",s1.getRow(sllastRowNum).getCell(13).getStringCellValue());

                mapList.add(map1);
            }/*else {
                map1.put("检测项目","沥青路面(左幅)");
                map1.put("规定值",0);
                map1.put("检测点数",0);
                map1.put("合格点数",0);
                map1.put("合格率",0);
                mapList.add(map1);
            }*/

            if (!lmyfsheet && proname.equals(s2.getRow(1).getCell(2).toString()) && htd.equals(s2.getRow(1).getCell(8).toString())){
                int sllastRowNum = s2.getLastRowNum();
                s2.getRow(sllastRowNum).getCell(4).setCellType(CellType.STRING);//检测点数
                s2.getRow(sllastRowNum).getCell(8).setCellType(CellType.STRING);//合格点数
                s2.getRow(sllastRowNum).getCell(10).setCellType(CellType.STRING);//合格率（%）
                s2.getRow(6).getCell(9).setCellType(CellType.STRING);//合格率（%）
                s2.getRow(sllastRowNum).getCell(12).setCellType(CellType.STRING);
                s2.getRow(sllastRowNum).getCell(13).setCellType(CellType.STRING);
                map2.put("检测项目","沥青路面(右幅)");
                map2.put("规定值",s1.getRow(6).getCell(9).getStringCellValue());
                map2.put("检测点数",decf.format(Double.valueOf(s2.getRow(sllastRowNum).getCell(4).getStringCellValue())));
                map2.put("合格点数",decf.format(Double.valueOf(s2.getRow(sllastRowNum).getCell(8).getStringCellValue())));
                map2.put("合格率",df.format(Double.valueOf(s2.getRow(sllastRowNum).getCell(10).getStringCellValue())));
                map2.put("最大值",s2.getRow(sllastRowNum).getCell(12).getStringCellValue());
                map2.put("最小值",s2.getRow(sllastRowNum).getCell(13).getStringCellValue());
                mapList.add(map2);
            }/*else {
                map2.put("检测项目","沥青路面(右幅)");
                map2.put("规定值",0);
                map2.put("检测点数",0);
                map2.put("合格点数",0);
                map2.put("合格率",0);
                mapList.add(map2);
            }*/

            if (!sdzfsheet && proname.equals(s3.getRow(1).getCell(2).toString()) && htd.equals(s3.getRow(1).getCell(8).toString())){
                int sllastRowNum = s3.getLastRowNum();
                s3.getRow(sllastRowNum).getCell(4).setCellType(CellType.STRING);//检测点数
                s3.getRow(sllastRowNum).getCell(8).setCellType(CellType.STRING);//合格点数
                s3.getRow(sllastRowNum).getCell(10).setCellType(CellType.STRING);//合格率（%）
                s3.getRow(6).getCell(9).setCellType(CellType.STRING);//合格率（%）

                s3.getRow(sllastRowNum).getCell(12).setCellType(CellType.STRING);
                s3.getRow(sllastRowNum).getCell(13).setCellType(CellType.STRING);
                map3.put("检测项目","隧道路面(左幅)");
                map3.put("规定值",s3.getRow(6).getCell(9).getStringCellValue());
                map3.put("检测点数",decf.format(Double.valueOf(s3.getRow(sllastRowNum).getCell(4).getStringCellValue())));
                map3.put("合格点数",decf.format(Double.valueOf(s3.getRow(sllastRowNum).getCell(8).getStringCellValue())));
                map3.put("合格率",df.format(Double.valueOf(s3.getRow(sllastRowNum).getCell(10).getStringCellValue())));
                map3.put("最大值",s3.getRow(sllastRowNum).getCell(12).getStringCellValue());
                map3.put("最小值",s3.getRow(sllastRowNum).getCell(13).getStringCellValue());
                mapList.add(map3);
            }/*else {
                map3.put("检测项目","隧道路面(左幅)");
                map3.put("规定值",0);
                map3.put("检测点数",0);
                map3.put("合格点数",0);
                map3.put("合格率",0);
                mapList.add(map3);
            }*/

            if (!sdyfsheet && proname.equals(s4.getRow(1).getCell(2).toString()) && htd.equals(s4.getRow(1).getCell(8).toString())){
                int sllastRowNum = s4.getLastRowNum();
                s4.getRow(sllastRowNum).getCell(4).setCellType(CellType.STRING);//检测点数
                s4.getRow(sllastRowNum).getCell(8).setCellType(CellType.STRING);//合格点数
                s4.getRow(sllastRowNum).getCell(10).setCellType(CellType.STRING);//合格率（%）
                s4.getRow(6).getCell(9).setCellType(CellType.STRING);//合格率（%）
                s4.getRow(sllastRowNum).getCell(12).setCellType(CellType.STRING);
                s4.getRow(sllastRowNum).getCell(13).setCellType(CellType.STRING);

                map4.put("检测项目","隧道路面(右幅)");
                map4.put("规定值",s4.getRow(6).getCell(9).getStringCellValue());
                map4.put("检测点数",decf.format(Double.valueOf(s4.getRow(sllastRowNum).getCell(4).getStringCellValue())));
                map4.put("合格点数",decf.format(Double.valueOf(s4.getRow(sllastRowNum).getCell(8).getStringCellValue())));
                map4.put("合格率",df.format(Double.valueOf(s4.getRow(sllastRowNum).getCell(10).getStringCellValue())));
                map4.put("最大值",s4.getRow(sllastRowNum).getCell(12).getStringCellValue());
                map4.put("最小值",s4.getRow(sllastRowNum).getCell(13).getStringCellValue());
                mapList.add(map4);
            }/*else {
                map4.put("检测项目","隧道路面(右幅)");
                map4.put("规定值",0);
                map4.put("检测点数",0);
                map4.put("合格点数",0);
                map4.put("合格率",0);
                mapList.add(map4);
            }*/

            if (!zdsheet && proname.equals(s5.getRow(1).getCell(2).toString()) && htd.equals(s5.getRow(1).getCell(8).toString())){
                int sllastRowNum = s5.getLastRowNum();
                s5.getRow(sllastRowNum).getCell(4).setCellType(CellType.STRING);//检测点数
                s5.getRow(sllastRowNum).getCell(8).setCellType(CellType.STRING);//合格点数
                s5.getRow(sllastRowNum).getCell(10).setCellType(CellType.STRING);//合格率（%）
                s5.getRow(6).getCell(9).setCellType(CellType.STRING);//合格率（%）
                s5.getRow(sllastRowNum).getCell(12).setCellType(CellType.STRING);
                s5.getRow(sllastRowNum).getCell(13).setCellType(CellType.STRING);
                map5.put("检测项目","匝道路面");
                map5.put("规定值",s5.getRow(6).getCell(9).getStringCellValue());
                map5.put("检测点数",decf.format(Double.valueOf(s5.getRow(sllastRowNum).getCell(4).getStringCellValue())));
                map5.put("合格点数",decf.format(Double.valueOf(s5.getRow(sllastRowNum).getCell(8).getStringCellValue())));
                map5.put("合格率",df.format(Double.valueOf(s5.getRow(sllastRowNum).getCell(10).getStringCellValue())));
                map5.put("最大值",s4.getRow(sllastRowNum).getCell(12).getStringCellValue());
                map5.put("最小值",s4.getRow(sllastRowNum).getCell(13).getStringCellValue());
                mapList.add(map5);
            }/*else {
                map5.put("检测项目","匝道路面");
                map5.put("规定值",0);
                map5.put("检测点数",0);
                map5.put("合格点数",0);
                map5.put("合格率",0);
                mapList.add(map5);
            }*/
            return mapList;
        }
    }

    @Override
    public void exportLmssxs(HttpServletResponse response) {
        String fileName = "04沥青路面渗水系数实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcLmgcLmssxsVo()).finish();

    }

    @Override
    public void importLmssxs(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcLmgcLmssxsVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcLmgcLmssxsVo>(JjgFbgcLmgcLmssxsVo.class) {
                                @Override
                                public void handle(List<JjgFbgcLmgcLmssxsVo> dataList) {
                                    for(JjgFbgcLmgcLmssxsVo lmssxsVo: dataList)
                                    {
                                        JjgFbgcLmgcLmssxs lmssxs = new JjgFbgcLmgcLmssxs();
                                        BeanUtils.copyProperties(lmssxsVo,lmssxs);
                                        lmssxs.setCreatetime(new Date());
                                        lmssxs.setProname(commonInfoVo.getProname());
                                        lmssxs.setHtd(commonInfoVo.getHtd());
                                        lmssxs.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcLmgcLmssxsMapper.insert(lmssxs);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }

    @Override
    public int selectnum(String proname, String htd) {
        int selectnum = jjgFbgcLmgcLmssxsMapper.selectnum(proname, htd);
        return selectnum;
    }
}
