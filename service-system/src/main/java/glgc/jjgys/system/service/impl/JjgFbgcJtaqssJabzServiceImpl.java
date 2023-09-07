package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabz;
import glgc.jjgys.model.projectvo.jagc.JjgFbgcJtaqssJabzVo;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcJtaqssJabzMapper;
import glgc.jjgys.system.service.JjgFbgcJtaqssJabzService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.ss.usermodel.CellType;


/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-03-01
 */
@Service
public class JjgFbgcJtaqssJabzServiceImpl extends ServiceImpl<JjgFbgcJtaqssJabzMapper, JjgFbgcJtaqssJabz> implements JjgFbgcJtaqssJabzService {

    @Autowired
    private JjgFbgcJtaqssJabzMapper jjgFbgcJtaqssJabzMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;


    @Override
    public void importjabz(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcJtaqssJabzVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcJtaqssJabzVo>(JjgFbgcJtaqssJabzVo.class) {
                                @Override
                                public void handle(List<JjgFbgcJtaqssJabzVo> dataList) {
                                    for(JjgFbgcJtaqssJabzVo jabzVo: dataList)
                                    {
                                        JjgFbgcJtaqssJabz fbgcJtaqssJabz = new JjgFbgcJtaqssJabz();
                                        BeanUtils.copyProperties(jabzVo,fbgcJtaqssJabz);
                                        fbgcJtaqssJabz.setCreatetime(new Date());
                                        fbgcJtaqssJabz.setProname(commonInfoVo.getProname());
                                        fbgcJtaqssJabz.setHtd(commonInfoVo.getHtd());
                                        fbgcJtaqssJabz.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcJtaqssJabzMapper.insert(fbgcJtaqssJabz);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }

    @Override
    public void exportjabz(HttpServletResponse response) {
        String fileName = "01交安标志实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcJtaqssJabzVo()).finish();

    }

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获取数据
        QueryWrapper<JjgFbgcJtaqssJabz> wrapper = new QueryWrapper<>();
        wrapper.like("proname", proname);
        wrapper.like("htd", htd);
        wrapper.like("fbgc", fbgc);
        wrapper.orderByAsc("wz");
        List<JjgFbgcJtaqssJabz> data = jjgFbgcJtaqssJabzMapper.selectList(wrapper);
        //鉴定表要存放的路径
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "56交安标志.xlsx");
        //健壮性判断如果没有数据返回"请导入数据"
        if (data == null || data.size() == 0) {
            return;
        } else {
            //存放鉴定表的目录
            File fdir = new File(filepath + File.separator + proname + File.separator + htd);
            if (!fdir.exists()) {
                //创建文件根目录
                fdir.mkdirs();
            }
            File directory = new File("service-system/src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String name = "标志.xlsx";
            String path = reportPath + File.separator + name;
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            createTable(gettableNum(data.size()),wb);
            if(DBtoExcel(data,wb)){
                for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                    calculateOneSheet(wb.getSheetAt(j),wb);
                }

                for (int j = 0; j < wb.getNumberOfSheets(); j++) {   //表内公式  计算 显示结果
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
        /**
         * 7 26 45
         * 19
         *
         */
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String,Object>> mapList = new ArrayList<>();
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        List<Map<String, Object>> resultlist = new ArrayList<>();


        String sheetname = "标志";
        //获取鉴定表文件
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"56交安标志.xlsx");
        if(!f.exists()){
            return null;
        }else {
            //创建工作簿
            XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));
            //读取工作表
            XSSFSheet slSheet = xwb.getSheet(sheetname);
            for (int i = 7 ; i<slSheet.getPhysicalNumberOfRows();i+=19){
                Map<String,Object> jgmap1 = new HashMap<>();
                Map<String,Object> jgmap2 = new HashMap<>();
                Map<String,Object> jgmap3 = new HashMap<>();
                slSheet.getRow(i).getCell(0).setCellType(CellType.STRING);//项目
                slSheet.getRow(i).getCell(3).setCellType(CellType.STRING);//规定值或允许偏差
                slSheet.getRow(i).getCell(12).setCellType(CellType.STRING);//总点数
                slSheet.getRow(i).getCell(13).setCellType(CellType.STRING);//合格点数
                String xm1= slSheet.getRow(i).getCell(0).getStringCellValue();
                String yxpc1= slSheet.getRow(i).getCell(3).getStringCellValue();
                double zds1 = 0;
                double hgds1 = 0;
                if (!slSheet.getRow(i).getCell(12).getStringCellValue().equals("-") && !slSheet.getRow(i).getCell(12).getStringCellValue().equals("")){
                     zds1= Double.valueOf(slSheet.getRow(i).getCell(12).getStringCellValue());
                     hgds1= Double.valueOf(slSheet.getRow(i).getCell(13).getStringCellValue());
                }else{
                     zds1 = 0;
                     hgds1 = 0;
                }
                String zdsz1 = decf.format(zds1);
                String hgdsz1 = decf.format(hgds1);
                jgmap1.put("总点数",zdsz1);
                jgmap1.put("合格点数",hgdsz1);
                jgmap1.put("项目",xm1);
                jgmap1.put("ys","");
                jgmap1.put("规定值或允许偏差",yxpc1);
                mapList.add(jgmap1);

                slSheet.getRow(i+2).getCell(0).setCellType(CellType.STRING);
                slSheet.getRow(i+2).getCell(3).setCellType(CellType.STRING);
                slSheet.getRow(i+2).getCell(12).setCellType(CellType.STRING);
                slSheet.getRow(i+2).getCell(13).setCellType(CellType.STRING);
                String xm2= slSheet.getRow(i+2).getCell(0).getStringCellValue();
                String yxpc2= slSheet.getRow(i+2).getCell(3).getStringCellValue();
                double zds2 = 0;
                double hgds2 = 0;
                if (!slSheet.getRow(i+2).getCell(12).getStringCellValue().equals("-") && !slSheet.getRow(i+2).getCell(12).getStringCellValue().equals("")){
                     zds2= Double.valueOf(slSheet.getRow(i+2).getCell(12).getStringCellValue());
                }else {
                     zds2 = 0;
                }
                if (!slSheet.getRow(i+2).getCell(13).getStringCellValue().equals("-") && !slSheet.getRow(i+2).getCell(13).getStringCellValue().equals("")){
                    hgds2= Double.valueOf(slSheet.getRow(i+2).getCell(13).getStringCellValue());
                }else {
                    hgds2 = 0;
                }
                String zdsz2 = decf.format(zds2);
                String hgdsz2 = decf.format(hgds2);
                jgmap2.put("总点数",zdsz2);
                jgmap2.put("合格点数",hgdsz2);
                jgmap2.put("项目",xm2);
                jgmap2.put("ys","");
                jgmap2.put("规定值或允许偏差",yxpc2);
                mapList.add(jgmap2);


                slSheet.getRow(i+3).getCell(0).setCellType(CellType.STRING);
                slSheet.getRow(i+3).getCell(3).setCellType(CellType.STRING);
                slSheet.getRow(i+3).getCell(12).setCellType(CellType.STRING);
                slSheet.getRow(i+3).getCell(13).setCellType(CellType.STRING);
                String xm3= slSheet.getRow(i+3).getCell(0).getStringCellValue();
                String yxpc3= slSheet.getRow(i+3).getCell(3).getStringCellValue();
                double zds3 = 0;
                double hgds3 = 0;
                if (!slSheet.getRow(i+3).getCell(12).getStringCellValue().equals("-") && !slSheet.getRow(i+3).getCell(12).getStringCellValue().equals("")){
                     zds3= Double.valueOf(slSheet.getRow(i+3).getCell(12).getStringCellValue());
                }else {
                    zds3 = 0;
                }
                if (!slSheet.getRow(i+3).getCell(13).getStringCellValue().equals("-") && !slSheet.getRow(i+3).getCell(13).getStringCellValue().equals("")){
                    hgds3= Double.valueOf(slSheet.getRow(i+3).getCell(13).getStringCellValue());
                }else {
                    hgds3 = 0;
                }
                String zdsz3 = decf.format(zds3);
                String hgdsz3 = decf.format(hgds3);
                jgmap3.put("总点数",zdsz3);
                jgmap3.put("合格点数",hgdsz3);
                jgmap3.put("项目",xm3);
                jgmap3.put("ys","");
                jgmap3.put("规定值或允许偏差",yxpc3);
                mapList.add(jgmap3);

                for (int j = 0;j<10;j++){
                    Map<String,Object> jgmap4 = new HashMap<>();
                    slSheet.getRow(i+5+j).getCell(0).setCellType(CellType.STRING);
                    slSheet.getRow(i+5+j).getCell(1).setCellType(CellType.STRING);
                    slSheet.getRow(i+5+j).getCell(2).setCellType(CellType.STRING);
                    slSheet.getRow(i+5+j).getCell(3).setCellType(CellType.STRING);
                    slSheet.getRow(i+5+j).getCell(12).setCellType(CellType.STRING);
                    slSheet.getRow(i+5+j).getCell(13).setCellType(CellType.STRING);
                    String xm4= slSheet.getRow(i+5+j).getCell(0).getStringCellValue();
                    String yxpc4= slSheet.getRow(i+5+j).getCell(3).getStringCellValue();
                    double zds4 = 0;
                    double hgds4 = 0;
                    if (!slSheet.getRow(i+5+j).getCell(12).getStringCellValue().equals("-") && !slSheet.getRow(i+5+j).getCell(12).getStringCellValue().equals("")){
                        zds4= Double.valueOf(slSheet.getRow(i+5+j).getCell(12).getStringCellValue());
                    }else {
                        zds4 = 0;
                    }
                    if (!slSheet.getRow(i+5+j).getCell(13).getStringCellValue().equals("-") && !slSheet.getRow(i+5+j).getCell(13).getStringCellValue().equals("")){
                        hgds4= Double.valueOf(slSheet.getRow(i+5+j).getCell(13).getStringCellValue());
                    }else {
                        hgds4 = 0;
                    }

                    String zdsz4 = decf.format(zds4);
                    String hgdsz4 = decf.format(hgds4);
                    jgmap4.put("总点数",zdsz4);
                    jgmap4.put("合格点数",hgdsz4);
                    jgmap4.put("项目","标志面反光膜逆反射系数");
                    String val = getmergecell(slSheet,slSheet.getRow(i+5+j),1);
                    jgmap4.put("ys",val+slSheet.getRow(i+5+j).getCell(2).getStringCellValue());
                    jgmap4.put("规定值或允许偏差",yxpc4);
                    mapList.add(jgmap4);
                }

            }
            Map<String, Map<String, Object>> resultMap = new HashMap<>();

            for (Map<String, Object> data : mapList) {
                String project = (String) data.get("项目");
                String deviation = (String) data.get("规定值或允许偏差");
                String ys = (String) data.get("ys");
                double totalPoints = Double.valueOf(data.get("总点数").toString());
                double qualifiedPoints = Double.valueOf(data.get("合格点数").toString());
                // 构建唯一键
                String key = project + "-" + ys;

                // 检查是否存在相同的键
                if (resultMap.containsKey(key)) {
                    Map<String, Object> existingData = resultMap.get(key);
                    double existingTotalPoints = Double.valueOf(existingData.get("总点数").toString());
                    double existingQualifiedPoints = Double.valueOf( existingData.get("合格点数").toString());

                    existingData.put("总点数", existingTotalPoints + totalPoints);
                    existingData.put("合格点数", existingQualifiedPoints + qualifiedPoints);
                } else {
                    Map<String, Object> newData = new HashMap<>();
                    newData.put("项目", project);
                    newData.put("规定值或允许偏差", deviation);
                    newData.put("总点数", totalPoints);
                    newData.put("合格点数", qualifiedPoints);
                    newData.put("ys", ys);
                    resultMap.put(key, newData);
                }
            }
            // 转换为列表
            List<Map<String, Object>> result = new ArrayList<>(resultMap.values());
            for (Map<String, Object> map : result) {
                Map mape = new HashMap(map);
                double zds = Double.valueOf(map.get("总点数").toString());
                double hgds = Double.valueOf(map.get("合格点数").toString());
                String hgl="";
                if (zds != 0|| hgds !=0 ){
                    hgl = df.format(hgds/zds*100);
                }else {
                    hgl = "0";
                }
                mape.put("合格率",hgl);
                mape.put("规定值或允许偏差",map.get("ys").toString()+map.get("规定值或允许偏差").toString());
                resultlist.add(mape);
            }
        }
        return resultlist;
    }

    @Override
    public Map<String, Object> selectchs(String proname, String htd) {
        Map<String, Object> map = jjgFbgcJtaqssJabzMapper.selectchs(proname,htd);
        return map;
    }

    @Override
    public int selectnum(String proname, String htd) {
        int selectnum = jjgFbgcJtaqssJabzMapper.selectnum(proname, htd);
        return selectnum;
    }

    /**
     *
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    private String getmergecell(XSSFSheet sheet, XSSFRow row, int column) {
        String cellValue = "";

        // 判断单元格是否为合并单元格
        if (sheet instanceof XSSFSheet) {
            XSSFSheet xssfSheet = (XSSFSheet) sheet;
            int numMergedRegions = xssfSheet.getNumMergedRegions();

            for (int i = 0; i < numMergedRegions; i++) {
                CellRangeAddress region = xssfSheet.getMergedRegion(i);
                int firstRow = region.getFirstRow();
                int lastRow = region.getLastRow();
                int firstCol = region.getFirstColumn();
                int lastCol = region.getLastColumn();

                if (row.getRowNum() >= firstRow && row.getRowNum() <= lastRow && column >= firstCol && column <= lastCol) {
                    XSSFRow mergedRow = xssfSheet.getRow(firstRow);
                    if (mergedRow != null) {
                        XSSFCell mergedCell = mergedRow.getCell(firstCol);
                        if (mergedCell != null) {
                            // 获取合并单元格的值
                            switch (mergedCell.getCellType()) {
                                case XSSFCell.CELL_TYPE_STRING:
                                    cellValue = mergedCell.getStringCellValue();
                                    break;
                                case XSSFCell.CELL_TYPE_NUMERIC:
                                    cellValue = String.valueOf(mergedCell.getNumericCellValue());
                                    break;
                                case XSSFCell.CELL_TYPE_BOOLEAN:
                                    cellValue = String.valueOf(mergedCell.getBooleanCellValue());
                                    break;
                                case XSSFCell.CELL_TYPE_FORMULA:
                                    cellValue = mergedCell.getCellFormula();
                                    break;
                                default:
                                    cellValue = "";
                            }
                        }
                    }
                    break;
                }
            }
        }

        return cellValue;

    }

    /**
     *
     * @param sheet
     * @param xwb
     */
    private void calculateOneSheet(XSSFSheet sheet,XSSFWorkbook xwb) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        String offset = "";
        XSSFFormulaEvaluator evaluator = new XSSFFormulaEvaluator(xwb);

        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if ("立柱竖直度\n(mm/m)".equals(row.getCell(0).toString())) {
                offset = analysisOffset(row.getCell(3).toString());
                rowstart = row;
                rowend = sheet.getRow(i + 1);
                rowstart.getCell(12).setCellFormula(
                        "COUNT(" + rowstart.getCell(4).getReference() + ":"
                                + rowend.getCell(11).getReference() + ")");// =COUNT(E8:L9)
                rowstart.getCell(13).setCellFormula(
                        "COUNTIF(" + rowstart.getCell(4).getReference() + ":"
                                + rowend.getCell(11).getReference() + ",\""
                                + offset + "\")");// =COUNTIF(E8:L9,"<=3")
                if (evaluator.evaluate(row.getCell(12)).getNumberValue() != 0.0) {
                    rowstart.getCell(14).setCellFormula(
                            rowstart.getCell(13).getReference() + "/"
                                    + rowstart.getCell(12).getReference()
                                    + "*100");// =N8/M8*100
                } else {
                    row.getCell(12).setCellValue("-");
                    row.getCell(13).setCellValue("-");
                    row.getCell(14).setCellValue("-");
                }
                i++;
                continue;
            }
            if ("标志板净空\n（mm）".equals(row.getCell(0).toString())) {
                row.getCell(12).setCellFormula(
                        "COUNT(" + row.getCell(4).getReference() + ":"
                                + row.getCell(11).getReference() + ")");// =COUNT(E8:L9)
                int count = 0;
                for (int j = 0; j < 4; j++) {
                    if(row.getCell(4+2*j) != null && !"".equals(row.getCell(4+2*j).toString())){
                        try{
                            if(Double.valueOf(row.getCell(4+2*j).toString()).intValue() >= 5500&&Double.valueOf(row.getCell(4+2*j).toString()).intValue() <= 5600){
                                count ++;
                            }
                        }catch(Exception e){
                        }
                    }
                }
                row.getCell(13).setCellValue(count);
                if (evaluator.evaluate(row.getCell(12)).getNumberValue() != 0.0) {
                    row.getCell(14).setCellFormula(
                            row.getCell(13).getReference() + "/"
                                    + row.getCell(12).getReference() + "*100");// =N8/M8*100
                } else {
                    row.getCell(12).setCellValue("-");
                    row.getCell(13).setCellValue("-");
                    row.getCell(14).setCellValue("-");
                }
                continue;
            }
            if ("标志板厚度\n（mm）".equals(row.getCell(0).toString())) {
                offset = analysisOffset(row.getCell(3).toString());
                rowstart = row;
                do {
                    rowend = sheet.getRow(i);
                    i++;
                } while (rowend.getCell(0)!= null && !"".equals(rowend.getCell(0).toString()) && !rowend.getCell(0).toString().contains("反光膜逆反射系数"));
                i -= 2;
                rowstart.getCell(12).setCellFormula(
                        "COUNT(" + rowstart.getCell(4).getReference() + ":"
                                + rowend.getCell(11).getReference() + ")");// =COUNT(E8:L9)
                if(rowstart.getCell(4+11) == null){
                    rowstart.createCell(4+11);
                }
                if(rowend.getCell(11+11) == null){
                    rowend.createCell(11+11);
                }
                rowstart.getCell(13).setCellFormula(
                        "COUNT(" + rowstart.getCell(4+11).getReference() + ":"
                                + rowend.getCell(11+11).getReference() + ")");// =COUNT(E8:L9)
                if (evaluator.evaluate(row.getCell(12)).getNumberValue() != 0.0) {
                    rowstart.getCell(14).setCellFormula(
                            rowstart.getCell(13).getReference() + "/"
                                    + rowstart.getCell(12).getReference()
                                    + "*100");// =N8/M8*100
                } else {
                    row.getCell(12).setCellValue("-");
                    row.getCell(13).setCellValue("-");
                    row.getCell(14).setCellValue("-");
                }
                continue;
            }
            if (row.getCell(0)!= null && !"".equals(row.getCell(0).toString()) && row.getCell(0).toString().contains("反光膜逆反射系数")) {
                while (row.getCell(3) != null && !"备注：".equals(row.getCell(0).toString())) {
                    if(!"".equals(row.getCell(3).toString())){
                        offset = analysisOffset(row.getCell(3).toString());
                        row.getCell(12).setCellFormula(
                                "COUNT(" + row.getCell(4).getReference() + ":"
                                        + row.getCell(11).getReference() + ")");// =COUNT(E8:L9)
                        row.getCell(13).setCellFormula(
                                "COUNTIF(" + row.getCell(4).getReference() + ":"
                                        + row.getCell(11).getReference() + ",\""
                                        + offset + "\")");// =COUNTIF(E8:L9,"<=3")

                        if (evaluator.evaluate(row.getCell(12)).getNumberValue() != 0.0) {
                            row.getCell(14).setCellFormula(
                                    row.getCell(13).getReference() + "/"
                                            + row.getCell(12).getReference()
                                            + "*100");// =N8/M8*100
                        }
                    } else {
                        row.getCell(12).setCellValue("-");
                        row.getCell(13).setCellValue("-");
                        row.getCell(14).setCellValue("-");
                    }
                    row = sheet.getRow(++i);
                }
            }
        }
    }

    /**
     *
     * @param offset
     * @return
     */
    private String analysisOffset(String offset) {
        String result = "";
        String operator = offset.substring(0, 1);
        switch (operator) {
            case "≤":
                result = "<=" + offset.substring(1);
                break;
            case "≥":
                result = ">=" + offset.substring(1);
                break;
            default:
                break;
        }
        return result;
    }

    /**
     *
     * @param data
     * @param wb
     * @return
     * @throws ParseException
     */
    private boolean DBtoExcel(List<JjgFbgcJtaqssJabz> data,XSSFWorkbook wb) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        XSSFSheet sheet = wb.getSheet("标志");
        sheet.getRow(1).getCell(3).setCellValue(data.get(0).getProname());
        sheet.getRow(1).getCell(11).setCellValue(data.get(0).getHtd());
        sheet.getRow(2).getCell(3).setCellValue(data.get(0).getFbgc());
        sheet.getRow(2).getCell(11).setCellValue("标志");
        sheet.getRow(3).getCell(2).setCellValue(data.get(0).getJcsj());
        String date = simpleDateFormat.format(data.get(0).getJcsj());
        for(int i =1; i < data.size(); i++){
            Date jcsj = data.get(i).getJcsj();
            date = JjgFbgcCommonUtils.getLastDate(date, simpleDateFormat.format(jcsj));
        }
        sheet.getRow(3).getCell(11).setCellValue(date);
        int tableNum = 0;
        for(int i =0; i < data.size(); i++){
            fillCommonCellData(sheet, tableNum, 4+(i%4)*2, data.get(i));
            if((i+1)%4 == 0){
                tableNum++;
            }
        }
        return true;
    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param index
     * @param row
     */
    private void fillCommonCellData(XSSFSheet sheet, int tableNum, int index, JjgFbgcJtaqssJabz row) {
        sheet.getRow(tableNum*(17+2)+5).getCell(index).setCellValue(row.getWz());
        sheet.getRow(tableNum*(17+2)+6).getCell(index).setCellValue(row.getLzlx());
        if(row.getLzlx().contains("单柱") ){
            if(!"".equals(row.getSzdyxps()) && row.getSzdyxps() != null){
                if(row.getSzdyxps().contains("≤")){
                    sheet.getRow(tableNum*(17+2)+7).getCell(3).setCellValue(row.getSzdyxps());
                }
                else{
                    sheet.getRow(tableNum*(17+2)+7).getCell(3).setCellValue("≤"+row.getSzdyxps());
                }
            }
            if(!"".equals(row.getFx1scz1()) && row.getFx1scz1() != null){
                sheet.addMergedRegion(new CellRangeAddress(tableNum*(17+2)+7, tableNum*(17+2)+7, index, index+1));
                sheet.getRow(tableNum*(17+2)+7).getCell(index).setCellValue(Double.valueOf(row.getFx1scz1()));
            }
            if(!"".equals(row.getFx2scz1()) && row.getFx2scz1() != null){
                sheet.addMergedRegion(new CellRangeAddress(tableNum*(17+2)+8, tableNum*(17+2)+8, index, index+1));
                sheet.getRow(tableNum*(17+2)+8).getCell(index).setCellValue(Double.valueOf(row.getFx2scz1()));
            }
        }
        else{
            if(!"".equals(row.getSzdyxps()) && row.getSzdyxps() != null){
                if(row.getSzdyxps().contains("≤")){
                    sheet.getRow(tableNum*(17+2)+7).getCell(3).setCellValue(row.getSzdyxps());
                }
                else{
                    sheet.getRow(tableNum*(17+2)+7).getCell(3).setCellValue("≤"+row.getSzdyxps());
                }
            }
            if(!"".equals(row.getFx1scz1()) && row.getFx1scz1() != null){
                sheet.getRow(tableNum*(17+2)+7).getCell(index).setCellValue(Double.valueOf(row.getFx1scz1()));
            }
            if(!"".equals(row.getFx1scz2()) && row.getFx1scz2() != null){
                sheet.getRow(tableNum*(17+2)+7).getCell(index+1).setCellValue(Double.valueOf(row.getFx1scz2()));
            }else {
                sheet.addMergedRegion(new CellRangeAddress(tableNum*(17+2)+7, tableNum*(17+2)+7, index, index+1));
            }
            if(!"".equals(row.getFx2scz1()) && row.getFx2scz1() !=null) {
                sheet.getRow(tableNum*(17+2)+8).getCell(index).setCellValue(Double.valueOf(row.getFx2scz1()));
            }
            if(!"".equals(row.getFx2scz2()) && row.getFx2scz2()!=null){
                sheet.getRow(tableNum*(17+2)+8).getCell(index+1).setCellValue(Double.valueOf(row.getFx2scz2()));
            }else {
                sheet.addMergedRegion(new CellRangeAddress(tableNum*(17+2)+8, tableNum*(17+2)+8, index, index+1));
            }
        }
        //净空允许偏差
        if(!"".equals(row.getJkgdz()) && row.getJkgdz() != null){
            if(!row.getJkgdz().contains("≥")){
                row.setJkgdz("≥"+Double.valueOf(row.getJkgdz()).intValue());
            }
            sheet.getRow(tableNum*(17+2)+9).getCell(3).setCellValue("+100,0");
        }
        //净空实测值
        if(!"".equals(row.getJkscz()) && row.getJkscz() != null){
            sheet.getRow(tableNum*(17+2)+9).getCell(index).setCellValue(Double.valueOf(row.getJkscz()));
        }else{
            sheet.getRow(tableNum*(17+2)+9).getCell(index).setCellValue("-");
        }
        //厚度允许偏差
        if(!"".equals(row.getHdyxps()) && row.getHdyxps() != null){
            if(!row.getHdyxps().contains("≥")){
                row.setHdyxps("≥"+row.getHdyxps());
                //row.getHdyxps() = "≥"+row.getHdyxps();
            }
            if(!sheet.getRow(tableNum*(17+2)+10).getCell(3).toString().contains(row.getHdyxps())){
                if(!"".equals(sheet.getRow(tableNum*(17+2)+10).getCell(3).toString())){
                    sheet.getRow(tableNum*(17+2)+10).getCell(3).setCellValue(
                            sheet.getRow(tableNum*(17+2)+10).getCell(3).toString()+","+row.getHdyxps());
                }
                else{
                    sheet.getRow(tableNum*(17+2)+10).getCell(3).setCellValue(row.getHdyxps());
                }
            }
        }

        //厚度实测值1
        if(!"".equals(row.getHdcz1()) && row.getHdcz1() != null){
            sheet.getRow(tableNum*(17+2)+10).getCell(index).setCellValue(Double.valueOf(row.getHdcz1()));
            double b = Double.valueOf(row.getHdcz1());
            double c = Double.valueOf(row.getHdyxps().substring(1));
            if(b >= c){
                sheet.getRow(tableNum*(17+2)+10).createCell(index+11).setCellValue(1);
            }

        }else{
            sheet.getRow(tableNum*(17+2)+10).getCell(index).setCellValue("-");
        }
        //厚度实测值2
        if(!"".equals(row.getHdcz2()) && row.getHdcz2() != null){
            sheet.getRow(tableNum*(17+2)+11).getCell(index).setCellValue(Double.valueOf(row.getHdcz2()));
            if(Double.valueOf(row.getHdcz2()) >= Double.valueOf(row.getHdyxps().substring(1))){
                sheet.getRow(tableNum*(17+2)+11).createCell(index+11).setCellValue(1);
            }
        }else{
            sheet.getRow(tableNum*(17+2)+11).getCell(index).setCellValue("-");
        }

        //白色V类
        if(!"".equals(row.getBsvlyxps()) && row.getBsvlyxps()!=null){
            sheet.getRow(tableNum*(17+2)+12).getCell(3).setCellValue(row.getBsvlyxps());
            /*if(row.getBsvlyxps().contains("≥")){
                sheet.getRow(tableNum*(17+2)+12).getCell(3).setCellValue(row.getBsvlyxps());
            }
            else{
                sheet.getRow(tableNum*(17+2)+12).getCell(3).setCellValue("≥"+row.getBsvlyxps());
            }*/
        }
        if(!"".equals(row.getBsvlscz1()) && row.getBsvlscz1()!=null){
            sheet.getRow(tableNum*(17+2)+12).getCell(index).setCellValue(Double.valueOf(row.getBsvlscz1()));
        }else{
            sheet.getRow(tableNum*(17+2)+12).getCell(index).setCellValue("-");
        }

        if(!"".equals(row.getBsvlscz2()) && row.getBsvlscz2() != null){
            sheet.getRow(tableNum*(17+2)+12).getCell(index+1).setCellValue(Double.valueOf(row.getBsvlscz2()));
        }else{
            sheet.getRow(tableNum*(17+2)+12).getCell(index+1).setCellValue("-");
        }
        //白色IV类
        if(!"".equals(row.getBswlyxps()) && row.getBswlyxps()!=null){
            sheet.getRow(tableNum*(17+2)+13).getCell(3).setCellValue(row.getBswlyxps());
           /* if((row.getBswlyxps()).contains("≥")){
                sheet.getRow(tableNum*(17+2)+13).getCell(3).setCellValue(row.getBswlyxps());
            }
            else{
                sheet.getRow(tableNum*(17+2)+13).getCell(3).setCellValue("≥"+row.getBswlyxps());
            }*/
        }
        if(!"".equals(row.getBswlscz1()) && row.getBswlscz1()!=null){
            sheet.getRow(tableNum*(17+2)+13).getCell(index).setCellValue(Double.valueOf(row.getBswlscz1()));
        }else{
            sheet.getRow(tableNum*(17+2)+13).getCell(index).setCellValue("-");
        }
        if(!"".equals(row.getBswlscz2()) && row.getBswlscz2()!=null){
            sheet.getRow(tableNum*(17+2)+13).getCell(index+1).setCellValue(Double.valueOf(row.getBswlscz2()));
        }else{
            sheet.getRow(tableNum*(17+2)+13).getCell(index+1).setCellValue("-");
        }

        //绿色V类
        if(!"".equals(row.getLsvlyxps()) && row.getLsvlyxps()!=null){
            sheet.getRow(tableNum*(17+2)+14).getCell(3).setCellValue(row.getLsvlyxps());
            /*if(row.getLsvlyxps().contains("≥")){
                sheet.getRow(tableNum*(17+2)+14).getCell(3).setCellValue(row.getLsvlyxps());
            }
            else{
                sheet.getRow(tableNum*(17+2)+14).getCell(3).setCellValue("≥"+row.getLsvlyxps());
            }*/
        }
        if(!"".equals(row.getLsvlscz1()) && row.getLsvlscz1()!=null){
            sheet.getRow(tableNum*(17+2)+14).getCell(index).setCellValue(Double.valueOf(row.getLsvlscz1()));
        }else{
            sheet.getRow(tableNum*(17+2)+14).getCell(index).setCellValue("-");
        }
        if(!"".equals(row.getLsvlscz2()) && row.getLsvlscz2()!=null){
            sheet.getRow(tableNum*(17+2)+14).getCell(index+1).setCellValue(Double.valueOf(row.getLsvlscz2()));
        }else{
            sheet.getRow(tableNum*(17+2)+14).getCell(index+1).setCellValue("-");
        }
        //绿色IV类
        if(!"".equals(row.getLswlyxps()) && row.getLswlyxps()!=null){
            sheet.getRow(tableNum*(17+2)+15).getCell(3).setCellValue(row.getLswlyxps());
           /* if(row.getLswlyxps().contains("≥")){
                sheet.getRow(tableNum*(17+2)+15).getCell(3).setCellValue(row.getLswlyxps());
            }
            else{
                sheet.getRow(tableNum*(17+2)+15).getCell(3).setCellValue("≥"+row.getLswlyxps());
            }*/
        }
        if(!"".equals(row.getLswlscz1()) && row.getLswlscz1()!=null){
            sheet.getRow(tableNum*(17+2)+15).getCell(index).setCellValue(Double.valueOf(row.getLswlscz1()));
        }else{
            sheet.getRow(tableNum*(17+2)+15).getCell(index).setCellValue("-");
        }
        if(!"".equals(row.getLswlscz2()) && row.getLswlscz2()!=null){
            sheet.getRow(tableNum*(17+2)+15).getCell(index+1).setCellValue(Double.valueOf(row.getLswlscz2()));
        }else{
            sheet.getRow(tableNum*(17+2)+15).getCell(index+1).setCellValue("-");
        }

        //黄色V类
        if(!"".equals(row.getHsvlyxps()) && row.getHsvlyxps()!=null){
            sheet.getRow(tableNum*(17+2)+16).getCell(3).setCellValue(row.getHsvlyxps());
           /* if(row.getHsvlyxps().contains("≥")){
                sheet.getRow(tableNum*(17+2)+16).getCell(3).setCellValue(row.getHsvlyxps());
            }
            else{
                sheet.getRow(tableNum*(17+2)+16).getCell(3).setCellValue("≥"+row.getHsvlyxps());
            }*/
        }
        if(!"".equals(row.getHsvlscz1()) && row.getHsvlscz1()!=null){
            sheet.getRow(tableNum*(17+2)+16).getCell(index).setCellValue(Double.valueOf(row.getHsvlscz1()));
        }else{
            sheet.getRow(tableNum*(17+2)+16).getCell(index).setCellValue("-");
        }
        if(!"".equals(row.getHsvlscz2()) && row.getHsvlscz2()!=null){
            sheet.getRow(tableNum*(17+2)+16).getCell(index+1).setCellValue(Double.valueOf(row.getHsvlscz2()));
        }else{
            sheet.getRow(tableNum*(17+2)+16).getCell(index+1).setCellValue("-");
        }
        //黄色IV类
        if(!"".equals(row.getHswlyxps()) && row.getHswlyxps()!=null){
            sheet.getRow(tableNum*(17+2)+17).getCell(3).setCellValue(row.getHswlyxps());
            /*if(row.getHswlyxps().contains("≥")){
                sheet.getRow(tableNum*(17+2)+17).getCell(3).setCellValue(row.getHswlyxps());
            }
            else{
                sheet.getRow(tableNum*(17+2)+17).getCell(3).setCellValue("≥"+row.getHswlyxps());
            }*/
        }
        if(!"".equals(row.getHswlscz1()) &&row.getHswlscz1()!=null){
            sheet.getRow(tableNum*(17+2)+17).getCell(index).setCellValue(Double.valueOf(row.getHswlscz1()));
        }else{
            sheet.getRow(tableNum*(17+2)+17).getCell(index).setCellValue("-");
        }
        if(!"".equals(row.getHswlscz2()) && row.getHswlscz2()!=null){
            sheet.getRow(tableNum*(17+2)+17).getCell(index+1).setCellValue(Double.valueOf(row.getHswlscz2()));
        }else{
            sheet.getRow(tableNum*(17+2)+17).getCell(index+1).setCellValue("-");
        }

        //蓝色V类
        if(!"".equals(row.getLsvlyxps()) && row.getLsvlyxps()!=null){
            sheet.getRow(tableNum*(17+2)+18).getCell(3).setCellValue(row.getLsvlyxps());
            /*if(row.getLsvlyxps().contains("≥")){
                sheet.getRow(tableNum*(17+2)+18).getCell(3).setCellValue(row.getLsvlyxps());
            }
            else{
                sheet.getRow(tableNum*(17+2)+18).getCell(3).setCellValue("≥"+row.getLsvlyxps());
            }*/
        }
        if(!"".equals(row.getLsvlscz1()) && row.getLsvlscz1()!=null){
            sheet.getRow(tableNum*(17+2)+18).getCell(index).setCellValue(Double.valueOf(row.getLsvlscz1()));
        }else{
            sheet.getRow(tableNum*(17+2)+18).getCell(index).setCellValue("-");
        }
        if(!"".equals(row.getLsvlscz2()) && row.getLsvlscz2()!=null){
            sheet.getRow(tableNum*(17+2)+18).getCell(index+1).setCellValue(Double.valueOf(row.getLsvlscz2()));
        }else{
            sheet.getRow(tableNum*(17+2)+16).getCell(index+1).setCellValue("-");
        }
        //蓝色IV类
        if(!"".equals(row.getLswlyxps()) && row.getLswlyxps()!=null){
            sheet.getRow(tableNum*(17+2)+19).getCell(3).setCellValue(row.getLswlyxps());
//            if(row.getLswlyxps().contains("≥")){
//                sheet.getRow(tableNum*(17+2)+19).getCell(3).setCellValue(row.getLswlyxps());
//            }
//            else{
//                sheet.getRow(tableNum*(17+2)+19).getCell(3).setCellValue("≥"+row.getLswlyxps());
//            }
        }
        if(!"".equals(row.getLswlscz1()) && row.getLswlscz1()!=null){
            sheet.getRow(tableNum*(17+2)+19).getCell(index).setCellValue(Double.valueOf(row.getLswlscz1()));
        }else{
            sheet.getRow(tableNum*(17+2)+19).getCell(index).setCellValue("-");
        }
        if(!"".equals(row.getLswlscz2()) && row.getLswlscz2()!=null){
            sheet.getRow(tableNum*(17+2)+19).getCell(index+1).setCellValue(Double.valueOf(row.getLswlscz2()));
        }else{
            sheet.getRow(tableNum*(17+2)+19).getCell(index+1).setCellValue("-");
        }

        //红色V类
        if(!"".equals(row.getRsvlyxps()) && row.getRsvlyxps()!=null){
            sheet.getRow(tableNum*(17+2)+20).getCell(3).setCellValue(row.getRsvlyxps());
            /*if(row.getRsvlyxps().contains("≥")){
                sheet.getRow(tableNum*(17+2)+20).getCell(3).setCellValue(row.getRsvlyxps());
            }
            else{
                sheet.getRow(tableNum*(17+2)+20).getCell(3).setCellValue("≥"+row.getRsvlyxps());
            }*/
        }
        if(!"".equals(row.getRsvlscz1()) && row.getRsvlscz1()!=null){
            sheet.getRow(tableNum*(17+2)+20).getCell(index).setCellValue(Double.valueOf(row.getRsvlscz1()));
        }else{
            sheet.getRow(tableNum*(17+2)+20).getCell(index).setCellValue("-");
        }
        if(!"".equals(row.getRsvlscz2()) && row.getRsvlscz2() != null){
            sheet.getRow(tableNum*(17+2)+20).getCell(index+1).setCellValue(Double.valueOf(row.getRsvlscz2()));
        }else{
            sheet.getRow(tableNum*(17+2)+20).getCell(index+1).setCellValue("-");
        }
        //红色IV类
        if(!"".equals(row.getRswlyxps()) && row.getRswlyxps()!=null){
            sheet.getRow(tableNum*(17+2)+21).getCell(3).setCellValue(row.getRswlyxps());
            /*if(row.getRswlyxps().contains("≥")){
                sheet.getRow(tableNum*(17+2)+21).getCell(3).setCellValue(row.getRswlyxps());
            }
            else{
                sheet.getRow(tableNum*(17+2)+21).getCell(3).setCellValue("≥"+row.getRswlyxps());
            }*/
        }
        if(!"".equals(row.getRswlscz1()) && row.getRswlscz1()!=null){
            sheet.getRow(tableNum*(17+2)+21).getCell(index).setCellValue(Double.valueOf(row.getRswlscz1()));
        }else{
            sheet.getRow(tableNum*(17+2)+21).getCell(index).setCellValue("-");
        }
        if(!"".equals(row.getRswlscz2()) && row.getRswlscz2()!=null){
            sheet.getRow(tableNum*(17+2)+21).getCell(index+1).setCellValue(Double.valueOf(row.getRswlscz2()));
        }else{
            sheet.getRow(tableNum*(17+2)+21).getCell(index+1).setCellValue("-");
        }


    }

    /**
     *
     * @param tableNum
     * @param wb
     */
    private void createTable(int tableNum,XSSFWorkbook wb) {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            RowCopy.copyRows(wb, "标志", "标志", 4, 20+2, (i - 1) * (17+2) + 21+2);
        }
        if(record > 1)
            wb.setPrintArea(wb.getSheetIndex("标志"), 0, 14, 0, record * (17+2)+3);
    }

    /**
     *
     * @param size
     * @return
     */
    private int gettableNum(int size) {
        return size%4==0?size/4:size/4+1;
    }
}
