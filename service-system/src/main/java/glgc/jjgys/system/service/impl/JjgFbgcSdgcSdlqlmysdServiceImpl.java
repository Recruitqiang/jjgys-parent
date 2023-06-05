package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLmgcLqlmysd;
import glgc.jjgys.model.project.JjgFbgcSdgcLmssxs;
import glgc.jjgys.model.project.JjgFbgcSdgcSdlqlmysd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.lmgc.JjgFbgcLmgcLqlmysdVo;
import glgc.jjgys.model.projectvo.sdgc.JjgFbgcSdgcSdlqlmysdVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcSdgcSdlqlmysdMapper;
import glgc.jjgys.system.service.JjgFbgcSdgcSdlqlmysdService;
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
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-05-07
 */
@Service
public class JjgFbgcSdgcSdlqlmysdServiceImpl extends ServiceImpl<JjgFbgcSdgcSdlqlmysdMapper, JjgFbgcSdgcSdlqlmysd> implements JjgFbgcSdgcSdlqlmysdService {

    @Autowired
    private JjgFbgcSdgcSdlqlmysdMapper jjgFbgcSdgcSdlqlmysdMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        List<Map<String,Object>> sdmclist = jjgFbgcSdgcSdlqlmysdMapper.selectsdmc(proname,htd,fbgc);
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
     */
    private void DBtoExcelsd(String proname, String htd, String fbgc, String sdmc) throws IOException {
        XSSFWorkbook wb = null;
        QueryWrapper<JjgFbgcSdgcSdlqlmysd> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.like("sdmc",sdmc);
        wrapper.orderByAsc("zh");
        List<JjgFbgcSdgcSdlqlmysd> data = jjgFbgcSdgcSdlqlmysdMapper.selectList(wrapper);

        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"43隧道沥青路面压实度-"+sdmc+".xlsx");
        if (data == null || data.size()==0 ){
            return;
        }else {
            //桥的数据是放在路面工作薄中
            List<JjgFbgcSdgcSdlqlmysd> zxzfdata = jjgFbgcSdgcSdlqlmysdMapper.selectzxzf(proname,htd,fbgc,sdmc);
            List<JjgFbgcSdgcSdlqlmysd> zxyfdata = jjgFbgcSdgcSdlqlmysdMapper.selectzxyf(proname,htd,fbgc,sdmc);
            List<JjgFbgcSdgcSdlqlmysd> sdzfdata = jjgFbgcSdgcSdlqlmysdMapper.selectsdzf(proname,htd,fbgc,sdmc);
            List<JjgFbgcSdgcSdlqlmysd> sdyfdata = jjgFbgcSdgcSdlqlmysdMapper.selectsdyf(proname,htd,fbgc,sdmc);
            List<JjgFbgcSdgcSdlqlmysd> zddata = jjgFbgcSdgcSdlqlmysdMapper.selectzd(proname,htd,fbgc,sdmc);
            List<JjgFbgcSdgcSdlqlmysd> ljxdata = jjgFbgcSdgcSdlqlmysdMapper.selectljx(proname,htd,fbgc,sdmc);
            List<JjgFbgcSdgcSdlqlmysd> ljxsddata = jjgFbgcSdgcSdlqlmysdMapper.selectljxsd(proname,htd,fbgc,sdmc);


            File fdir = new File(filepath + File.separator + proname + File.separator + htd);
            if (!fdir.exists()) {
                //创建文件根目录
                fdir.mkdirs();
            }
            File directory = new File("service-system/src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String path = reportPath + File.separator + "沥青路面压实度.xlsx";
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);


            //沥青路面压实度主线左幅
            lqlmysdzx(wb, zxzfdata, "沥青路面压实度左幅", "路面面层（主线左幅）");
            //沥青路面压实度主线右幅
            lqlmysdzx(wb, zxyfdata, "沥青路面压实度右幅", "路面面层（主线右幅）");
            //沥青路面压实度隧道左幅
            lqlmysdOther(wb, sdzfdata, "隧道左幅", "隧道路面");
            //沥青路面压实度隧道右幅
            lqlmysdOther(wb, sdyfdata, "隧道右幅", "隧道路面");
            //沥青路面压实度匝道
            lqlmysdOther(wb, zddata, "沥青路面压实度匝道", "匝道路面");
            //沥青路面压实度连接线
            lqlmysdOther(wb, ljxdata, "连接线", "连接线");
            //沥青路面压实度连接线隧道
            lqlmysdOther(wb, ljxsddata, "连接线隧道", "连接线隧道");
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                if (shouldBeCalculate(wb.getSheetAt(j))) {
                    calculateOneSheet(wb.getSheetAt(j));
                    JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
                }
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
        }
        wb.close();
    }

    /**
     * 
     * @param wb
     * @param data
     * @param sheetname
     * @param fbgcname
     */
    private void lqlmysdzx(XSSFWorkbook wb, List<JjgFbgcSdgcSdlqlmysd> data, String sheetname, String fbgcname) {
        if (data.size() > 0 && data != null){
            createZXZTable(wb,getZXZtableNum(data.size()),sheetname);
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
            XSSFSheet sheet = wb.getSheet(sheetname);
            String type = data.get(0).getLqs();
            String zh = data.get(0).getZh();
            sheet.getRow(1).getCell(2).setCellValue(data.get(0).getProname());
            sheet.getRow(1).getCell(9).setCellValue(data.get(0).getHtd());
            sheet.getRow(2).getCell(2).setCellValue(data.get(0).getFbgc());
            sheet.getRow(2).getCell(9).setCellValue(fbgcname);
            sheet.getRow(3).getCell(2).setCellValue(data.get(0).getLmlx());
            String date = simpleDateFormat.format(data.get(0).getJcsj());
            sheet.getRow(3).getCell(9).setCellValue(date);

            int index = 6;

            for(int i =0; i < data.size(); i++){
                if (zh.equals(data.get(i).getZh())){
                    if(type.contains("隧")){
                        i+=2;
                        index += 3;
                        sheet.addMergedRegion(new CellRangeAddress(index-3, index-1, 0, 1));
                        sheet.getRow(index-3).getCell(0).setCellValue(zh+type);
                        sheet.addMergedRegion(new CellRangeAddress(index-3, index-1, 2, 8));
                        sheet.getRow(index-3).getCell(2).setCellValue("见隧道路面压实度鉴定表");

                    }else {
                        for(int j = 0; j < 3; j++){
                            if(i+j < data.size()){
                                if("上面层".equals(data.get(i+j).getCw())){
                                    sheet.getRow(index).getCell(2).setCellValue(data.get(i+j).getCw());
                                    fillLmzfCommonCellData(sheet, index, data.get(i+j));
                                    //flag[0] = true;
                                }
                                if("中面层".equals(data.get(i+j).getCw())){
                                    sheet.getRow(index+1).getCell(2).setCellValue(data.get(i+j).getCw());
                                    fillLmzfCommonCellData(sheet, index+1, data.get(i+j));
                                    //flag[1] = true;
                                }
                                if("下面层".equals(data.get(i+j).getCw())){
                                    sheet.getRow(index+2).getCell(2).setCellValue(data.get(i+j).getCw());
                                    fillLmzfCommonCellData(sheet, index+2, data.get(i+j));
                                    //flag[2] = true;
                                }
                            }else {
                                break;
                            }
                        }
                        i+=2;
                        index+=3;
                        sheet.addMergedRegion(new CellRangeAddress(index-3, index-1, 0, 0));
                        sheet.getRow(index-3).getCell(0).setCellValue(zh);
                        sheet.addMergedRegion(new CellRangeAddress(index-3, index-1, 1, 1));
                    }
                }else {
                    type = data.get(i).getLqs();
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
     * @param zfdata
     */
    private void fillLmzfCommonCellData(XSSFSheet sheet, int index, JjgFbgcSdgcSdlqlmysd zfdata) {
        sheet.getRow(index).getCell(0).setCellValue(zfdata.getZh());
        sheet.getRow(index).getCell(1).setCellValue(zfdata.getQywz());
        //sheet.getRow(index).getCell(2).setCellValue(zfdata.getCw());
        if(!"".equals(zfdata.getGzsjzl()) && zfdata.getGzsjzl() != null){
            sheet.getRow(index).getCell(3).setCellValue(Double.parseDouble(zfdata.getGzsjzl()));
        }else {
            sheet.getRow(index).getCell(3).setCellValue("-");
        }
        if(!"".equals(zfdata.getSjszzl()) && zfdata.getSjszzl()!=null){
            sheet.getRow(index).getCell(4).setCellValue(Double.parseDouble(zfdata.getSjszzl()));
        }else {
            sheet.getRow(index).getCell(4).setCellValue("-");
        }
        if(!"".equals(zfdata.getSjbgzl()) && zfdata.getSjbgzl()!=null){
            sheet.getRow(index).getCell(5).setCellValue(Double.parseDouble(zfdata.getSjbgzl()));
        }else {
            sheet.getRow(index).getCell(5).setCellValue("-");
        }
        if(!"".equals(zfdata.getSysbzmd()) && zfdata.getSysbzmd()!=null){
            sheet.getRow(index).getCell(7).setCellValue(Double.parseDouble(zfdata.getSysbzmd()));
        }else {
            sheet.getRow(index).getCell(7).setCellValue("-");
        }

        if(!"".equals(zfdata.getZdllmd()) && zfdata.getZdllmd()!=null){
            sheet.getRow(index).getCell(8).setCellValue(Double.parseDouble(zfdata.getZdllmd()));
        }else {
            sheet.getRow(index).getCell(8).setCellValue("-");
        }

        if(!"".equals(zfdata.getSysbzmdgdz()) && zfdata.getSysbzmdgdz()!=null){
            sheet.getRow(index).getCell(11).setCellValue(Double.parseDouble(zfdata.getSysbzmdgdz()));
        }else {
            sheet.getRow(index).getCell(11).setCellValue("-");
        }

        if(!"".equals(zfdata.getZdllmdgdz()) && zfdata.getZdllmdgdz()!=null){
            sheet.getRow(index).getCell(12).setCellValue(Double.parseDouble(zfdata.getZdllmdgdz()));
        }else {
            sheet.getRow(index).getCell(12).setCellValue("-");
        }
        sheet.getRow(sheet.getLastRowNum()-1).getCell(2).setCellValue(Double.parseDouble(zfdata.getSysbzmdgdz()));
        sheet.getRow(sheet.getLastRowNum()-1).getCell(8).setCellValue(Double.parseDouble(zfdata.getZdllmdgdz()));
    }

    /**
     *
     * @param sheet
     */
    private void calculateOneSheet(XSSFSheet sheet) {
        XSSFRow row = null;
        for (int i = sheet.getFirstRowNum(); i <= sheet
                .getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            // 计算表格中间的数据
            fillBodyRow(row);
            // 完成表格尾部数据的计算及填写
            if ("评定".equals(row.getCell(0).toString())) {
                fillBottomConclusion(sheet, i);
                // 文件的所有计算都已完成，跳出循环，结束剩下2行的读取
                break;
            }
        }
        /*
         * 统计最大值最小值，生成报告表格的时候要用
         */
        row = sheet.getRow(sheet.getPhysicalNumberOfRows()-3);
        row.createCell(15).setCellValue("最大值");
        row.createCell(16).setCellFormula("ROUND("+
                "MAX("+sheet.getRow(6).getCell(9).getReference()+":"+
                sheet.getRow(sheet.getPhysicalNumberOfRows()-4).getCell(9).getReference()+"),1)");//=MAX(J7:J57)
        row.createCell(17).setCellFormula("ROUND("+"MAX("+sheet.getRow(6).getCell(10).getReference()+":"+
                sheet.getRow(sheet.getPhysicalNumberOfRows()-4).getCell(10).getReference()+"),1)");//=MAX(J7:J57)

        row = sheet.getRow(sheet.getPhysicalNumberOfRows()-2);
        row.createCell(15).setCellValue("最小值");
        row.createCell(16).setCellFormula("ROUND("+"MIN("+sheet.getRow(6).getCell(9).getReference()+":"+
                sheet.getRow(sheet.getPhysicalNumberOfRows()-4).getCell(9).getReference()+"),1)");//=MIN(J7:J57)
        row.createCell(17).setCellFormula("ROUND("+"MIN("+sheet.getRow(6).getCell(10).getReference()+":"+
                sheet.getRow(sheet.getPhysicalNumberOfRows()-4).getCell(10).getReference()+"),1)");//=MIN(J7:J57)
    }

    /**
     *
     * @param sheet
     * @param i
     */
    private void fillBottomConclusion(XSSFSheet sheet, int i) {
        XSSFRow row1, row2, row3;
        row1 = sheet.getRow(i);
        row2 = sheet.getRow(i + 1);
        row3 = sheet.getRow(i + 2);

        // 试验室标准密度控制->平均值
        row1.getCell(4).setCellType(Cell.CELL_TYPE_NUMERIC);
        row1.getCell(4).setCellFormula(
                "AVERAGE(" + "J7" + ":" + "J" + row1.getRowNum() + ")");// AVERAGE(J7:J24)
        // 试验室标准密度控制->监测点数
        row1.getCell(6).setCellType(Cell.CELL_TYPE_NUMERIC);
        row1.getCell(6).setCellFormula(
                "COUNT(" + "J7" + ":" + "J" + row1.getRowNum() + ")");// COUNT(J7:J24)

        // 最大理论密度控制->平均值
        row1.getCell(10).setCellType(Cell.CELL_TYPE_NUMERIC);
        row1.getCell(10).setCellFormula(
                "AVERAGE(" + "K7" + ":" + "K" + row1.getRowNum() + ")");// AVERAGE(K7:K24)
        // 最大理论密度控制->监测点数
        row1.getCell(12).setCellType(Cell.CELL_TYPE_NUMERIC);
        row1.getCell(12).setCellFormula(
                "COUNT(" + "K7" + ":" + "K" + row1.getRowNum() + ")");// COUNT(K7:K24)

        // 标准差
        row2.getCell(4).setCellType(Cell.CELL_TYPE_NUMERIC);
        row2.getCell(4).setCellFormula(
                "STDEV(" + "J7" + ":" + "J" + (row2.getRowNum() - 1) + ")");// =STDEV(J7:J24)
        // 标准差
        row2.getCell(10).setCellType(Cell.CELL_TYPE_NUMERIC);
        row2.getCell(10).setCellFormula(
                "STDEV(" + "K7" + ":" + "K" + (row2.getRowNum() - 1) + ")");// =STDEV(K7:K24)

        // 代表值
        row3.getCell(4).setCellType(Cell.CELL_TYPE_NUMERIC);
        row3.getCell(4).setCellFormula(
                "ROUND(" + row1.getCell(4).getReference() + "-("
                        + row2.getCell(4).getReference() + "*VLOOKUP("
                        + row1.getCell(6).getReference()
                        + ",保证率系数!$A:$D,3)),1)");// ROUND(E40-(E41*VLOOKUP(G40,保证率系数!$A:$D,3,)),1)
        row3.getCell(10).setCellType(Cell.CELL_TYPE_NUMERIC);
        row3.getCell(10).setCellFormula(
                "ROUND(" + row1.getCell(10).getReference() + "-("
                        + row2.getCell(10).getReference() + "*VLOOKUP("
                        + row1.getCell(12).getReference()
                        + ",保证率系数!$A:$D,3)),1)");// =ROUND(K40-(K41*VLOOKUP(M40,保证率系数!$A:$D,3,)),1)
        // 右下角结论
        row1.getCell(14).setCellType(Cell.CELL_TYPE_NUMERIC);
        row1.getCell(14).setCellFormula(
                "IF((" + row3.getCell(4).getReference() + ">="
                        + row2.getCell(2).getReference() + ")*AND("
                        + row3.getCell(10).getReference() + ">="
                        + row2.getCell(8).getReference() + "),\"合格\",\"不合格\")");// IF((E42>=C41)*AND(K42>=I41),"合格","不合格")
        // 计算右侧结论
        fillRightConclusion(sheet);
        // 合格点数 =COUNTIF(N7:N24,"√")
        row2.getCell(6).setCellType(Cell.CELL_TYPE_NUMERIC);
        row2.getCell(6).setCellFormula(
                "COUNTIF(" + "N7" + ":" + row1.getCell(13).getReference()
                        + ",\"√\")");
        row2.getCell(12).setCellFormula(
                "COUNTIF(" + "N7" + ":" + row1.getCell(13).getReference()
                        + ",\"√\")");
        // 合格率 =G41/G40*100
        row3.getCell(6).setCellFormula(
                row2.getCell(6).getReference() + "/"
                        + row1.getCell(6).getReference() + "*100");
        row3.getCell(12).setCellFormula(
                row2.getCell(12).getReference() + "/"
                        + row1.getCell(12).getReference() + "*100");
    }

    /**
     *
     * @param sheet
     */
    private void fillRightConclusion(XSSFSheet sheet) {
        XSSFRow row = null;
        for (int i = sheet.getFirstRowNum(); i <= sheet
                .getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            // 判断前面的数据有值时，才进行结论的计算
            if (row.getCell(9).getCellType() == Cell.CELL_TYPE_FORMULA
                    && row.getCell(10).getCellType() == Cell.CELL_TYPE_FORMULA
                    && row.getCell(11).getCellType() == Cell.CELL_TYPE_NUMERIC
                    && row.getCell(12).getCellType() == Cell.CELL_TYPE_NUMERIC) {
                // 计算结论
                row.getCell(13).setCellFormula(
                        "IF(" + getLastCell(sheet).getReference()
                                + "=\"合格\",IF(("
                                + row.getCell(9).getReference() + ">="
                                + row.getCell(11).getReference() + ")*("
                                + row.getCell(10).getReference() + ">"
                                + row.getCell(12).getReference()
                                + "),\"√\",\"\"),\"\")");// =IF($O$40="合格",IF((J7>=L7)*(K7>M7),"√",""),"")
                row.getCell(14).setCellFormula(
                        "IF(" + row.getCell(9).getReference()
                                + "=\"\",\"\",IF("
                                + row.getCell(13).getReference()
                                + "=\"\",\"×\",\"\"))");// =IF(J13="","",IF(N13="","×",""))

            }
        }

    }

    /**
     *
     * @param sheet
     * @return
     */
    private XSSFCell getLastCell(XSSFSheet sheet) {
        XSSFRow row = null;
        XSSFCell cell = null;
        row = sheet.getRow(sheet.getLastRowNum() - 2);
        cell = row.getCell(14);
        return cell;
    }

    /**
     *
     * @param row
     */
    private void fillBodyRow(XSSFRow row) {
        if (row.getCell(3).getCellType() == Cell.CELL_TYPE_NUMERIC
                && row.getCell(4).getCellType() == Cell.CELL_TYPE_NUMERIC
                && row.getCell(5).getCellType() == Cell.CELL_TYPE_NUMERIC) {

            row.getCell(6).setCellType(Cell.CELL_TYPE_NUMERIC);
            row.getCell(6).setCellFormula(
                    row.getCell(3).getReference() + "/" + "("
                            + row.getCell(5).getReference() + "-"
                            + row.getCell(4).getReference() + ")");// D7/(F7-E7)
        }
        if (row.getCell(7).getCellType() == Cell.CELL_TYPE_NUMERIC
                && row.getCell(8).getCellType() == Cell.CELL_TYPE_NUMERIC) {
            // 试验室
            row.getCell(9).setCellType(Cell.CELL_TYPE_NUMERIC);
            row.getCell(9).setCellFormula("ROUND("+
                    row.getCell(6).getReference() + "/"
                    + row.getCell(7).getReference() + "*" + "100,1)");// G7/H7*100
            // 最大理论
            row.getCell(10).setCellType(Cell.CELL_TYPE_NUMERIC);
            row.getCell(10).setCellFormula("ROUND("+
                    row.getCell(6).getReference() + "/"
                    + row.getCell(8).getReference() + "*" + "100,1)");// G7/I7*100

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
     * @param sheet
     */
    private void getTunnelTotal(XSSFSheet sheet){
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
            if(flag && row.getCell(0) != null && !"".equals(row.getCell(0).toString())){
                if(!name.equals(row.getCell(0).toString().replaceAll("[^\u4E00-\u9FA5]", "")) && !"".equals(name)){
                    endrow = sheet.getRow(i-1);
                    startrow.createCell(15).setCellFormula("COUNT("
                            +startrow.getCell(9).getReference()+":"
                            +endrow.getCell(9).getReference()+")");//=COUNT(J7:J21)
                    startrow.createCell(16).setCellFormula("COUNTIF("
                            +startrow.getCell(13).getReference()+":"
                            +endrow.getCell(13).getReference()+",\"√\")");//=COUNTIF(N7:N22,"√")

                    startrow.createCell(17).setCellFormula(startrow.getCell(16).getReference()+"*100/"
                            +startrow.getCell(15).getReference());
                    name = row.getCell(0).toString().replaceAll("[^\u4E00-\u9FA5]", "");
                    startrow = row;
                }
                if("".equals(name)){
                    name = row.getCell(0).toString().replaceAll("[^\u4E00-\u9FA5]", "");
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
     * @param wb
     * @param data
     * @param sheetname
     * @param fbgcname
     */
    private void lqlmysdOther(XSSFWorkbook wb, List<JjgFbgcSdgcSdlqlmysd> data, String sheetname, String fbgcname) {
        if (data.size() > 0 && data != null) {
            createZXZTable(wb, getZXZtableNum(data.size()), sheetname);
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
            XSSFSheet sheet = wb.getSheet(sheetname);
            String type = data.get(0).getLqs();
            String zh = data.get(0).getZh();
            sheet.getRow(1).getCell(2).setCellValue(data.get(0).getProname());
            sheet.getRow(1).getCell(9).setCellValue(data.get(0).getHtd());
            sheet.getRow(2).getCell(2).setCellValue(data.get(0).getFbgc());
            sheet.getRow(2).getCell(9).setCellValue(fbgcname);
            sheet.getRow(3).getCell(2).setCellValue(data.get(0).getLmlx());
            String date = simpleDateFormat.format(data.get(0).getJcsj());
            sheet.getRow(3).getCell(9).setCellValue(date);
            int index = 6;
            for(int i =0; i < data.size(); i++){
                if (zh.equals(data.get(i).getZh())) {
                    for(int j = 0; j < 4; j++){
                        if(i+j < data.size()){
                            /*if(!data.get(i)[5].equals(data.get(i+j)[5])){
                                i+=j-1;
                                break;
                            }
                            else */if("上面层".equals(data.get(i+j).getCw())){
                                sheet.getRow(index).getCell(2).setCellValue("上面层");
                                fillCommonCellData(sheet, index, data.get(i+j));
                            }
                            else if("中面层".equals(data.get(i+j).getCw())){
                                sheet.getRow(index+1).getCell(2).setCellValue("中面层");
                                fillCommonCellData(sheet, index+1, data.get(i+j));
                            }
                            else if("下面层".equals(data.get(i+j).getCw())){
                                sheet.getRow(index+2).getCell(2).setCellValue("下面层");
                                fillCommonCellData(sheet, index+2, data.get(i+j));
                            }
                        }
                        else{
                            i+=j-1;
                            break;
                        }
                    }
                    index+=3;
                    sheet.addMergedRegion(new CellRangeAddress(index-3, index-1, 0, 0));
                    sheet.getRow(index-3).getCell(0).setCellValue(zh);
                    sheet.addMergedRegion(new CellRangeAddress(index-3, index-1, 1, 1));
                    //sheet.getRow(index-3).getCell(1).setCellValue(data.get(i)[8]);
                }
                else{
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
    private void fillCommonCellData(XSSFSheet sheet, int index, JjgFbgcSdgcSdlqlmysd row) {
        sheet.getRow(index).getCell(0).setCellValue(row.getZh());
        sheet.getRow(index).getCell(1).setCellValue(row.getQywz());
        sheet.getRow(index).getCell(3).setCellValue(Double.parseDouble(row.getGzsjzl()));
        sheet.getRow(index).getCell(4).setCellValue(Double.parseDouble(row.getSjszzl()));
        sheet.getRow(index).getCell(5).setCellValue(Double.parseDouble(row.getSjbgzl()));
        sheet.getRow(index).getCell(7).setCellValue(Double.parseDouble(row.getSysbzmd()));
        sheet.getRow(index).getCell(8).setCellValue(Double.parseDouble(row.getZdllmd()));
        sheet.getRow(index).getCell(11).setCellValue(Double.parseDouble(row.getSysbzmdgdz()));
        sheet.getRow(index).getCell(12).setCellValue(Double.parseDouble(row.getZdllmdgdz()));
        sheet.getRow(sheet.getLastRowNum()-1).getCell(2).setCellValue(Double.parseDouble(row.getSysbzmdgdz()));
        sheet.getRow(sheet.getLastRowNum()-1).getCell(8).setCellValue(Double.parseDouble(row.getZdllmdgdz()));
    }

    /**
     *
     * @param wb
     * @param tableNum
     * @param sheetname
     */
    private void createZXZTable(XSSFWorkbook wb, int tableNum, String sheetname) {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            if(i < record-1){
                RowCopy.copyRows(wb, "沥青路面压实度", sheetname, 6, 23, (i - 1) * 18 + 24);
            }
            else{
                RowCopy.copyRows(wb, "沥青路面压实度", sheetname, 6, 20, (i - 1) * 18 + 24);
            }
        }
        if(record == 1){
            wb.getSheet(sheetname).shiftRows(22, 24, -1);
        }
        RowCopy.copyRows(wb, "source", sheetname, 0, 3,(record) * 18 + 3);
        /*wb.getSheet("隧道左幅").getRow((record) * 18 + 4).getCell(2).setCellValue(lab);
        wb.getSheet("隧道左幅").getRow((record) * 18 + 4).getCell(8).setCellValue(max);*/
        wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 14, 0,(record) * 18 + 5);
    }

    /**
     *
     * @param size
     * @return
     */
    private int getZXZtableNum(int size) {
        return size%18 <= 15 ? size/18+1 : size/18+2;
    }


    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        String title = "沥青路面压实度质量鉴定表";
        List<Map<String, Object>> mapList = new ArrayList<>();

        List<Map<String,Object>> sdmclist = jjgFbgcSdgcSdlqlmysdMapper.selectsdmc(proname,htd,fbgc);
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
        DecimalFormat df = new DecimalFormat(".00");
        DecimalFormat decf = new DecimalFormat("0.##");
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "43隧道沥青路面压实度-"+sdmc+".xlsx");
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
                    XSSFCell htdname = slSheet.getRow(1).getCell(9);//合同段名
                    Map map = new HashMap();

                    if (proname.equals(xmname.toString()) && title.equals(bt.toString()) && htd.equals(htdname.toString())) {
                        //获取到最后一行
                        int lastRowNum = slSheet.getLastRowNum();
                        slSheet.getRow(lastRowNum-2).getCell(12).setCellType(CellType.STRING);//总点数
                        slSheet.getRow(lastRowNum-1).getCell(12).setCellType(CellType.STRING);//合格点数
                        slSheet.getRow(lastRowNum).getCell(12).setCellType(CellType.STRING);//合格率
                        double zds = Double.valueOf(slSheet.getRow(lastRowNum-2).getCell(12).getStringCellValue());
                        double hgds = Double.valueOf(slSheet.getRow(lastRowNum-1).getCell(12).getStringCellValue());
                        double hgl = Double.valueOf(slSheet.getRow(lastRowNum).getCell(12).getStringCellValue());
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
    public void exportsdlqlmysd(HttpServletResponse response) {
        String fileName = "06隧道沥青路面压实度实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcSdgcSdlqlmysdVo()).finish();

    }

    @Override
    public void importsdlqlmysd(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcSdgcSdlqlmysdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcSdgcSdlqlmysdVo>(JjgFbgcSdgcSdlqlmysdVo.class) {
                                @Override
                                public void handle(List<JjgFbgcSdgcSdlqlmysdVo> dataList) {
                                    for(JjgFbgcSdgcSdlqlmysdVo sdlqlmysdVo: dataList)
                                    {
                                        JjgFbgcSdgcSdlqlmysd fbgcLmgcLqlmysd = new JjgFbgcSdgcSdlqlmysd();
                                        BeanUtils.copyProperties(sdlqlmysdVo,fbgcLmgcLqlmysd);
                                        fbgcLmgcLqlmysd.setCreatetime(new Date());
                                        fbgcLmgcLqlmysd.setProname(commonInfoVo.getProname());
                                        fbgcLmgcLqlmysd.setHtd(commonInfoVo.getHtd());
                                        fbgcLmgcLqlmysd.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcSdgcSdlqlmysdMapper.insert(fbgcLmgcLqlmysd);
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
        List<Map<String,Object>> sdmclist = jjgFbgcSdgcSdlqlmysdMapper.selectsdmc(proname,htd,fbgc);
        return sdmclist;
    }
}
