package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabx;
import glgc.jjgys.model.projectvo.jagc.JjgFbgcJtaqssJabxVo;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcJtaqssJabxMapper;
import glgc.jjgys.system.service.JjgFbgcJtaqssJabxService;
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
 *  交安标线服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-03-01
 */
@Service
public class JjgFbgcJtaqssJabxServiceImpl extends ServiceImpl<JjgFbgcJtaqssJabxMapper, JjgFbgcJtaqssJabx> implements JjgFbgcJtaqssJabxService {

    @Autowired
    private JjgFbgcJtaqssJabxMapper jjgFbgcJtaqssJabxMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;


    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws Exception {
        XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获取数据
        QueryWrapper<JjgFbgcJtaqssJabx> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByDesc("wz","hdscz1");
        List<JjgFbgcJtaqssJabx> data = jjgFbgcJtaqssJabxMapper.selectList(wrapper);
        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"57交安标线厚度.xlsx");
        if (data == null || data.size()==0){
            return;
        }else {
            //存放鉴定表的目录
            File fdir = new File(filepath + File.separator + proname + File.separator + htd);
            if (!fdir.exists()) {
                //创建文件根目录
                fdir.mkdirs();
            }
            File directory = new File("service-system/src/main/resources");
            String reportPath = directory.getCanonicalPath();
            String path = reportPath + "/static/交安标线厚度新.xlsx";
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            createTable(gettableNum(data.size()),wb);
            if(DBtoExcel(data,wb,proname,htd,fbgc)){
                for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                    calculateOneSheet(wb.getSheetAt(j));
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
            //查询是否有白线逆反射系数
            List<JjgFbgcJtaqssJabx> bxnfsxsdata = jjgFbgcJtaqssJabxMapper.selectbxnfsxs(proname,htd,fbgc);
            List<JjgFbgcJtaqssJabx> hxnfsxsdata = jjgFbgcJtaqssJabxMapper.selecthxnfsxs(proname,htd,fbgc);
            if (bxnfsxsdata !=null || bxnfsxsdata.size()>0){
                bxnfsxs(data);
            }
            if (hxnfsxsdata !=null || hxnfsxsdata.size()>0){
                hxnfsxs(data);
            }


        }

    }

    /**
     * 交安黄线逆反射系数
     * @param data
     * @throws IOException
     * @throws ParseException
     */
    public void hxnfsxs(List<JjgFbgcJtaqssJabx> data) throws IOException, ParseException {
        String proname = data.get(0).getProname();
        String htd = data.get(0).getHtd();
        String fbgc = data.get(0).getFbgc();
        XSSFWorkbook xwb = null;
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "57交安标线黄线逆反射系数.xlsx");
        if (data == null || data.size() == 0) {
            return;
        } else {
            //存放鉴定表的目录
            File fdir = new File(filepath + File.separator + proname + File.separator + htd);
            if (!fdir.exists()) {
                //创建文件根目录
                fdir.mkdirs();
            }
            File directory = new File("src/main/resources");
            String reportPath = directory.getCanonicalPath();
            String path = reportPath + "/static/交安标线逆反射系数新.xlsx";
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            xwb = new XSSFWorkbook(out);
            createTable(gettableNum(data.size()), xwb);
            if(DBtohxnfsxsExcel(data,xwb,proname,htd,fbgc)){
                for (int j = 0; j < xwb.getNumberOfSheets(); j++) {
                    calculatebxnfsxs(xwb.getSheetAt(j));

                }
                for (int j = 0; j < xwb.getNumberOfSheets(); j++) {   //表内公式  计算 显示结果
                    JjgFbgcCommonUtils.updateFormula(xwb, xwb.getSheetAt(j));
                }

                FileOutputStream fileOut = new FileOutputStream(f);
                xwb.write(fileOut);
                fileOut.flush();
                fileOut.close();

            }
            out.close();
            xwb.close();
        }
    }

    /**
     * 黄线
     * @param data
     * @param xwb
     * @param proname
     * @param htd
     * @param fbgc
     * @return
     * @throws ParseException
     */
    public boolean DBtohxnfsxsExcel(List<JjgFbgcJtaqssJabx> data,XSSFWorkbook xwb,String proname,String htd,String fbgc) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        int z=1;
        String type="";
        //规定值或允许偏差
        HashMap<String, String> error=new HashMap<>();
        for(JjgFbgcJtaqssJabx row:data){
            if(!type.contains(row.getBxlx())) {
                type+=row.getBxlx()+",";
                error.put(row.getBxlx(), row.getBxnfsxsgdz());
            }else if(!error.get(row.getBxlx()).contains(row.getBxnfsxsgdz())) {
                error.put(row.getBxlx(), error.get(type)+","+row.getBxnfsxsgdz());
            }
            z++;
        }
        type=type.substring(0, type.length()-1);
        String errors="";
        for(String s:error.keySet()){
            errors+=(s+"≥"+error.get(s))+";";
        }
        errors=errors.substring(0, errors.length()-1);

        //填写表头
        XSSFSheet sheet = xwb.getSheet("标线");
        sheet.getRow(1).getCell(4).setCellValue(proname);
        sheet.getRow(1).getCell(14).setCellValue(htd);
        sheet.getRow(2).getCell(14).setCellValue(fbgc);
        sheet.getRow(3).getCell(4).setCellValue(type);//标线类型
        sheet.getRow(5).getCell(4).setCellValue(errors);//标线允许误差
        String date = simpleDateFormat.format(data.get(0).getJcsj());
        for(int i =1; i < data.size(); i++){
            date = JjgFbgcCommonUtils.getLastDate(date, simpleDateFormat.format(data.get(i).getJcsj()));
        }
        sheet.getRow(3).getCell(14).setCellValue(date);//检测日期

        int index = 0;
        int tableNum = 0;
        for(int i =0; i < data.size(); i++){
            if(i%5 == 0){
                sheet.addMergedRegion(new CellRangeAddress(index+7, index+11, 0, 0));
                sheet.getRow(index+7).getCell(0).setCellValue(data.get(i).getWz());  //位置
            }
            sheet.getRow(index+7).getCell(1).setCellValue(Double.valueOf((i%5+1)));//序号
            fillhxCommonCellData(sheet, tableNum, index+7, data.get(i));
            index++;

        }
        return true;
    }

    /**
     * 黄线
     * @param sheet
     * @param tableNum
     * @param index
     * @param row
     */
    public void fillhxCommonCellData(XSSFSheet sheet, int tableNum, int index,JjgFbgcJtaqssJabx row) {
        sheet.getRow(index).getCell(0).setCellValue(row.getWz());  //位置
        sheet.getRow(index).getCell(2).setCellValue(row.getBxlx());
        sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(row.getHxnfsxsgdz()));
        if(row.getHxscz1() != null && !"".equals(row.getHxscz1())){//第1条线
            sheet.getRow(tableNum*37+index).getCell(4).setCellValue(Double.valueOf(row.getHxscz1()));
        }else{
            sheet.getRow(tableNum*37+index).getCell(4).setCellValue("-");
        }

        if(row.getHxscz2() != null && !"".equals(row.getHxscz2())){//第2
            sheet.getRow(tableNum*37+index).getCell(7).setCellValue(Double.valueOf(row.getHxscz2()));
        }else{
            sheet.getRow(tableNum*37+index).getCell(7).setCellValue("-");
        }

        if(row.getHxscz3() != null && !"".equals(row.getHxscz3())){//第三
            sheet.getRow(tableNum*37+index).getCell(10).setCellValue(Double.valueOf(row.getHxscz3()));
        }else{
            sheet.getRow(tableNum*37+index).getCell(10).setCellValue("-");
        }
        if(row.getHxscz4() != null && !"".equals(row.getHxscz4())){//第4
            sheet.getRow(tableNum*37+index).getCell(13).setCellValue(Double.valueOf(row.getHxscz4()));
        }else{
            sheet.getRow(tableNum*37+index).getCell(13).setCellValue("-");
        }
        if(row.getHxscz5() != null && !"".equals(row.getHxscz5().toString())){//5
            sheet.getRow(tableNum*37+index).getCell(16).setCellValue(Double.valueOf(row.getHxscz5()));
        }else{
            sheet.getRow(tableNum*37+index).getCell(16).setCellValue("-");
        }

    }

    /**
     * 交安白线逆反射系数
     * @param data
     */
    public void bxnfsxs(List<JjgFbgcJtaqssJabx> data) throws IOException, ParseException {
        String proname = data.get(0).getProname();
        String htd = data.get(0).getHtd();
        String fbgc = data.get(0).getFbgc();
        XSSFWorkbook xwb = null;
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"57交安标线白线逆反射系数.xlsx");
        if (data == null || data.size()==0){
            return;
        }else {
            //存放鉴定表的目录
            File fdir = new File(filepath + File.separator + proname + File.separator + htd);
            if (!fdir.exists()) {
                //创建文件根目录
                fdir.mkdirs();
            }
            File directory = new File("src/main/resources");
            String reportPath = directory.getCanonicalPath();
            String path = reportPath + "/static/交安标线逆反射系数新.xlsx";
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            xwb = new XSSFWorkbook(out);
            createTable(gettableNum(data.size()),xwb);
            //createTable(13,xwb);
            if(DBtobxnfsxsExcel(data,xwb,proname,htd,fbgc)){

                for (int j = 0; j < xwb.getNumberOfSheets(); j++) {
                    calculatebxnfsxs(xwb.getSheetAt(j));

                }
                for (int j = 0; j < xwb.getNumberOfSheets(); j++) {   //表内公式  计算 显示结果
                    JjgFbgcCommonUtils.updateFormula(xwb, xwb.getSheetAt(j));
                }

                FileOutputStream fileOut = new FileOutputStream(f);
                xwb.write(fileOut);
                fileOut.flush();
                fileOut.close();

            }
            out.close();
            xwb.close();

        }

    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        String title = "道路交通标线施工质量鉴定表";
        String sheetname = "标线";

        DecimalFormat df = new DecimalFormat(".00");
        DecimalFormat decf = new DecimalFormat("0.##");

        String fileName1 = "57交安标线厚度";
        String fileName2 = "57交安标线白线逆反射系数";
        String fileName3 = "57交安标线黄线逆反射系数";
        List list = new ArrayList();
        list.add(fileName1);
        list.add(fileName2);
        list.add(fileName3);

        List<Map<String,Object>> mapList = new ArrayList<>();

        for (int i=0;i<list.size();i++) {
            //获取鉴定表文件
            File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+list.get(i)+".xlsx");
            if(!f.exists()){
                return null;
            }else {
                XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));
                //读取工作表
                XSSFSheet slSheet = xwb.getSheet(sheetname);
                XSSFCell bt = slSheet.getRow(0).getCell(0);
                XSSFCell xmname = slSheet.getRow(1).getCell(4);//陕西高速
                XSSFCell htdname = slSheet.getRow(1).getCell(14);//LJ-1
                XSSFCell hd = slSheet.getRow(2).getCell(14);//涵洞
                Map<String,Object> jgmap = new HashMap<>();
                if(proname.equals(xmname.toString()) && title.equals(bt.toString()) && htd.equals(htdname.toString()) && fbgc.equals(hd.toString())) {
                    //获取到最后一行
                    int lastRowNum = slSheet.getLastRowNum();
                    slSheet.getRow(lastRowNum - 4).getCell(14).setCellType(CellType.STRING);
                    slSheet.getRow(lastRowNum - 3).getCell(14).setCellType(CellType.STRING);
                    slSheet.getRow(lastRowNum - 2).getCell(14).setCellType(CellType.STRING);
                    slSheet.getRow(lastRowNum - 1).getCell(14).setCellType(CellType.STRING);
                    double zds = Double.valueOf(slSheet.getRow(lastRowNum - 4).getCell(14).getStringCellValue());
                    double hgds = Double.valueOf(slSheet.getRow(lastRowNum - 3).getCell(14).getStringCellValue());
                    double bhgds = Double.valueOf(slSheet.getRow(lastRowNum - 2).getCell(14).getStringCellValue());
                    double hgl = Double.valueOf(slSheet.getRow(lastRowNum - 1).getCell(14).getStringCellValue());
                    String zdsz = decf.format(zds);
                    String hgdsz = decf.format(hgds);
                    String bhgdsz = decf.format(bhgds);
                    String hglz = df.format(hgl);
                    jgmap.put("检测项目",list.get(i).toString().substring(2));
                    jgmap.put("总点数", zdsz);
                    jgmap.put("合格点数", hgdsz);
                    jgmap.put("不合格点数", bhgdsz);
                    jgmap.put("合格率", hglz);
                    mapList.add(jgmap);
                }

            }
        }
        return mapList;

        //获取鉴定表文件
        /*File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"57交安标线厚度.xlsx");
        if(!f.exists()){
            return null;
        }else {
            XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));
            //读取工作表
            XSSFSheet slSheet = xwb.getSheet(sheetname);
            XSSFCell bt = slSheet.getRow(0).getCell(0);
            XSSFCell xmname = slSheet.getRow(1).getCell(4);//陕西高速
            XSSFCell htdname = slSheet.getRow(1).getCell(14);//LJ-1
            XSSFCell hd = slSheet.getRow(2).getCell(14);//涵洞
            List<Map<String,Object>> mapList = new ArrayList<>();
            Map<String,Object> jgmap = new HashMap<>();
            if(proname.equals(xmname.toString()) && title.equals(bt.toString()) && htd.equals(htdname.toString()) && fbgc.equals(hd.toString())){
                //获取到最后一行
                int lastRowNum = slSheet.getLastRowNum();
                slSheet.getRow(lastRowNum-4).getCell(14).setCellType(CellType.STRING);
                slSheet.getRow(lastRowNum-3).getCell(14).setCellType(CellType.STRING);
                slSheet.getRow(lastRowNum-2).getCell(14).setCellType(CellType.STRING);
                slSheet.getRow(lastRowNum-1).getCell(14).setCellType(CellType.STRING);
                double zds= Double.valueOf(slSheet.getRow(lastRowNum-4).getCell(14).getStringCellValue());
                double hgds= Double.valueOf(slSheet.getRow(lastRowNum-3).getCell(14).getStringCellValue());
                double bhgds= Double.valueOf(slSheet.getRow(lastRowNum-2).getCell(14).getStringCellValue());
                double hgl = Double.valueOf(slSheet.getRow(lastRowNum-1).getCell(14).getStringCellValue());
                String zdsz = decf.format(zds);
                String hgdsz = decf.format(hgds);
                String bhgdsz = decf.format(bhgds);
                String hglz = df.format(hgl);
                jgmap.put("总点数",zdsz);
                jgmap.put("合格点数",hgdsz);
                jgmap.put("不合格点数",bhgdsz);
                jgmap.put("合格率",hglz);
                mapList.add(jgmap);
                return mapList;
            }
            return null;
        }*/
    }

    /**
     * 计算白线逆反射系数
     * @param sheet
     */
    public void calculatebxnfsxs(XSSFSheet sheet) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        boolean flag = false;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if (flag) {
                if (row.getCell(3).getCellType() != Cell.CELL_TYPE_NUMERIC
                        && !"合计".equals(row.getCell(0).toString())) {
                    continue;
                }
                /*
                 * 当数据计算完成时，将数据汇总，计算合格率
                 */
                if ("合计".equals(row.getCell(0).toString())) {
                    calculateTotal(sheet, i, rowstart, rowend);

                    break;
                }
                /*
                 * 判断当前所需数据存在，然后进行单元格计算
                 */
                if (row.getCell(4).getCellType() == Cell.CELL_TYPE_NUMERIC ) {//值大于等于规定值既为合格

                    row.getCell(5).setCellFormula(
                            "IF(" + row.getCell(4).getReference() + ">="
                                    +row.getCell(3).getReference()
                                    +",\"√\",\"\")");					// =IF(实测值>=规定值,"√","")

                    row.getCell(6).setCellFormula(
                            "IF(" + row.getCell(5).getReference() + "=\"\"" + ",\"×\",\"\")");

                }

                if(row.getCell(7).getCellType() == Cell.CELL_TYPE_NUMERIC ) { //
                    row.getCell(8).setCellFormula(
                            "IF(" + row.getCell(5).getReference() + ">="
                                    +row.getCell(3).getReference()
                                    +",\"√\",\"\")");					// =IF(实测值>=规定值,"√","")

                    row.getCell(9).setCellFormula(
                            "IF(" + row.getCell(8).getReference() + "=\"\"" + ",\"×\",\"\")");

                }

                if(row.getCell(10).getCellType() == Cell.CELL_TYPE_NUMERIC ) { //
                    row.getCell(11).setCellFormula(
                            "IF(" + row.getCell(10).getReference() + ">="
                                    +row.getCell(3).getReference()
                                    +",\"√\",\"\")");

                    row.getCell(12).setCellFormula(
                            "IF(" + row.getCell(11).getReference() + "=\"\"" + ",\"×\",\"\")");

                }

                if(row.getCell(13).getCellType() == Cell.CELL_TYPE_NUMERIC) { //
                    row.getCell(14).setCellFormula(
                            "IF(" + row.getCell(13).getReference() + ">="
                                    +row.getCell(3).getReference()
                                    +",\"√\",\"\")");

                    row.getCell(15).setCellFormula(
                            "IF(" + row.getCell(14).getReference() + "=\"\"" + ",\"×\",\"\")");

                }

                if(row.getCell(16).getCellType() == Cell.CELL_TYPE_NUMERIC) { //
                    row.getCell(17).setCellFormula(
                            "IF(" + row.getCell(16).getReference() + ">="
                                    +row.getCell(3).getReference()
                                    +",\"√\",\"\")");

                    row.getCell(18).setCellFormula(
                            "IF(" + row.getCell(17).getReference() + "=\"\"" + ",\"×\",\"\")");

                }
                rowend = row;
            }
            if ("位置".equals(row.getCell(0).toString())) {
                rowstart = sheet.getRow(i + 1);
                flag = true;
            }

        }
    }

    /**
     * 将白线逆反射系数据写入
     * @param data
     * @param xwb
     * @param proname
     * @param htd
     * @param fbgc
     * @return
     * @throws ParseException
     */
    public boolean DBtobxnfsxsExcel(List<JjgFbgcJtaqssJabx> data,XSSFWorkbook xwb,String proname,String htd,String fbgc) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        int z=1;
        String type="";
        //规定值或允许偏差
        HashMap<String, String> error=new HashMap<>();
        for(JjgFbgcJtaqssJabx row:data){
            if(!type.contains(row.getBxlx())) {
                type+=row.getBxlx()+",";
                error.put(row.getBxlx(), row.getBxnfsxsgdz());
            }else if(!error.get(row.getBxlx()).contains(row.getBxnfsxsgdz())) {
                error.put(row.getBxlx(), error.get(type)+","+row.getBxnfsxsgdz());
            }
            z++;
        }
        type=type.substring(0, type.length()-1);
        String errors="";
        for(String s:error.keySet()){
            errors+=(s+"≥"+error.get(s))+";";
        }
        errors=errors.substring(0, errors.length()-1);

        //填写表头
        XSSFSheet sheet = xwb.getSheet("标线");
        sheet.getRow(1).getCell(4).setCellValue(proname);
        sheet.getRow(1).getCell(14).setCellValue(htd);
        sheet.getRow(2).getCell(14).setCellValue(fbgc);
        sheet.getRow(3).getCell(4).setCellValue(type);//标线类型
        sheet.getRow(5).getCell(4).setCellValue(errors);//标线允许误差
        String date = simpleDateFormat.format(data.get(0).getJcsj());
        for(int i =1; i < data.size(); i++){
            date = JjgFbgcCommonUtils.getLastDate(date, simpleDateFormat.format(data.get(i).getJcsj()));
        }
        sheet.getRow(3).getCell(14).setCellValue(date);//检测日期

        int index = 0;
        int tableNum = 0;
        for(int i =0; i < data.size(); i++){
            if(i%5 == 0){
                sheet.addMergedRegion(new CellRangeAddress(index+7, index+11, 0, 0));
                sheet.getRow(index+7).getCell(0).setCellValue(data.get(i).getWz());  //位置
            }
            sheet.getRow(index+7).getCell(1).setCellValue(Double.valueOf((i%5+1)));//序号
            fillCommonCellData(sheet, tableNum, index+7, data.get(i));
            index++;

        }
        return true;
    }

    /**
     * 填写白线数据
     * @param sheet
     * @param tableNum
     * @param index
     * @param row
     */
    public void fillCommonCellData(XSSFSheet sheet, int tableNum, int index,JjgFbgcJtaqssJabx row) {
        sheet.getRow(index).getCell(0).setCellValue(row.getWz());  //位置
        sheet.getRow(index).getCell(2).setCellValue(row.getBxlx());
        sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(row.getBxnfsxsgdz()));
        if(row.getBxscz1() != null && !"".equals(row.getBxscz1())){//第1条线

            sheet.getRow(tableNum*37+index).getCell(4).setCellValue(Double.valueOf(row.getBxscz1()));
        }else{
            sheet.getRow(tableNum*37+index).getCell(4).setCellValue("-");
        }

        if(row.getBxscz2() != null && !"".equals(row.getBxscz2())){//第2
            sheet.getRow(tableNum*37+index).getCell(7).setCellValue(Double.valueOf(row.getBxscz2()));
        }else{
            sheet.getRow(tableNum*37+index).getCell(7).setCellValue("-");
        }

        if(row.getBxscz3() != null && !"".equals(row.getBxscz3())){//第三
            sheet.getRow(tableNum*37+index).getCell(10).setCellValue(Double.valueOf(row.getBxscz3()));
        }else{
            sheet.getRow(tableNum*37+index).getCell(10).setCellValue("-");
        }
        if(row.getBxscz4() != null && !"".equals(row.getBxscz4())){//第4
            sheet.getRow(tableNum*37+index).getCell(13).setCellValue(Double.valueOf(row.getBxscz4()));
        }else{
            sheet.getRow(tableNum*37+index).getCell(13).setCellValue("-");
        }
        if(row.getBxscz5() != null && !"".equals(row.getBxscz5().toString())){//5
            sheet.getRow(tableNum*37+index).getCell(16).setCellValue(Double.valueOf(row.getBxscz5()));
        }else{
            sheet.getRow(tableNum*37+index).getCell(16).setCellValue("-");
        }

    }

    /**
     * 填写标题
     * @param sheet
     * @param tableNum
     * @param proname
     * @param htd
     * @param fbgc
     */
    public void fillTitleCellData(XSSFSheet sheet, int tableNum,String proname,String htd,String fbgc) {
        sheet.getRow(tableNum*37+1).getCell(4).setCellValue(proname);
        sheet.getRow(tableNum*37+1).getCell(14).setCellValue(htd);
        sheet.getRow(tableNum*37+2).getCell(14).setCellValue(fbgc);
    }

    /**
     * 计算合计
     * @param sheet
     * @param i
     * @param rowstart
     * @param rowend
     */
    public void calculateTotal(XSSFSheet sheet, int i, XSSFRow rowstart, XSSFRow rowend) {
        //总点数
        sheet.getRow(i).getCell(14).setCellFormula(
                "COUNT(" + rowstart.getCell(4).getReference() + ":"
                        + rowend.getCell(4).getReference() + ")+COUNT("
                        + rowstart.getCell(7).getReference() + ":"
                        + rowend.getCell(7).getReference() + ")+COUNT("
                        + rowstart.getCell(10).getReference() + ":"
                        + rowend.getCell(10).getReference() + ")+COUNT("
                        + rowstart.getCell(13).getReference() + ":"
                        + rowend.getCell(13).getReference() + ")+COUNT("
                        + rowstart.getCell(16).getReference() + ":"
                        + rowend.getCell(16).getReference() + ")");

        //合格点数
        sheet.getRow(i+1).getCell(14).setCellFormula(
                "COUNTIF(" + rowstart.getCell(5).getReference() + ":"
                        + rowend.getCell(5).getReference()
                        + ",\"√\")+COUNTIF("
                        + rowstart.getCell(8).getReference() + ":"
                        + rowend.getCell(8).getReference() + ",\"√\")+COUNTIF("
                        + rowstart.getCell(11).getReference() + ":"
                        + rowend.getCell(11).getReference() + ",\"√\")+COUNTIF("
                        + rowstart.getCell(14).getReference() + ":"
                        + rowend.getCell(14).getReference() + ",\"√\")+COUNTIF("
                        + rowstart.getCell(17).getReference() + ":"
                        + rowend.getCell(17).getReference() + ",\"√\")"
        );

        //不合格点数
        sheet.getRow(i+2).getCell(14).setCellFormula(
                "COUNTIF(" + rowstart.getCell(6).getReference() + ":"
                        + rowend.getCell(6).getReference()
                        + ",\"×\")+COUNTIF("
                        + rowstart.getCell(9).getReference() + ":"
                        + rowend.getCell(9).getReference() + ",\"×\")+COUNTIF("
                        + rowstart.getCell(12).getReference() + ":"
                        + rowend.getCell(12).getReference() + ",\"×\")+COUNTIF("
                        + rowstart.getCell(15).getReference() + ":"
                        + rowend.getCell(15).getReference() + ",\"×\")+COUNTIF("
                        + rowstart.getCell(18).getReference() + ":"
                        + rowend.getCell(18).getReference() + ",\"×\")"
        );
        //合格率
        sheet.getRow(i+3).getCell(14).setCellFormula(
                sheet.getRow(i + 1).getCell(14).getReference() + "/"
                        + sheet.getRow(i).getCell(14).getReference()
                        + "*100"
        );
    }

    /**
     * 计算
     * @param sheet
     */
    public void calculateOneSheet(XSSFSheet sheet) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        boolean flag = false;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if (flag) {
                if (row.getCell(3).getCellType() != Cell.CELL_TYPE_NUMERIC && !"合计".equals(row.getCell(0).toString())) {
                    continue;
                }
                /*
                 * 当数据计算完成时，将数据汇总，计算合格率
                 */
                if ("合计".equals(row.getCell(0).toString())) {
                    calculateTotal(sheet, i, rowstart, rowend);
                    break;
                }
                /*
                 * 判断当前所需数据存在，然后进行单元格计算
                 */
                //F
                if (row.getCell(4).getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    row.getCell(5).setCellFormula(
                            "IF(AND(" + row.getCell(4).getReference() + "<="
                                    +row.getCell(19).getReference()+"+"
                                    +row.getCell(20).getReference()+","
                                    + row.getCell(4).getReference()
                                    + ">="
                                    +row.getCell(19).getReference()+"-"
                                    +row.getCell(21).getReference()+"),\"√\",\"\")");// =IF(AND(单元格<=最大允许值,单元格>=最小允许值),"√","")
                    row.getCell(6).setCellFormula(
                            "IF(" + row.getCell(5).getReference() + "=\"\"" + ",\"×\",\"\")");
                }
                //I
                if(row.getCell(7).getCellType() == Cell.CELL_TYPE_NUMERIC) { //厚度实测值2是否有值
                    row.getCell(8).setCellFormula(
                            "IF(AND(" + row.getCell(7).getReference() + "<="
                                    +row.getCell(19).getReference()+"+"
                                    +row.getCell(20).getReference()+","
                                    + row.getCell(7).getReference()
                                    + ">="
                                    +row.getCell(19).getReference()+"-"
                                    +row.getCell(21).getReference()+"),\"√\",\"\")");// =IF(AND(单元格<=最大允许值,单元格>=最小允许值),"√","")
                    row.getCell(9).setCellFormula(
                            "IF(" + row.getCell(8).getReference() + "=\"\"" + ",\"×\",\"\")");
                }
                //L
                if(row.getCell(10).getCellType() == Cell.CELL_TYPE_NUMERIC) { //厚度实测值3是否有值
                    row.getCell(11).setCellFormula(
                            "IF(AND(" + row.getCell(10).getReference() + "<="
                                    +row.getCell(19).getReference()+"+"
                                    +row.getCell(20).getReference()+","
                                    + row.getCell(10).getReference()
                                    + ">="
                                    +row.getCell(19).getReference()+"-"
                                    +row.getCell(21).getReference()+"),\"√\",\"\")");// =IF(AND(单元格<=最大允许值,单元格>=最小允许值),"√","")

                    row.getCell(12).setCellFormula(
                            "IF(" + row.getCell(11).getReference() + "=\"\"" + ",\"×\",\"\")");
                }

                //O
                if(row.getCell(13).getCellType() == Cell.CELL_TYPE_NUMERIC) { //厚度实测值4是否有值
                    row.getCell(14).setCellFormula(
                            "IF(AND(" + row.getCell(13).getReference() + "<="
                                    +row.getCell(19).getReference()+"+"
                                    +row.getCell(20).getReference()+","
                                    + row.getCell(13).getReference()
                                    + ">="
                                    +row.getCell(19).getReference()+"-"
                                    +row.getCell(21).getReference()+"),\"√\",\"\")");// =IF(AND(单元格<=最大允许值,单元格>=最小允许值),"√","")

                    row.getCell(15).setCellFormula(
                            "IF(" + row.getCell(14).getReference() + "=\"\"" + ",\"×\",\"\")");

                }
                //R
                if(row.getCell(16).getCellType() == Cell.CELL_TYPE_NUMERIC) { //厚度实测值5是否有值
                    row.getCell(17).setCellFormula(
                            "IF(AND(" + row.getCell(16).getReference() + "<="
                                    +row.getCell(19).getReference()+"+"
                                    +row.getCell(20).getReference()+","
                                    + row.getCell(16).getReference()
                                    + ">="
                                    +row.getCell(19).getReference()+"-"
                                    +row.getCell(21).getReference()+"),\"√\",\"\")");// =IF(AND(单元格<=最大允许值,单元格>=最小允许值),"√","")

                    row.getCell(18).setCellFormula(
                            "IF(" + row.getCell(17).getReference() + "=\"\"" + ",\"×\",\"\")");
                }
                rowend = row;
            }
            if ("位置".equals(row.getCell(0).toString())) {
                rowstart = sheet.getRow(i + 1);
                flag = true;
            }

        }
    }

    /**
     * 写入数据
     * @param data
     * @param wb
     * @param proname
     * @param htd
     * @param fbgc
     * @return
     * @throws Exception
     */
    public boolean DBtoExcel(List<JjgFbgcJtaqssJabx> data,XSSFWorkbook wb,String proname,String htd,String fbgc) throws Exception {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        XSSFSheet sheet = wb.getSheet("标线");
        sheet.getRow(1).getCell(4).setCellValue(proname);
        sheet.getRow(1).getCell(14).setCellValue(htd);
        sheet.getRow(2).getCell(14).setCellValue(fbgc);
        sheet.getRow(3).getCell(4).setCellValue(data.get(0).getBxlx());//标线类型

        String date = simpleDateFormat.format(data.get(0).getJcsj());
        for(int i =1; i < data.size(); i++){
            Date jcsj = data.get(i).getJcsj();
            date = JjgFbgcCommonUtils.getLastDate(date, simpleDateFormat.format(jcsj));
        }
        sheet.getRow(3).getCell(14).setCellValue(date);//检测日期
        sheet.getRow(5).getCell(4).setCellValue(getError(data));
        int index = 0;

        for(int i =0; i < data.size(); i++){
            if(i%5 == 0){
                sheet.addMergedRegion(new CellRangeAddress(index+7, index+11, 0, 0));
                sheet.getRow(index+7).getCell(0).setCellValue(data.get(i).getWz());
            }
            sheet.getRow(index+7).getCell(1).setCellValue(Double.valueOf((i%5+1)));//序号
            sheet.getRow(index+7).getCell(2).setCellValue(data.get(i).getBxlx());
            sheet.getRow(index+7).getCell(3).setCellValue(Double.valueOf(data.get(i).getHdgdz()));
            if(!" ".equals(data.get(i).getHdscz1()) && data.get(i).getHdscz1() !=null){//厚度实测值1
                sheet.getRow(index+7).getCell(4).setCellValue(Double.valueOf(data.get(i).getHdscz1()));
                sheet.getRow(index+7).createCell(19).setCellValue(data.get(i).getHdgdz());//厚度规定值
                if(!"".equals(data.get(i).getHdyxpsz())){
                    sheet.getRow(index+7).createCell(20).setCellValue(Double.valueOf(data.get(i).getHdyxpsz()));//厚度允许偏差+
                }
                if(!"".equals(data.get(i).getHdyxpsf())){
                    sheet.getRow(index+7).createCell(21).setCellValue(Double.valueOf(data.get(i).getHdyxpsf()));//厚度允许偏差-
                }

            }else{
                sheet.getRow(index+7).getCell(4).setCellValue("-");
            }

            if(data.get(i).getHdscz2() != null && !"".equals(data.get(i).getHdscz2())){//厚度实测值2
                sheet.getRow(index+7).getCell(7).setCellValue(Double.valueOf(data.get(i).getHdscz2()));
            }else{
                sheet.getRow(index+7).getCell(7).setCellValue("-");
            }

            if(data.get(i).getHdscz3() != null && !"".equals(data.get(i).getHdscz3())){//厚度实测值3
                sheet.getRow(index+7).getCell(10).setCellValue(Double.valueOf(data.get(i).getHdscz3()));
            }else{
                sheet.getRow(index+7).getCell(10).setCellValue("-");
            }

            if(data.get(i).getHdscz4() != null && !"".equals(data.get(i).getHdscz4())){//厚度实测值4
                sheet.getRow(index+7).getCell(13).setCellValue(Double.valueOf(data.get(i).getHdscz4()));
            }else{
                sheet.getRow(index+7).getCell(13).setCellValue("-");
            }

            if(data.get(i).getHdscz5() != null && !"".equals(data.get(i).getHdscz5())){//厚度实测值5
                sheet.getRow(index+7).getCell(16).setCellValue(Double.valueOf(data.get(i).getHdscz5()));
            }else{
                sheet.getRow(index+7).getCell(16).setCellValue("-");
            }
            index ++;
        }
        return true;
    }

    /**
     *
     * @param data
     * @return
     */
    public String getError(List<JjgFbgcJtaqssJabx> data){
        String error = "";
        String temp = "";
        ArrayList<String> errorlist = new ArrayList<String>();
        for(int i =0; i < data.size(); i++){
            if(!errorlist.contains(data.get(i).getWz())){
                errorlist.add((data.get(i).getWz()));
                if(!data.get(i).getWz().contains("振动标线")){
                    temp+="白、黄线"+data.get(i).getHdgdz();
                    if(!data.get(i).getHdyxpsz().equals("0") && !data.get(i).getHdyxpsf().equals("0")){
                        if(data.get(i).getHdyxpsz().equals(data.get(i).getHdyxpsf())){
                            temp+="±"+data.get(i).getHdyxpsz();
                        }
                        else{
                            temp+="+"+data.get(i).getHdyxpsz();
                            temp+=",-"+data.get(i).getHdyxpsf();
                        }
                    }else{
                        if(!data.get(i).getHdyxpsz().equals("0")){
                            temp+="+"+data.get(i).getHdyxpsz();
                        }
                        if(!data.get(i).getHdyxpsf().equals("0")){
                            temp+="-"+data.get(i).getHdyxpsf();
                        }
                    }

                }else if(data.get(i).getWz().contains("振动标线")){
                    temp+="振动标线"+ data.get(i).getHdgdz();
                    if(!data.get(i).getHdyxpsz().equals("0") && !data.get(i).getHdyxpsf().equals("0")){
                        if(data.get(i).getHdyxpsz().equals(data.get(i).getHdyxpsf())){
                            temp+="±"+data.get(i).getHdyxpsz();
                        }
                    }else{
                        if(!data.get(i).getHdyxpsz().equals("0")){
                            temp+="+"+data.get(i).getHdyxpsz();
                        }
                        if(!data.get(i).getHdyxpsf().equals("0")){
                            temp+="-"+data.get(i).getHdyxpsf();
                        }
                    }
                }
                if(!error.contains(temp)){
                    error += temp;
                }temp = "";
            }
        }
        return error;
    }

    /**
     * 根据数据获取页数
     * @param size
     * @return
     */
    public int gettableNum(int size){
        return size%30==0?size/30:size/30+1;
    }

    /**
     * 创建页
     * @param tableNum
     * @param wb
     * @throws IOException
     */
    public void createTable(int tableNum,XSSFWorkbook wb) throws IOException {
        int record = 0;
        record = tableNum;
        for (int i = 1; i <record; i++) {
            if(i < record-1){
                RowCopy.copyRows(wb, "标线", "标线", 7, 36, (i - 1) * 30 + 37);  // 每页表 6组数据  共30行，第一页 加上表头 37行
            }
            else{
                RowCopy.copyRows(wb, "标线", "标线", 7, 31, (i - 1) * 30 + 37);//  最后一页 表 五组数据
            }
        }
        if(record == 1){
            wb.getSheet("标线").shiftRows(33, 37, -1);
        }
        //合计的
        RowCopy.copyRows(wb, "source", "标线", 0, 5,(record) * 30 + 2);
        wb.setPrintArea(wb.getSheetIndex("标线"), 0, 18, 0, record * 30+6);//第二，第三参数是列，四五是行
        for(int i = 0 ; i < 17; i++) {
            wb.getSheet("标线").setColumnWidth(i,1200); //设置第一列宽度为1200
        }
    }

    /**
     * 导出实测数据模板文件
     * @param response
     */
    @Override
    public void exportjabx(HttpServletResponse response) {
        String fileName = "02交安标线实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcJtaqssJabxVo()).finish();

    }

    /**
     * 导入数据
     * @param file
     * @param commonInfoVo
     */
    @Override
    public void importjabx(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcJtaqssJabxVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcJtaqssJabxVo>(JjgFbgcJtaqssJabxVo.class) {
                                @Override
                                public void handle(List<JjgFbgcJtaqssJabxVo> dataList) {
                                    for(JjgFbgcJtaqssJabxVo jabxVo: dataList)
                                    {
                                        JjgFbgcJtaqssJabx fbgcJtaqssJabx = new JjgFbgcJtaqssJabx();
                                        BeanUtils.copyProperties(jabxVo,fbgcJtaqssJabx);
                                        fbgcJtaqssJabx.setCreatetime(new Date());
                                        fbgcJtaqssJabx.setProname(commonInfoVo.getProname());
                                        fbgcJtaqssJabx.setHtd(commonInfoVo.getHtd());
                                        fbgcJtaqssJabx.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcJtaqssJabxMapper.insert(fbgcJtaqssJabx);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }
}
