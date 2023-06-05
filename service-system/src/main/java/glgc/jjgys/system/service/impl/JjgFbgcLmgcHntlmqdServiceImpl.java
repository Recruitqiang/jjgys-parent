package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLmgcHntlmqd;
import glgc.jjgys.model.project.JjgFbgcLmgcLmssxs;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.lmgc.JjgFbgcLmgcHntlmqdVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcLmgcHntlmqdMapper;
import glgc.jjgys.system.service.JjgFbgcLmgcHntlmqdService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
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
 * @since 2023-04-20
 */
@Service
public class JjgFbgcLmgcHntlmqdServiceImpl extends ServiceImpl<JjgFbgcLmgcHntlmqdMapper, JjgFbgcLmgcHntlmqd> implements JjgFbgcLmgcHntlmqdService {

    @Autowired
    private JjgFbgcLmgcHntlmqdMapper jjgFbgcLmgcHntlmqdMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获取数据
        QueryWrapper<JjgFbgcLmgcHntlmqd> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("zh");
        List<JjgFbgcLmgcHntlmqd> data = jjgFbgcLmgcHntlmqdMapper.selectList(wrapper);
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"16混凝土路面强度.xlsx");
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
            String path =reportPath +File.separator+ "混凝土路面强度.xlsx";
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            createTable(gettableNum(data.size()),wb);
            if(DBtoExcel(data,wb)){
                for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                    if (shouldBeCalculate(wb.getSheetAt(j))) {
                        calculateOneSheet(wb.getSheetAt(j));
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
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if(flag && row.getCell(0) != null && !"".equals(row.getCell(0).toString())){
                if(!name.equals(
                        row.getCell(1).toString().contains("Z")?
                                row.getCell(1).toString().replaceAll("[^\u4E00-\u9FA5]", "")+"Z" :
                                row.getCell(1).toString().replaceAll("[^\u4E00-\u9FA5]", "")+"Y"
                ) && !"".equals(name)){
                    endrow = sheet.getRow(i-1);
                    startrow.createCell(10).setCellFormula("COUNT("
                            +startrow.getCell(7).getReference()+":"
                            +endrow.getCell(7).getReference()+")");//=COUNT(H7:H27)
                    startrow.createCell(11).setCellFormula("COUNTIF("
                            +startrow.getCell(8).getReference()+":"
                            +endrow.getCell(8).getReference()+",\"√\")");//=COUNTIF(I7:I27,"√")
                    startrow.createCell(12).setCellFormula(startrow.getCell(11).getReference()+"*100/"
                            +startrow.getCell(10).getReference());
                    name = row.getCell(1).toString().contains("Z")?
                            row.getCell(1).toString().replaceAll("[^\u4E00-\u9FA5]", "")+"Z" :
                            row.getCell(1).toString().replaceAll("[^\u4E00-\u9FA5]", "")+"Y";
                    startrow = row;
                }
                if("".equals(name)){
                    /*
                     * 隧道要分左右幅统计，但渗水系数没有分开统计，所以要根据桩号的z/y来判断
                     */
                    name = row.getCell(1).toString().contains("Z")?
                            row.getCell(1).toString().replaceAll("[^\u4E00-\u9FA5]", "")+"Z" :
                            row.getCell(1).toString().replaceAll("[^\u4E00-\u9FA5]", "")+"Y";
                    startrow = row;
                }

            }
            if ("试件编号".equals(row.getCell(0).toString())) {
                flag = true;
                i++;
            }
            if ("合计".equals(row.getCell(0).toString())) {
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
        boolean flag = false;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if (!"".equals(row.getCell(1).toString()) && flag) {
                row.getCell(6).setCellFormula(
                        "ROUND(2*" + row.getCell(5).getReference() + "/(PI()*"
                                + row.getCell(3).getReference() + "*"
                                + row.getCell(4).getReference() + "),2)");// G=ROUND(2*F7/(PI()*D7*E7),2)
                row.getCell(7).setCellFormula(
                        "ROUND((" + row.getCell(6).getReference()
                                + "^0.871)*1.868,2)");// H=ROUND((G7^0.871)*1.868,2)
                row.getCell(8).setCellFormula(
                        "IF(" + row.getCell(7).getReference() + ":"
                                + row.getCell(7).getReference() + ">="
                                + row.getCell(2).getReference()
                                + ",\"√\",\"\")");// I=IF(H7:H7>=C4,"√","")
                row.getCell(9).setCellFormula(
                        "IF(" + row.getCell(8).getReference()
                                + "=\"\",\"×\",\"\")");// J=IF(I7="","×","")
            }
            if ("合计".equals(row.getCell(0).toString())) {
                rowend = sheet.getRow(row.getRowNum() - 1);
                calculateTotal(sheet, i, rowstart, rowend);
            }
            if ("试件编号".equals(row.getCell(0).toString())) {
                flag = true;
                i += 1;
                rowstart = sheet.getRow(i + 1);
            }

        }
    }

    /**
     *
     * @param sheet
     * @param i
     * @param rowstart
     * @param rowend
     */
    private void calculateTotal(XSSFSheet sheet, int i, XSSFRow rowstart, XSSFRow rowend) {
        sheet.getRow(i)
                .getCell(4)
                .setCellFormula(
                        "COUNT(" + rowstart.getCell(7).getReference() + ":"
                                + rowend.getCell(7).getReference() + ")");// =COUNT(H7:H32)
        sheet.getRow(i)
                .getCell(6)
                .setCellFormula(
                        "COUNTIF(" + rowstart.getCell(8).getReference() + ":"
                                + rowend.getCell(8).getReference() + ",\"√\")");// =COUNTIF(I7:I32,"√")
        sheet.getRow(i)
                .getCell(8)
                .setCellFormula(
                        sheet.getRow(i).getCell(6).getReference() + "/"
                                + sheet.getRow(i).getCell(4).getReference()
                                + "*100");// =G33/E33*100
        sheet.getRow(i)
                .createCell(10)
                .setCellFormula(
                        "MAX(" + rowstart.getCell(7).getReference() + ":"
                                + rowend.getCell(7).getReference() + ")");//最大值
        sheet.getRow(i)
                .createCell(11)
                .setCellFormula(
                        "MIN(" + rowstart.getCell(7).getReference() + ":"
                                + rowend.getCell(7).getReference() + ")");// 最小值
        sheet.getRow(i)
                .createCell(12)
                .setCellFormula(
                        "AVERAGE(" + rowstart.getCell(7).getReference() + ":"
                                + rowend.getCell(7).getReference() + ")");// 平均值
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
    private boolean DBtoExcel(List<JjgFbgcLmgcHntlmqd> data, XSSFWorkbook wb) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        XSSFSheet sheet = wb.getSheet("砼路面抗弯拉强度");
        int index = 6;
        sheet.getRow(1).getCell(2).setCellValue(data.get(0).getProname());
        sheet.getRow(1).getCell(7).setCellValue(data.get(0).getHtd());
        sheet.getRow(2).getCell(2).setCellValue(data.get(0).getFbgc());
        sheet.getRow(2).getCell(7).setCellValue("路面面层");
        sheet.getRow(3).getCell(2).setCellValue(Double.valueOf(data.get(0).getLmqdgdz()));
        String date = simpleDateFormat.format(data.get(0).getJcsj());
        for(int i =1; i < data.size(); i++){
            date = JjgFbgcCommonUtils.getLastDate(date, simpleDateFormat.format(data.get(i).getJcsj()));
        }
        sheet.getRow(3).getCell(7).setCellValue(date);
        for(int i =0; i < data.size(); i++){
            sheet.getRow(index+i).getCell(0).setCellValue(i+1);
            sheet.addMergedRegion(new CellRangeAddress(index+i, index+i, 1, 2));
            sheet.getRow(index+i).getCell(1).setCellValue(data.get(i).getQywzmc()+ data.get(i).getZh());
            sheet.getRow(index+i).getCell(3).setCellValue(Double.parseDouble(data.get(i).getSypjzj()));
            sheet.getRow(index+i).getCell(4).setCellValue(Double.parseDouble(data.get(i).getSypjhd()));
            sheet.getRow(index+i).getCell(5).setCellValue(Double.parseDouble(data.get(i).getJxhz()));
        }
        return true;
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
                RowCopy.copyRows(wb, "砼路面抗弯拉强度", "砼路面抗弯拉强度", 6, 27, (i - 1) * 22 + 28);
            }
            else{
                RowCopy.copyRows(wb, "砼路面抗弯拉强度", "砼路面抗弯拉强度", 6, 26, (i - 1) * 22 + 28);
            }
        }
        if(record == 1){
            wb.getSheet("砼路面抗弯拉强度").shiftRows(28, 28, -1);
        }
        RowCopy.copyRows(wb, "source", "砼路面抗弯拉强度", 0, 0,(record) * 22 + 5);
        wb.setPrintArea(wb.getSheetIndex("砼路面抗弯拉强度"), 0, 9, 0,(record) * 22 + 5);
    }

    /**
     *
     * @param size
     * @return
     */
    private int gettableNum(int size) {
        return size%22 <= 21 ? size/22+1 : size/22+2;
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat(".00");
        DecimalFormat decf = new DecimalFormat("0.##");
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        String title = "混凝土路面弯拉强度鉴定表";
        String sheetname = "砼路面抗弯拉强度";

        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "16混凝土路面强度.xlsx");
        if (!f.exists()) {
            return null;
        } else {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
            XSSFSheet slSheet = wb.getSheet(sheetname);
            XSSFCell bt = slSheet.getRow(0).getCell(0);//标题
            XSSFCell xmname = slSheet.getRow(1).getCell(2);//项目名
            XSSFCell htdname = slSheet.getRow(1).getCell(7);//合同段名
            XSSFCell hd = slSheet.getRow(2).getCell(2);//分布工程名
            List<Map<String, Object>> mapList = new ArrayList<>();
            Map<String, Object> jgmap = new HashMap<>();

            if (proname.equals(xmname.toString()) && title.equals(bt.toString()) && htd.equals(htdname.toString()) && fbgc.equals(hd.toString())) {
                //获取到最后一行
                int lastRowNum = slSheet.getLastRowNum();
                slSheet.getRow(lastRowNum).getCell(4).setCellType(CellType.STRING);//总点数
                slSheet.getRow(lastRowNum).getCell(6).setCellType(CellType.STRING);//合格点数
                slSheet.getRow(lastRowNum).getCell(8).setCellType(CellType.STRING);//合格率
                double zds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(4).getStringCellValue());
                double hgds = Double.valueOf(slSheet.getRow(lastRowNum).getCell(6).getStringCellValue());
                double hgl = Double.valueOf(slSheet.getRow(lastRowNum).getCell(8).getStringCellValue());
                String zdsz1 = decf.format(zds);
                String hgdsz1 = decf.format(hgds);
                String hglz1 = df.format(hgl);
                jgmap.put("总点数", zdsz1);
                jgmap.put("合格点数", hgdsz1);
                jgmap.put("合格率", hglz1);
                mapList.add(jgmap);
            }else {
                jgmap.put("总点数", 0);
                jgmap.put("合格点数", 0);
                jgmap.put("合格率", 0);
                mapList.add(jgmap);
            }
            return mapList;

        }
    }

    @Override
    public void exporthntlmqd(HttpServletResponse response) {
        String fileName = "05混凝土路面强度实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcLmgcHntlmqdVo()).finish();

    }

    @Override
    public void importhntlmqd(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcLmgcHntlmqdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcLmgcHntlmqdVo>(JjgFbgcLmgcHntlmqdVo.class) {
                                @Override
                                public void handle(List<JjgFbgcLmgcHntlmqdVo> dataList) {
                                    for(JjgFbgcLmgcHntlmqdVo hntlmqdVo: dataList)
                                    {
                                        JjgFbgcLmgcHntlmqd hntlmqd = new JjgFbgcLmgcHntlmqd();
                                        BeanUtils.copyProperties(hntlmqdVo,hntlmqd);
                                        hntlmqd.setCreatetime(new Date());
                                        hntlmqd.setProname(commonInfoVo.getProname());
                                        hntlmqd.setHtd(commonInfoVo.getHtd());
                                        hntlmqd.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcLmgcHntlmqdMapper.insert(hntlmqd);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }
}
