package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcJtaqssJathldmcc;
import glgc.jjgys.model.project.JjgFbgcLjgcHdjgcc;
import glgc.jjgys.model.project.JjgFbgcLjgcXqjgcc;
import glgc.jjgys.model.projectvo.jagc.JjgFbgcJtaqssJathldmccVo;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.ljgc.JjgFbgcLjgcHdjgccVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcJtaqssJathldmccMapper;
import glgc.jjgys.system.service.JjgFbgcJtaqssJathldmccService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
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
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-03-01
 */
@Service
public class JjgFbgcJtaqssJathldmccServiceImpl extends ServiceImpl<JjgFbgcJtaqssJathldmccMapper, JjgFbgcJtaqssJathldmcc> implements JjgFbgcJtaqssJathldmccService {

    @Autowired
    private JjgFbgcJtaqssJathldmccMapper jjgFbgcJtaqssJathldmccMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException {
        XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获得数据
        QueryWrapper<JjgFbgcJtaqssJathldmcc> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("zh","bw","lb");
        List<JjgFbgcJtaqssJathldmcc> data = jjgFbgcJtaqssJathldmccMapper.selectList(wrapper);

        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"60交安砼护栏断面尺寸.xlsx");
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
            String name = "砼护栏断面尺寸.xlsx";
            String path = reportPath + File.separator + name;
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream in = new FileInputStream(f);
            wb = new XSSFWorkbook(in);
            createTable(gettableNum(data.size()),wb);
            if(DBtoExcel(data,proname,htd,fbgc,wb)){
                calculateSheet(wb.getSheet("砼护栏断面尺寸"),wb);
                for (int j = 0; j < wb.getNumberOfSheets(); j++) {   //表内公式  计算 显示结果
                    JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
                }
                FileOutputStream fileOut = new FileOutputStream(f);
                wb.write(fileOut);
                fileOut.flush();
                fileOut.close();
            }
            in.close();
            wb.close();
        }

    }

    private void calculateSheet(XSSFSheet sheet,XSSFWorkbook wb) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        boolean flag = false;
        XSSFCellStyle cellstyle = JjgFbgcCommonUtils.dBtoExcelUtils(wb);
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if(flag && "检测总点数".equals(row.getCell(0).toString())){
                rowend = sheet.getRow(i-1);
                row.getCell(1).setCellFormula("COUNT("
                        +rowstart.getCell(4).getReference()+":"
                        +rowend.getCell(4).getReference()+")");//B32=COUNT(D6:D31) B=COUNT(E6:E68)
                row.getCell(3).setCellFormula("COUNTIF("
                        +rowstart.getCell(6).getReference()+":"
                        +rowend.getCell(6).getReference()+",\"√\")");//D32=COUNTIF(H6:H31,"√") D=COUNTIF(G6:G34,"√")+COUNTIF(G41:G68,"√")
                row.getCell(6).setCellFormula("ROUND("+
                        row.getCell(3).getReference()+"/"
                        +row.getCell(1).getReference()+"*100,1)");//H32=D32/B32*100 G=ROUND(D69/B69*100,1)
                break;
            }

            if(flag && !"".equals(row.getCell(4).toString())){
                if("不小于设计值".equals(row.getCell(5).toString())){
                    row.getCell(6).setCellFormula("IF("
                            +row.getCell(4).getReference()+">="
                            +row.getCell(3).getReference()+",\"√\",\"\")");//G=IF((E41+30>=D41)*(E41-30<=D41),"√","")

                    row.getCell(7).setCellFormula("IF("
                            +row.getCell(4).getReference()+">="
                            +row.getCell(3).getReference()+",\"\",\"×\")");//H=IF((E41+30>=D41)*(E41-30<=D41),"","×")
                }
                else{

                    row.getCell(6).setCellFormula("IF(("
                            +row.getCell(3).getReference()+"+"
                            +JjgFbgcCommonUtils.getAllowError1(row.getCell(5).toString())+">="
                            +row.getCell(4).getReference()+")*("
                            +row.getCell(3).getReference()+"-"
                            +JjgFbgcCommonUtils.getAllowError2(row.getCell(5).toString())+"<="
                            +row.getCell(4).getReference()+"),\"√\",\"\")");//G=IF((D41+30>=E41)*(D41-30<=E41),"√","")

                    row.getCell(7).setCellFormula("IF(("
                            +row.getCell(3).getReference()+"+"
                            +JjgFbgcCommonUtils.getAllowError1(row.getCell(5).toString())+">="
                            +row.getCell(4).getReference()+")*("
                            +row.getCell(3).getReference()+"-"
                            +JjgFbgcCommonUtils.getAllowError2(row.getCell(5).toString())+"<="
                            +row.getCell(4).getReference()+"),\"\",\"×\")");//H=IF((E41+30>=D41)*(E41-30<=D41),"","×")

                }
                XSSFFormulaEvaluator evaluate = new XSSFFormulaEvaluator(wb);
                if("×".equals(row.getCell(7).getRawValue())){
                    row.getCell(7).setCellStyle(cellstyle);
                }
            }
            if ("桩号".equals(row.getCell(0).toString())) {
                rowstart = sheet.getRow(i+2);
                rowend = rowstart;
                i++;
                flag = true;
            }
        }
    }

    private boolean DBtoExcel(List<JjgFbgcJtaqssJathldmcc> data, String proname, String htd, String fbgc,XSSFWorkbook wb) {
        XSSFCellStyle cellstyle = JjgFbgcCommonUtils.dBtoExcelUtils(wb);
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        String testtime = simpleDateFormat.format(data.get(0).getJcrq());
        for (JjgFbgcJtaqssJathldmcc row:data){
            try {
                Date dt1 = simpleDateFormat.parse(testtime);
                Date dt2 = simpleDateFormat.parse(simpleDateFormat.format(row.getJcrq()));
                if(dt1.getTime() < dt2.getTime()){
                    testtime = simpleDateFormat.format(row.getJcrq());
                }
            } catch (ParseException e) {
                e.printStackTrace();
            }
        }
        XSSFSheet sheet = wb.getSheet("砼护栏断面尺寸");

        sheet.getRow(1).getCell(1).setCellValue(proname);
        sheet.getRow(1).getCell(6).setCellValue(htd);
        sheet.getRow(2).getCell(1).setCellValue(fbgc);
        sheet.getRow(2).getCell(6).setCellValue(testtime);//检测时间

        String zh = data.get(0).getZh();
        String bw = data.get(0).getBw();
        String lb = data.get(0).getLb();
        int start = 5;
        int categroystart = 5;
        int positionstart = 5;
        int index = 0;
        for(JjgFbgcJtaqssJathldmcc row:data){
            if(zh.equals(row.getZh())){
                fillCommonCellData(sheet,index,row);
                if(!bw.equals(row.getBw())){
                    if(categroystart < 5 + index-1){
                        sheet.addMergedRegion(new CellRangeAddress(categroystart, 5 + index-1 , 2, 2));
                    }
                    sheet.getRow(categroystart).getCell(2).setCellValue(lb);
                    categroystart = index+5;
                    lb = row.getLb();
                    if(positionstart < 5 + index-1){
                        sheet.addMergedRegion(new CellRangeAddress(positionstart, 5 + index-1 , 1, 1));
                    }
                    sheet.getRow(positionstart).getCell(1).setCellValue(bw);
                    positionstart = index+5;
                    bw = row.getBw();

                }else {
                    //部位相等的情况下  类别
                    if(!lb.equals(row.getLb())){
                        if(categroystart < 5 + index-1){
                            sheet.addMergedRegion(new CellRangeAddress(categroystart, 5 + index-1 , 2, 2));
                        }
                        sheet.getRow(categroystart).getCell(2).setCellValue(lb);
                        categroystart = index+5;
                        lb = row.getLb();
                    }
                }
                index++;
            }else {
                if(categroystart < 5 + index-1){
                    sheet.addMergedRegion(new CellRangeAddress(categroystart, 5 + index-1 , 2, 2));
                }
                sheet.getRow(categroystart).getCell(2).setCellValue(lb);
                categroystart = index+5;
                lb = row.getLb();
                if(positionstart < 5 + index-1){
                    sheet.addMergedRegion(new CellRangeAddress(positionstart, 5 + index-1 , 1, 1));
                }

                sheet.getRow(positionstart).getCell(1).setCellValue(bw);
                if(start < 5 + index-1){
                    sheet.addMergedRegion(new CellRangeAddress(start, 5 + index-1 , 0, 0));
                }

                sheet.getRow(start).getCell(0).setCellValue(row.getZh());
                zh = row.getZh();
                bw = row.getBw();
                start = index+5;
                positionstart = start;
                categroystart = positionstart;
                fillCommonCellData(sheet,index,row);
                index ++;

            }
        }
        if(categroystart < 5 + index-1){
            sheet.addMergedRegion(new CellRangeAddress(categroystart, 5 + index-1 , 2, 2));
        }

        sheet.getRow(categroystart).getCell(2).setCellValue(lb);
        if(positionstart < 5 + index-1){
            sheet.addMergedRegion(new CellRangeAddress(positionstart, 5 + index-1 , 1, 1));
        }

        sheet.getRow(positionstart).getCell(1).setCellValue(bw);
        if(start < 5 + index-1){
            sheet.addMergedRegion(new CellRangeAddress(start, 5 + index-1 , 0, 0));
        }

        sheet.getRow(start).getCell(0).setCellValue(zh);
        return true;
    }

    public void fillCommonCellData(XSSFSheet sheet, int index, JjgFbgcJtaqssJathldmcc row) {
        sheet.getRow(index+5).getCell(0).setCellValue(row.getZh());//桩号
        sheet.getRow(index+5).getCell(1).setCellValue(row.getBw());//部位
        sheet.getRow(index+5).getCell(2).setCellValue(row.getLb());//类别
        sheet.getRow(index+5).getCell(3).setCellValue(Double.parseDouble(row.getSjz()));//设计值
        sheet.getRow(index+5).getCell(4).setCellValue(Double.parseDouble(row.getScz()));//实测值
        if(row.getYxwcz().equals(row.getYxwcf())){
            if("0".equals(row.getYxwcf())){
                sheet.getRow(index+5).getCell(5).setCellValue("不小于设计值");
            }
            else{
                sheet.getRow(index+5).getCell(5).setCellValue("±"+row.getYxwcz());
            }
        }
        else{
            sheet.getRow(index+5).getCell(5).setCellValue("+"+row.getYxwcz()+",-"+row.getYxwcf());
        }

    }

    private void createTable(int tableNum,XSSFWorkbook wb) {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            if(i < record-1){
                RowCopy.copyRows(wb, "砼护栏断面尺寸", "砼护栏断面尺寸", 5, 33, (i - 1) * 29 + 34);
            }
            else{
                RowCopy.copyRows(wb, "砼护栏断面尺寸", "砼护栏断面尺寸", 5, 31, (i - 1) * 29 + 34);
            }
        }
        if(record == 1){
            wb.getSheet("砼护栏断面尺寸").shiftRows(33, 34, -1);
        }
        RowCopy.copyRows(wb, "source", "砼护栏断面尺寸", 0, 1,(record) * 29 + 3);
        wb.setPrintArea(wb.getSheetIndex("砼护栏断面尺寸"), 0, 7, 0,(record) * 29 + 4);

    }

    private int gettableNum(int size) {
        return size%29 <= 27 ? size/29+1 : size/29+2;
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        String title = "结构（断面）尺寸质量鉴定表";
        String sheetname = "砼护栏断面尺寸";
        //获取鉴定表文件
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"60交安砼护栏断面尺寸.xlsx");
        if(!f.exists()){
            return null;
        }
        Map<String,Object> map = new HashMap<>();
        map.put("proname",proname);
        map.put("title",title);
        map.put("htd",htd);
        map.put("fbgc",fbgc);
        map.put("f",f);
        map.put("sheetname",sheetname);
        List<Map<String, Object>> mapList = JjgFbgcCommonUtils.getdmcjjcjg(map);
        return mapList;
    }

    @Override
    public void exportjathldmcc(HttpServletResponse response) {
        String fileName = "交安砼护栏断面尺寸实测数据";
        String sheetName = "交安砼护栏断面尺寸实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcJtaqssJathldmccVo()).finish();

    }

    @Override
    public void importjathldmcc(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcJtaqssJathldmccVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcJtaqssJathldmccVo>(JjgFbgcJtaqssJathldmccVo.class) {
                                @Override
                                public void handle(List<JjgFbgcJtaqssJathldmccVo> dataList) {
                                    for(JjgFbgcJtaqssJathldmccVo jathldmccVo: dataList)
                                    {
                                        JjgFbgcJtaqssJathldmcc fbgcJtaqssJathldmcc = new JjgFbgcJtaqssJathldmcc();
                                        BeanUtils.copyProperties(jathldmccVo,fbgcJtaqssJathldmcc);
                                        fbgcJtaqssJathldmcc.setCreatetime(new Date());
                                        fbgcJtaqssJathldmcc.setProname(commonInfoVo.getProname());
                                        fbgcJtaqssJathldmcc.setHtd(commonInfoVo.getHtd());
                                        fbgcJtaqssJathldmcc.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcJtaqssJathldmccMapper.insert(fbgcJtaqssJathldmcc);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }
}
