package glgc.jjgys.system.service.impl;

import cn.hutool.core.io.resource.ClassPathResource;
import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLjgcPspqhd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.ljgc.JjgFbgcLjgcPspqhdVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcLjgcPspqhdMapper;
import glgc.jjgys.system.service.JjgFbgcLjgcPspqhdService;
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
 * @since 2023-02-16
 */
@Service
public class JjgFbgcLjgcPspqhdServiceImpl extends ServiceImpl<JjgFbgcLjgcPspqhdMapper, JjgFbgcLjgcPspqhd> implements JjgFbgcLjgcPspqhdService {

    @Autowired
    private JjgFbgcLjgcPspqhdMapper jjgFbgcLjgcPspqhdMapper;

    private static XSSFWorkbook wb = null;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;


    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        QueryWrapper<JjgFbgcLjgcPspqhd> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("zh","bw");
        List<JjgFbgcLjgcPspqhd> data = jjgFbgcLjgcPspqhdMapper.selectList(wrapper);
        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"05路基排水铺砌厚度.xlsx");
        if (data == null || data.size()==0){
            return;
        }else {
            //存放鉴定表的目录
            File fdir = new File(filepath+File.separator+proname+File.separator+htd);
            if(!fdir.exists()){
                //创建文件根目录
                fdir.mkdirs();
            }
            File directory = new File("service-system/src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String name = "排水铺砌厚度.xlsx";
            String path =reportPath+File.separator+name;
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            createTable(gettableNum(data.size()));
            if(DBtoExcel(data,proname,htd,fbgc)){
                //设置公式,计算合格点数
                calculateSheet(wb.getSheet("排水铺砌厚度"));
                System.out.println(wb.getNumberOfSheets());
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

    private void calculateSheet(XSSFSheet sheet) {
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
                row.getCell(6).setCellFormula("IF("
                        +row.getCell(4).getReference()+">="
                        +row.getCell(3).getReference()+",\"√\",\"\")");//G=IF(E6>=D6,"√","")
                row.getCell(7).setCellFormula("IF("
                        +row.getCell(4).getReference()+">="
                        +row.getCell(3).getReference()+",\"\",\"×\")");//H=IF((E6>=D6),"","X")
                XSSFFormulaEvaluator evaluate = new XSSFFormulaEvaluator(wb);
                evaluate.evaluateFormulaCell(row.getCell(7));
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

    public boolean DBtoExcel(List<JjgFbgcLjgcPspqhd> data,String proname,String htd,String fbgc) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        String time = simpleDateFormat.format(data.get(0).getJcsj());
        for(JjgFbgcLjgcPspqhd row:data){
            Date jcsj = row.getJcsj();
            JjgFbgcCommonUtils.getLastDate(time,simpleDateFormat.format(jcsj));
        }
        XSSFSheet sheet = wb.getSheet("排水铺砌厚度");

        String zh = data.get(0).getZh();
        String position = data.get(0).getBw();
        String categroy = data.get(0).getLb();

        int start = 5;
        int positionstart = 5;
        int categroystart = 5;

        int index = 0;
        ArrayList<String> compaction = new ArrayList<String>();
        sheet.getRow(1).getCell(1).setCellValue(proname);
        sheet.getRow(1).getCell(6).setCellValue(htd);
        sheet.getRow(2).getCell(1).setCellValue(fbgc);
        sheet.getRow(2).getCell(6).setCellValue(time);
        for(JjgFbgcLjgcPspqhd row:data){
            if(zh.equals(row.getZh())){  //同一桩号区间
                fillCommonCellData(sheet,index, row);
                if(!position.equals(row.getBw())){
                    if(categroystart < 5+index-1){  //合并categroy
                        sheet.addMergedRegion(new CellRangeAddress(categroystart, 5 + index-1 , 2, 2));
                    }
                    sheet.getRow(categroystart).getCell(2).setCellValue(categroy);
                    categroystart = index+5;
                    categroy = row.getLb();

                    if(positionstart < 5+index-1){
                        sheet.addMergedRegion(new CellRangeAddress(positionstart, 5 + index-1 , 1, 1));
                    }
                    sheet.getRow(positionstart).getCell(1).setCellValue(position);
                    positionstart = index+5;
                    position = row.getBw();
                }
                else{  //同一位置
                    if(!categroy.equals(row.getLb())){
                        sheet.addMergedRegion(new CellRangeAddress(categroystart, 5 + index-1 , 2, 2));
                        sheet.getRow(categroystart).getCell(2).setCellValue(categroy);
                        categroystart = index+5;
                        categroy = row.getLb();
                    }
                }
                index++;
            }
            else{
                if(categroystart < 5+index-1){
                    sheet.addMergedRegion(new CellRangeAddress(categroystart, 5 + index-1 , 2, 2));
                }
                sheet.getRow(categroystart).getCell(2).setCellValue(categroy);

                if(positionstart < 5+index-1){
                    sheet.addMergedRegion(new CellRangeAddress(positionstart, 5 + index-1 , 1, 1));
                }
                sheet.getRow(positionstart).getCell(1).setCellValue(position);
                if(start < 5+index-1){
                    sheet.addMergedRegion(new CellRangeAddress(start, 5 + index-1 , 0, 0));
                }
                categroystart = index+5;
                categroy = row.getLb();
                sheet.getRow(start).getCell(0).setCellValue(row.getZh());
                zh = row.getZh();
                position = row.getBw();
                start = index+5;
                positionstart = start;
                categroystart = positionstart;
                fillCommonCellData(sheet,index, row);
                index ++;
            }
        }
        if(categroystart < 5+index-1){
            sheet.addMergedRegion(new CellRangeAddress(categroystart, 5 + index-1 , 2, 2));
        }
        sheet.getRow(categroystart).getCell(2).setCellValue(categroy);
        if(positionstart < 5+index-1){
            sheet.addMergedRegion(new CellRangeAddress(positionstart, 5 + index-1 , 1, 1));
        }
        sheet.getRow(positionstart).getCell(1).setCellValue(position);
        if(start < 5+index-1){
            sheet.addMergedRegion(new CellRangeAddress(start, 5 + index-1 , 0, 0));
        }
        sheet.getRow(start).getCell(0).setCellValue(zh);
        return true;
    }

    public void fillCommonCellData(XSSFSheet sheet, int index, JjgFbgcLjgcPspqhd row) {
        sheet.getRow(index+5).getCell(0).setCellValue(row.getZh());
        sheet.getRow(index+5).getCell(2).setCellValue(row.getLb());//类别
        sheet.getRow(index+5).getCell(3).setCellValue(Double.parseDouble(row.getSjz()));
        sheet.getRow(index+5).getCell(4).setCellValue(Double.parseDouble(row.getScz()));
        sheet.getRow(index+5).getCell(5).setCellValue(row.getYxwc());

    }

    public int gettableNum(int size){
        return size%(29-6) <= (27-6) ? size/(29-6)+1 : size/(29-6)+2;
    }

    public void createTable(int tableNum) throws IOException {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            if(i < record-1){
                RowCopy.copyRows(wb, "排水铺砌厚度", "排水铺砌厚度", 5, 33, (i - 1) * (29-6) + 34);
            }
            else{
                RowCopy.copyRows(wb, "排水铺砌厚度", "排水铺砌厚度", 5, 31, (i - 1) * (29-6) + 34);
            }
        }
        if(record == 1){
            wb.getSheet("排水铺砌厚度").shiftRows(33-6, 34-6, -1);
        }
        RowCopy.copyRows(wb, "source", "排水铺砌厚度", 0, 1,(record) * (29-6) + 3);
        wb.setPrintArea(wb.getSheetIndex("排水铺砌厚度"), 0, 7, 0,(record) * (29-6) + 4);
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        String title = "排水工程铺砌厚度质量鉴定表";
        String sheetname = "排水铺砌厚度";
        //获取鉴定表文件
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"05路基排水铺砌厚度.xlsx");
        if(!f.exists()){
            return null;
        }else {
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

    }

    @Override
    public void exportpspqhd(HttpServletResponse response) {
        String fileName = "排水铺砌厚度实测数据";
        String sheetName = "排水铺砌厚度实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcLjgcPspqhdVo()).finish();
    }

    @Override
    public void importpspqhd(MultipartFile file,CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcLjgcPspqhdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcLjgcPspqhdVo>(JjgFbgcLjgcPspqhdVo.class) {
                                @Override
                                public void handle(List<JjgFbgcLjgcPspqhdVo> dataList) {
                                    for(JjgFbgcLjgcPspqhdVo pspqhdVo: dataList)
                                    {
                                        JjgFbgcLjgcPspqhd fbgcLjgcPspqhd = new JjgFbgcLjgcPspqhd();
                                        BeanUtils.copyProperties(pspqhdVo,fbgcLjgcPspqhd);
                                        fbgcLjgcPspqhd.setCreatetime(new Date());
                                        fbgcLjgcPspqhd.setProname(commonInfoVo.getProname());
                                        fbgcLjgcPspqhd.setHtd(commonInfoVo.getHtd());
                                        fbgcLjgcPspqhd.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcLjgcPspqhdMapper.insert(fbgcLjgcPspqhd);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }
}
