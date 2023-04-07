package glgc.jjgys.system.service.impl;

import cn.hutool.core.io.resource.ClassPathResource;
import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLjgcHdjgcc;
import glgc.jjgys.model.project.JjgFbgcLjgcPsdmcc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.ljgc.JjgFbgcLjgcHdjgccVo;
import glgc.jjgys.model.projectvo.ljgc.JjgFbgcLjgcPsdmccVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcLjgcPsdmccMapper;
import glgc.jjgys.system.service.JjgFbgcLjgcPsdmccService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
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
public class JjgFbgcLjgcPsdmccServiceImpl extends ServiceImpl<JjgFbgcLjgcPsdmccMapper, JjgFbgcLjgcPsdmcc> implements JjgFbgcLjgcPsdmccService {

    @Autowired
    private JjgFbgcLjgcPsdmccMapper jjgFbgcLjgcPsdmccMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    private static XSSFWorkbook wb = null;


    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获得数据
        QueryWrapper<JjgFbgcLjgcPsdmcc> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("zh","lb");
        List<JjgFbgcLjgcPsdmcc> data = jjgFbgcLjgcPsdmccMapper.selectList(wrapper);
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"04路基排水断面尺寸.xlsx");
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
            String name = "排水断面尺寸.xlsx";
            String path =reportPath+File.separator+name;
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream in = new FileInputStream(f);
            wb = new XSSFWorkbook(in);
            createTable(gettableNum(data.size()));
            String sheetname = "排水断面尺寸";
            if(DBtoExcel(data,proname,htd,fbgc,sheetname)){
                calculateSheet(wb.getSheet("排水断面尺寸"));
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

    private void calculateSheet(XSSFSheet sheet) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        boolean flag = false;
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
                    row.getCell(6).setCellFormula(
                            "IF("+row.getCell(3).getReference()+"=\"-\",\"\",IF(("+row.getCell(4).getReference()+"+"+JjgFbgcCommonUtils.getAllowError1(row.getCell(5).toString())+">="+
                                    row.getCell(3).getReference()+")*("+row.getCell(4).getReference()+"-"+JjgFbgcCommonUtils.getAllowError1(row.getCell(5).toString())+"<="+
                                    row.getCell(3).getReference()+"),\"√\",\"\"))");
                    row.getCell(7).setCellFormula(
                            "IF("+row.getCell(3).getReference()+"=\"-\",\"\",IF(("+row.getCell(4).getReference()+"+"+JjgFbgcCommonUtils.getAllowError1(row.getCell(5).toString())+">="+
                                    row.getCell(3).getReference()+")*("+row.getCell(4).getReference()+"-"+JjgFbgcCommonUtils.getAllowError1(row.getCell(5).toString())+"<="+
                                    row.getCell(3).getReference()+"),\"\",\"×\"))");
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

    public boolean DBtoExcel(List<JjgFbgcLjgcPsdmcc> data,String proname,String htd,String fbgc,String sheetname) throws IOException, ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        String time = simpleDateFormat.format(data.get(0).getJcsj());
        for(JjgFbgcLjgcPsdmcc row:data){
            Date jcsj = row.getJcsj();
            JjgFbgcCommonUtils.getLastDate(time,simpleDateFormat.format(jcsj));
        }
        XSSFSheet sheet = wb.getSheet(sheetname);

        String zh = data.get(0).getZh();
        String position = data.get(0).getBw();
        String categroy = data.get(0).getLb();

        int start = 5;
        int positionstart = 5;
        int categroystart = 5;

        int index = 0;
        String pileNO = "";
        ArrayList<String> compaction = new ArrayList<String>();
        sheet.getRow(1).getCell(1).setCellValue(proname);
        sheet.getRow(1).getCell(6).setCellValue(htd);
        sheet.getRow(2).getCell(1).setCellValue(fbgc);
        sheet.getRow(2).getCell(6).setCellValue(time);
        for(JjgFbgcLjgcPsdmcc row:data){
            if(zh.equals(row.getZh())){
                fillCommonCellData(sheet, index, row);
                if(!position.equals(row.getBw())){
                    sheet.addMergedRegion(new CellRangeAddress(categroystart, 5 + index-1 , 2, 2));
                    sheet.getRow(categroystart).getCell(2).setCellValue(categroy);
                    categroystart = index+5;
                    categroy = row.getLb();

                    sheet.addMergedRegion(new CellRangeAddress(positionstart, 5 + index-1 , 1, 1));
                    sheet.getRow(positionstart).getCell(1).setCellValue(position);
                    positionstart = index+5;
                    position = row.getBw();
                }
                else{
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
                sheet.addMergedRegion(new CellRangeAddress(categroystart, 5 + index-1 , 2, 2));
                sheet.getRow(categroystart).getCell(2).setCellValue(categroy);
                categroystart = index+5;
                categroy = row.getLb();
                sheet.addMergedRegion(new CellRangeAddress(positionstart, 5 + index-1 , 1, 1));
                sheet.getRow(positionstart).getCell(1).setCellValue(position);
                sheet.addMergedRegion(new CellRangeAddress(start, 5 + index-1 , 0, 0));
                sheet.getRow(start).getCell(0).setCellValue(row.getZh());
                zh = row.getZh();
                position = row.getBw();
                start = index+5;
                positionstart = start;
                categroystart = positionstart;
                fillCommonCellData(sheet, index, row);
                index ++;

            }
        }
        sheet.addMergedRegion(new CellRangeAddress(categroystart, 5 + index-1 , 2, 2));
        sheet.getRow(categroystart).getCell(2).setCellValue(categroy);
        sheet.addMergedRegion(new CellRangeAddress(positionstart, 5 + index-1 , 1, 1));
        sheet.getRow(positionstart).getCell(1).setCellValue(position);
        sheet.addMergedRegion(new CellRangeAddress(start, 5 + index-1 , 0, 0));
        sheet.getRow(start).getCell(0).setCellValue(zh);
        return true;
    }

    public void fillCommonCellData(XSSFSheet sheet, int index, JjgFbgcLjgcPsdmcc row) {
        sheet.getRow(index+5).getCell(2).setCellValue(row.getLb());//类别
        if(row.getSjz()==null || "".equals(row.getSjz()))
        {
            sheet.getRow(index+5).getCell(3).setCellValue("-");
            sheet.getRow(index+5).getCell(4).setCellValue("-");
        }else
        {
            sheet.getRow(index+5).getCell(3).setCellValue(Double.parseDouble(row.getSjz()));
            sheet.getRow(index+5).getCell(4).setCellValue(Double.parseDouble(row.getScz()));
        }

        if(row.getYxwcz().equals(row.getYxwcf())){
            if("不小于设计值".equals(row.getYxwcf()) || "不小于设计值".equals(row.getYxwcz()) ){
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

    //判断要生成几个表格
    public int gettableNum(int size){
        return size%29 <= 27 ? size/29+1 : size/29+2;
    }

    public void createTable(int tableNum) throws IOException {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            if(i < record-1){
                //pStartRow 标题的前5行   34是每页结束行数的下标  30是每页数据的数量  35是没页总共的数量
                RowCopy.copyRows(wb, "排水断面尺寸", "排水断面尺寸", 5, 34, (i - 1) * 30 + 35);
            }
            else{
                RowCopy.copyRows(wb, "排水断面尺寸", "排水断面尺寸", 5, 32, (i - 1) * 30 + 35);
            }
        }
        if(record == 1){
            wb.getSheet("排水断面尺寸").shiftRows(33, 34, -1);
        }
        RowCopy.copyRows(wb, "source", "排水断面尺寸", 0, 1,(record) * 30 + 3);//好像是合计的部分
        wb.setPrintArea(wb.getSheetIndex("排水断面尺寸"), 0, 7, 0,(record) * 30 + 4);
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        String title = "结构（断面）尺寸质量鉴定表";
        String sheetname = "排水断面尺寸";
        //获取鉴定表文件
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"04路基排水断面尺寸.xlsx");
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
    public void exportpsdmcc(HttpServletResponse response) {
        String fileName = "排水断面尺寸实测数据";
        String sheetName = "排水断面尺寸实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcLjgcPsdmccVo()).finish();

    }

    @Override
    public void importpsdmcc(MultipartFile file,CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcLjgcPsdmccVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcLjgcPsdmccVo>(JjgFbgcLjgcPsdmccVo.class) {
                                @Override
                                public void handle(List<JjgFbgcLjgcPsdmccVo> dataList) {
                                    for(JjgFbgcLjgcPsdmccVo psdmccVo: dataList)
                                    {
                                        JjgFbgcLjgcPsdmcc jjgFbgcLjgcPsdmcc = new JjgFbgcLjgcPsdmcc();
                                        BeanUtils.copyProperties(psdmccVo,jjgFbgcLjgcPsdmcc);
                                        jjgFbgcLjgcPsdmcc.setCreatetime(new Date());
                                        jjgFbgcLjgcPsdmcc.setProname(commonInfoVo.getProname());
                                        jjgFbgcLjgcPsdmcc.setHtd(commonInfoVo.getHtd());
                                        jjgFbgcLjgcPsdmcc.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcLjgcPsdmccMapper.insert(jjgFbgcLjgcPsdmcc);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }
}
