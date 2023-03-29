package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLjgcLjbp;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.ljgc.JjgFbgcLjgcLjbpVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcLjgcLjbpMapper;
import glgc.jjgys.system.service.JjgFbgcLjgcLjbpService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
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
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-02-21
 */
@Service
public class JjgFbgcLjgcLjbpServiceImpl extends ServiceImpl<JjgFbgcLjgcLjbpMapper, JjgFbgcLjgcLjbp> implements JjgFbgcLjgcLjbpService {

    @Autowired
    private JjgFbgcLjgcLjbpMapper jjgFbgcLjgcLjbpMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    private static XSSFWorkbook wb = null;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获得数据
        QueryWrapper<JjgFbgcLjgcLjbp> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("zh","wz");
        List<JjgFbgcLjgcLjbp> data = jjgFbgcLjgcLjbpMapper.selectList(wrapper);

        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"03路基边坡.xlsx");
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
            String path =reportPath + "/路基边坡.xlsx";
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream in = new FileInputStream(f);
            wb = new XSSFWorkbook(in);
            createTable(gettableNum(data.size()));
            String sheetname = "路基边坡";
            if(DBtoExcel(data,proname,htd,fbgc,sheetname)){
                calculateSheet(wb.getSheet("路基边坡"));
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
        XSSFCellStyle cellStyle = JjgFbgcCommonUtils.dBtoExcelUtils(wb);
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if(flag && "检测总点数".equals(row.getCell(0).toString())){
                rowend = sheet.getRow(i-1);
                row.getCell(1).setCellFormula("COUNT("
                        +rowstart.getCell(3).getReference()+":"
                        +rowend.getCell(3).getReference()+")");//B32=COUNT(D6:D31)
                row.getCell(3).setCellFormula("COUNTIF("
                        +rowstart.getCell(7).getReference()+":"
                        +rowend.getCell(7).getReference()+",\"√\")");//D32=COUNTIF(H6:H31,"√")
                row.getCell(7).setCellFormula(
                        row.getCell(3).getReference()+"/"
                                +row.getCell(1).getReference()+"*100");//H32=D32/B32*100
                break;
            }
            if(flag && !"".equals(row.getCell(1).toString())){
                row.createCell(9).setCellFormula("ROUND("+row.getCell(3).getReference()+"/"+row.getCell(4).getReference()+",2)");//J6=ROUND(D6/E6,2)

                row.getCell(5).setCellFormula("\"1\"&\":\"&"
                        +row.getCell(9).getReference());//F6="1"&":"&J6
                row.getCell(7).setCellFormula("IF(("
                        +row.getCell(9).getReference()+"-"
                        +row.getCell(10).getReference()+")>=0,\"√\",\"\")");//H6=IF((J6-K6)>=0,"√","")
                row.getCell(8).setCellFormula("IF(("
                        +row.getCell(9).getReference()+"-"
                        +row.getCell(10).getReference()+")>=0,\"\",\"×\")");//I6=IF((J6-K6)>=0,"","×")
                XSSFFormulaEvaluator evaluate = new XSSFFormulaEvaluator(wb);
                if("×".equals(row.getCell(8).getRawValue())){
                    row.getCell(8).setCellStyle(cellStyle);
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


    public boolean DBtoExcel(List<JjgFbgcLjgcLjbp> data,String proname,String htd,String fbgc,String sheetname) throws IOException, ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        String time = simpleDateFormat.format(data.get(0).getJcsj());
        for(JjgFbgcLjgcLjbp row:data){
            Date jcsj = row.getJcsj();
            JjgFbgcCommonUtils.getLastDate(time,simpleDateFormat.format(jcsj));
        }
        XSSFSheet sheet = wb.getSheet(sheetname);

        String zh = data.get(0).getZh();
        int start = 5;
        int end = 5;
        int index = 0;
        String pileNO = "";
        ArrayList<String> compaction = new ArrayList<String>();
        sheet.getRow(1).getCell(1).setCellValue(proname);
        sheet.getRow(1).getCell(6).setCellValue(htd);
        sheet.getRow(2).getCell(1).setCellValue(fbgc);
        sheet.getRow(2).getCell(6).setCellValue(time);
        for(JjgFbgcLjgcLjbp row:data){
            if(zh.equals(row.getZh())){
                fillCommonCellData(sheet,index, row);
                index++;
            }
            else{
                sheet.addMergedRegion(new CellRangeAddress(start, 5 + index-1 , 0, 0));
                sheet.getRow(start).getCell(0).setCellValue(zh);
                zh = row.getZh();
                start = index+5;
                fillCommonCellData(sheet,index, row);
                index ++;
            }
        }
        sheet.addMergedRegion(new CellRangeAddress(start, 5 + index-1 , 0, 0));
        sheet.getRow(start).getCell(0).setCellValue(zh);
        return true;
    }

    public void fillCommonCellData(XSSFSheet sheet,int index,JjgFbgcLjgcLjbp row) {
        sheet.getRow(index+5).getCell(1).setCellValue(row.getWz());
        sheet.getRow(index+5).createCell(10).setCellValue(bizhizhuanxiaoshu(row.getSjz()));//K
        sheet.getRow(index+5).getCell(2).setCellFormula("\"1\"&\":\"&"
                +sheet.getRow(index+5).getCell(10).getReference());//="1"&":"&K6
        sheet.getRow(index+5).getCell(3).setCellValue(Double.parseDouble(row.getLength()));
        sheet.getRow(index+5).getCell(4).setCellValue(Double.parseDouble(row.getHigh()));
        sheet.getRow(index+5).getCell(6).setCellValue("不陡于设计");

    }

    //比值转小数
    private double bizhizhuanxiaoshu(String str){
        double fenzi=Double.parseDouble(str.substring(0, str.indexOf(':')));
        double fenmu=Double.parseDouble(str.substring(str.indexOf(':')+1,str.length()));
        return fenmu;
    }

    public void createTable(int tableNum) throws IOException {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            if(i < record-1){
                RowCopy.copyRows(wb, "路基边坡", "路基边坡", 5, 34, (i - 1) * 30 + 35);
            }
            else{
                RowCopy.copyRows(wb, "路基边坡", "路基边坡", 5, 32, (i - 1) * 30 + 35);
            }
        }
        if(record == 1){
            wb.getSheet("路基边坡").shiftRows(34, 35, -1);
        }
        RowCopy.copyRows(wb, "source", "路基边坡", 0, 1,(record) * 30 + 3);
        wb.setPrintArea(wb.getSheetIndex("路基边坡"), 0, 8, 0,(record) * 30 + 3);
    }


    //判断要生成几个表格
    public int gettableNum(int size){
        return size%30 <= 28 ? size/30+1 : size/30+2;
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        String title = "路基边坡质量检测鉴定表";
        String sheetname = "路基边坡";
        //获取鉴定表文件
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"03路基边坡.xlsx");
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
    public void exportljbp(HttpServletResponse response)  {
        String fileName = "路基边坡实测数据";
        String sheetName = "路基边坡实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcLjgcLjbpVo()).finish();
    }


    @Override
    public void importljbp(MultipartFile file,CommonInfoVo commonInfoVo) throws IOException, ParseException {
        /*// 将文件流传过来，变成workbook对象
        XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
        //获得文本
        XSSFSheet sheet = workbook.getSheetAt(0);
        //获得行数
        int rows = sheet.getPhysicalNumberOfRows();
        //获得列数
        int columns = 0;
        for(int i=1;i<rows;i++){
            XSSFRow row = sheet.getRow(i);
            columns  = row.getPhysicalNumberOfCells();
        }
        JjgFbgcLjgcLjbpVo jjgFbgcLjgcLjbpVo = new JjgFbgcLjgcLjbpVo();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        List<String> titlelist = new ArrayList();
        for (int n=0;n<1;n++){
            for (int m=0;m<rows;m++){
                XSSFRow row = sheet.getRow(m);
                XSSFCell cell = row.getCell(n);
                titlelist.add(cell.toString());
            }
        }
        List<String> checklist = Arrays.asList("头1","头2","头3","头4","头5","头6");
        if(checklist.equals(titlelist)){
            for (int j = 1;j<columns;j++){//列
                Map<String,Object> map = new HashMap<>();
                Field[] fields = jjgFbgcLjgcLjbpVo.getClass().getDeclaredFields();
                JjgFbgcLjgcLjbp jjgFbgcLjgcLjbp = new JjgFbgcLjgcLjbp();
                for(int k=0;k<rows;k++){//行
                    //列是不变的 行增加
                    XSSFRow row = sheet.getRow(k);
                    XSSFCell cell = row.getCell(j);
                    switch (cell.getCellType()){
                        case XSSFCell.CELL_TYPE_STRING ://String
                            map.put(fields[k].getName(),cell.getStringCellValue());//属性赋值
                            break;
                        case XSSFCell.CELL_TYPE_BOOLEAN ://bealean
                            map.put(fields[k].getName(),Boolean.valueOf(cell.getBooleanCellValue()).toString());//属性赋值
                            break;
                        case XSSFCell.CELL_TYPE_NUMERIC ://number
                            //默认日期读取出来是数字，判断是否是日期格式的数字
                            if(DateUtil.isCellDateFormatted(cell)){
                                //读取的数字是日期，转换一下格式
                                DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                                Date date = cell.getDateCellValue();
                                map.put(fields[k].getName(),dateFormat.format(date));//属性赋值
                            }else {//不是日期直接赋值
                                map.put(fields[k].getName(),Double.valueOf(cell.getNumericCellValue()).toString());//属性赋值
                            }
                            break;
                        case XSSFCell.CELL_TYPE_BLANK :
                            map.put(fields[k].getName(),"");//属性赋值
                            break;
                        default:
                            System.out.println("未知类型------>"+cell);
                    }
                }
                jjgFbgcLjgcLjbp.setZh((String) map.get("zh"));
                jjgFbgcLjgcLjbp.setWz((String) map.get("wz"));
                jjgFbgcLjgcLjbp.setSjz((String) map.get("sjz"));
                jjgFbgcLjgcLjbp.setLength((String) map.get("length"));
                jjgFbgcLjgcLjbp.setHigh((String) map.get("high"));
                System.out.println(map.get("jcsj"));
                jjgFbgcLjgcLjbp.setJcsj(simpleDateFormat.parse((String) map.get("jcsj")));
                jjgFbgcLjgcLjbpMapper.insert(jjgFbgcLjgcLjbp);
            }
        }else {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }*/

        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcLjgcLjbpVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcLjgcLjbpVo>(JjgFbgcLjgcLjbpVo.class) {
                                @Override
                                public void handle(List<JjgFbgcLjgcLjbpVo> dataList) {
                                    for(JjgFbgcLjgcLjbpVo ljbpVo: dataList)
                                    {
                                        JjgFbgcLjgcLjbp jjgFbgcLjgcLjbp = new JjgFbgcLjgcLjbp();
                                        BeanUtils.copyProperties(ljbpVo,jjgFbgcLjgcLjbp);
                                        jjgFbgcLjgcLjbp.setCreatetime(new Date());
                                        jjgFbgcLjgcLjbp.setProname(commonInfoVo.getProname());
                                        jjgFbgcLjgcLjbp.setHtd(commonInfoVo.getHtd());
                                        jjgFbgcLjgcLjbp.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcLjgcLjbpMapper.insert(jjgFbgcLjgcLjbp);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }
}
