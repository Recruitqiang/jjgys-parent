package glgc.jjgys.system.service.impl;

import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.model.project.JjgFbgcLjgcLjtsfysdHt;
import glgc.jjgys.model.project.JjgFbgcLjgcLjtsfysdSl;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.ljgc.JjgFbgcLjgcLjtsfysdHtVo;
import glgc.jjgys.model.projectvo.ljgc.JjgFbgcLjgcLjtsfysdSlVo;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcLjgcLjtsfysdSlMapper;
import glgc.jjgys.system.service.JjgFbgcLjgcLjtsfysdSlService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.Field;
import java.text.DateFormat;
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
 * @since 2023-02-21
 */
@Service
public class JjgFbgcLjgcLjtsfysdSlServiceImpl extends ServiceImpl<JjgFbgcLjgcLjtsfysdSlMapper, JjgFbgcLjgcLjtsfysdSl> implements JjgFbgcLjgcLjtsfysdSlService {


    @Autowired
    private JjgFbgcLjgcLjtsfysdSlMapper jjgFbgcLjgcLjtsfysdSlMapper;


    //private static XSSFWorkbook wb = null;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    //存评定单元需要的数据
    private LinkedHashMap<String,ArrayList<String>> evaluateData = new LinkedHashMap<>();


    @Override
    public LinkedHashMap<String, ArrayList<String>> writeAndGetData(XSSFWorkbook wb,String proname, String htd, String fbgc) {
        try {
            //获取砂砾的数据
            List<JjgFbgcLjgcLjtsfysdSl> sldata = getdata(proname, htd, fbgc);
            //获取要生成页数的数量
            List<String> numsList = jjgFbgcLjgcLjtsfysdSlMapper.selectNums(proname, htd, fbgc);
            int i = gettableSlNum(numsList);
            System.out.println(i);
            //生成表格，并设置要打印的区域
            createTable(wb,i);
            //写入数据
            DBtoExcel(wb,sldata,proname,htd,fbgc);
            JjgFbgcCommonUtils.updateFormula(wb, wb.getSheet("压实度单点(砂砾)"));
            return evaluateData;
        }catch (Exception e) {
            throw new JjgysException(20001,"处理沙砾数据错误");
        }
    }


    /**
     * 根据数据的条数，计算要生成几页
     * @param numlist
     * @return
     */
    private int gettableSlNum(List<String> numlist) {
        int tableNum = 0;
        for (String i: numlist){
            int dataNum=0;
            dataNum = Integer.parseInt(i);
            tableNum+= dataNum%4 == 0 ? dataNum/4 : dataNum/5;

        }
        return tableNum;
    }

    /** 生成表格，并设置要打印的区域
     *
     * @param num
     */
    private void createTable(XSSFWorkbook wb,int num) {
        for(int i = 1; i < num; i++){
        RowCopy.copyRows(wb, "压实度单点(砂砾)", "压实度单点(砂砾)", 0, 33, i*34);
        }
        if(num > 1){
            wb.setPrintArea(wb.getSheetIndex("压实度单点(砂砾)"), 0, 10, 0, num*34-1);
        }
    }

    /**
     * 写入沙砾sheet数据
     * @param sldata
     * @param proname
     * @param htd
     * @param fbgc
     */
    private void DBtoExcel(XSSFWorkbook wb,List<JjgFbgcLjgcLjtsfysdSl> sldata, String proname, String htd, String fbgc) {
        if(sldata.size() > 0){
            XSSFCellStyle cellstyle = wb.createCellStyle();
            XSSFFont font=wb.createFont();
            font.setFontHeightInPoints((short)9);
            font.setFontName("宋体");
            cellstyle.setFont(font);
            cellstyle.setBorderBottom(BorderStyle.THIN); //下边框
            cellstyle.setBorderLeft(BorderStyle.THIN);//左边框
            cellstyle.setBorderTop(BorderStyle.THIN);//上边框
            cellstyle.setBorderRight(BorderStyle.THIN);//右边框
            cellstyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直
            cellstyle.setAlignment(HorizontalAlignment.CENTER);//水平
            cellstyle.setWrapText(true);//自动换行
            XSSFSheet sheet = wb.getSheet("压实度单点(砂砾)");
            String xuhao = sldata.get(0).getXh();
            HashMap<Integer, String> map = new HashMap<Integer, String>();
            String zh = sldata.get(0).getZh();
            //相同的序号表示这些数据在一起评定
            for(JjgFbgcLjgcLjtsfysdSl row:sldata){
                if(xuhao.equals(row.getXh())){
                    if(!zh.equals(row.getZh())){
                        zh = row.getZh();
                    }
                }
                else{
                    if(map.get(Integer.valueOf(xuhao)) == null){
                        map.put(Integer.valueOf(xuhao), zh);
                    }
                    else{
                        map.put(Integer.valueOf(xuhao), map.get(Integer.valueOf(xuhao))+"\n"+zh);
                    }
                    xuhao = row.getXh();
                    zh = row.getZh()+";\n";
                }
            }
            if(!"".equals(zh)){
                if(map.get(Integer.valueOf(xuhao)) == null){
                    map.put(Integer.valueOf(xuhao), zh);
                }
                else{
                    map.put(Integer.valueOf(xuhao), map.get(Integer.valueOf(xuhao))+"\n"+zh);
                }
            }
            int index = 0;
            int tableNum = 0;
            zh = sldata.get(0).getZh();
            xuhao = sldata.get(0).getXh();
            String pileNO = "";
            ArrayList<String> compaction = new ArrayList<String>();
            pileNO = map.get(Integer.valueOf(xuhao));

            //填写表头
            fillTitleCellData(sheet, tableNum, sldata.get(0), zh,proname,htd);
            for(JjgFbgcLjgcLjtsfysdSl row:sldata){
                if(xuhao.equals(row.getXh())){
                    if(index == 4){
                        tableNum ++;
                        fillTitleCellData(sheet, tableNum, row, pileNO,proname,htd);
                        index = 0;
                    }
                    fillCommonCellData(wb,sheet, tableNum, index, row);
                    compaction.add(sheet.getRow(tableNum*34+32).getCell(3+index*2).getReference());
                    index ++;
                }
                else{
                    ArrayList<String> temp = new ArrayList<String>();
                    temp = (ArrayList<String>) compaction.clone();
                    evaluateData.put(pileNO+"@"+xuhao, temp);
                    compaction.clear();
                    xuhao = row.getXh();
                    pileNO = map.get(Integer.valueOf(xuhao));
                    tableNum ++;
                    index = 0;
                    fillTitleCellData(sheet, tableNum, row, pileNO,proname,htd);
                    fillCommonCellData(wb,sheet, tableNum, index, row);
                    compaction.add(sheet.getRow(tableNum*34+32).getCell(3+index*2).getReference());
                    index = 1;
                }
            }
            evaluateData.put(pileNO+"@"+xuhao, compaction);
        }
    }

    /**
     * 填写中间下方的普通单元格
     * @param sheet
     * @param tableNum
     * @param index
     * @param row
     */
    public void fillCommonCellData(XSSFWorkbook wb,XSSFSheet sheet, int tableNum, int index, JjgFbgcLjgcLjtsfysdSl row){
        sheet.getRow(tableNum*34+7).getCell(3+index*2).setCellValue(row.getQyzhjwz());//E8 K58+120/左距中6.0m
        sheet.getRow(tableNum*34+8).getCell(3+index*2).setCellValue(row.getSksd());//E9 试坑深度
        sheet.getRow(tableNum*34+9).getCell(3+index*2).setCellValue(row.getGsqtszl());//E10 灌砂前筒+砂重
        sheet.getRow(tableNum*34+10).getCell(3+index*2).setCellValue(row.getGshtszl());//E11 灌砂后筒+砂重
        sheet.getRow(tableNum*34+11).getCell(3+index*2).setCellValue(row.getZtsz());//E12 锥体砂重

        sheet.getRow(tableNum*34+12).getCell(3+index*2).setCellFormula(
                sheet.getRow(tableNum*34+9).getCell(3+index*2).getReference()+"-"
                        +sheet.getRow(tableNum*34+10).getCell(3+index*2).getReference()+"-"
                        +sheet.getRow(tableNum*34+11).getCell(3+index*2).getReference());//E13=E10-E11-E12

        sheet.getRow(tableNum*34+13).getCell(3+index*2).setCellValue(row.getLsdmd());//E14 量砂的密度ρs

        sheet.getRow(tableNum*34+14).getCell(3+index*2).setCellFormula("IF("
                +sheet.getRow(tableNum*34+13).getCell(3+index*2).getReference()+"=0,\"\","
                +sheet.getRow(tableNum*34+12).getCell(3+index*2).getReference()+"/"
                +sheet.getRow(tableNum*34+13).getCell(3+index*2).getReference()+")");//E15=IF(E14=0," ",E13/E14)

        sheet.getRow(tableNum*34+15).getCell(3+index*2).setCellValue(row.getHhldszl());//E16混合料的湿质量

        sheet.getRow(tableNum*34+16).getCell(3+index*2).setCellFormula("IF("
                +sheet.getRow(tableNum*34+14).getCell(3+index*2).getReference()+"=0,\"\","
                +sheet.getRow(tableNum*34+15).getCell(3+index*2).getReference()+"/"
                +sheet.getRow(tableNum*34+14).getCell(3+index*2).getReference()+")");//E17=IF(E15=0," ",E16/E15)

        sheet.getRow(tableNum*34+20).getCell(3+index*2).setCellValue(row.getHzl());//E21 盒质量

        sheet.getRow(tableNum*34+18).getCell(3+index*2).setCellFormula(
                sheet.getRow(tableNum*34+15).getCell(3+index*2).getReference()+"+"
                        +sheet.getRow(tableNum*34+20).getCell(3+index*2).getReference());//E19=E16+E21

        sheet.getRow(tableNum*34+19).getCell(3+index*2).setCellValue(row.getHgzl());//E20 盒+干质量
        sheet.getRow(tableNum*34+19-2).getCell(3+index*2).setCellValue(row.getHh());//盒号

        sheet.getRow(tableNum*34+21).getCell(3+index*2).setCellFormula(
                sheet.getRow(tableNum*34+18).getCell(3+index*2).getReference()+"-"
                        +sheet.getRow(tableNum*34+19).getCell(3+index*2).getReference());//E22=E19-E20

        sheet.getRow(tableNum*34+22).getCell(3+index*2).setCellFormula(
                sheet.getRow(tableNum*34+19).getCell(3+index*2).getReference()+"-"
                        +sheet.getRow(tableNum*34+20).getCell(3+index*2).getReference());//E23=E20-E21

        sheet.getRow(tableNum*34+23).getCell(3+index*2).setCellFormula("IF("
                +sheet.getRow(tableNum*34+22).getCell(3+index*2).getReference()+"=0,\"\","
                +sheet.getRow(tableNum*34+21).getCell(3+index*2).getReference()+"/"
                +sheet.getRow(tableNum*34+22).getCell(3+index*2).getReference()+"*100)");//E24=IF(E23=0," ",E22/E23*100)

        sheet.getRow(tableNum*34+24).getCell(3+index*2).setCellValue(row.getKlzl());//E25 5-38mm颗粒质量

        sheet.getRow(tableNum*34+26).getCell(3+index*2).setCellFormula("IF("
                +sheet.getRow(tableNum*34+22).getCell(3+index*2).getReference()+"=0,\"\","
                +sheet.getRow(tableNum*34+24).getCell(3+index*2).getReference()+"/"
                +sheet.getRow(tableNum*34+22).getCell(3+index*2).getReference()+"*100)");//E27=IF(E23=0," ",E25/E23*100)

        sheet.getRow(tableNum*34+28).getCell(3+index*2).setCellFormula("IF("
                +sheet.getRow(tableNum*34+23).getCell(3+index*2).getReference()+"=0,\"\","
                +sheet.getRow(tableNum*34+16).getCell(3+index*2).getReference()+"/(1+"
                +sheet.getRow(tableNum*34+23).getCell(3+index*2).getReference()+"/"
                +100+"))");//E29=IF(E24=0," ",E17/(1+E24/100))
        sheet.getRow(tableNum*34+30).getCell(3+index*2).setCellFormula("$"
                +sheet.getRow(tableNum*34+6).getCell(8).getReference()+"*"
                +sheet.getRow(tableNum*34+26).getCell(3+index*2).getReference()+"+$"
                +sheet.getRow(tableNum*34+6).getCell(10).getReference());//E31=$J7*E27+$L7

        sheet.getRow(tableNum*34+32).getCell(3+index*2).setCellFormula("IF(ISERROR("
                +sheet.getRow(tableNum*34+28).getCell(3+index*2).getReference()+"/"
                +sheet.getRow(tableNum*34+30).getCell(3+index*2).getReference()+"*100),\"\","
                +sheet.getRow(tableNum*34+28).getCell(3+index*2).getReference()+"/"
                +sheet.getRow(tableNum*34+30).getCell(3+index*2).getReference()+"*100)");//E33=IF(ISERROR(E29/E31*100),"",E29/E31*100)

        XSSFCellStyle cellstyle2 = wb.createCellStyle();
        XSSFFont font=wb.createFont();
        font.setFontHeightInPoints((short)9);
        font.setColor(Font.COLOR_RED);
        font.setFontName("宋体");
        cellstyle2.setFont(font);
        cellstyle2.setBorderBottom(BorderStyle.THIN); //下边框
        cellstyle2.setBorderLeft(BorderStyle.THIN);//左边框
        cellstyle2.setBorderTop(BorderStyle.THIN);//上边框
        cellstyle2.setBorderRight(BorderStyle.THIN);//右边框
        cellstyle2.setVerticalAlignment(VerticalAlignment.CENTER);//垂直
        cellstyle2.setAlignment(HorizontalAlignment.CENTER);//水平
        cellstyle2.setWrapText(true);//自动换行
        sheet.getRow(tableNum*34+33).getCell(3+index*2).setCellFormula("IF("
                +sheet.getRow(tableNum*34+32).getCell(3+index*2).getReference()+">=94,\"√\",\"×\")");//E34=IF(E33>=94,"√","×")
        XSSFFormulaEvaluator evaluate = new XSSFFormulaEvaluator(wb);
        if("×".equals(sheet.getRow(tableNum*34+33).getCell(3+index*2).getRawValue())){
            sheet.getRow(tableNum*34+33).getCell(3+index*2).setCellStyle(cellstyle2);
        }
    }

    /**
     * 填写表头
     * @param sheet
     * @param tableNum
     * @param row
     * @param position
     * @param projectname
     * @param htd
     */
    public void fillTitleCellData(XSSFSheet sheet, int tableNum, JjgFbgcLjgcLjtsfysdSl row, String position,String projectname,String htd){
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        sheet.getRow(tableNum*34+1).getCell(2).setCellValue(projectname);
        sheet.getRow(tableNum*34+1).getCell(8).setCellValue(htd);
        sheet.getRow(tableNum*34+2).getCell(2).setCellValue("路基土石方");
        sheet.getRow(tableNum*34+2).getCell(8).setCellValue(simpleDateFormat.format(row.getSysj()));//实验时间
        sheet.getRow(tableNum*34+3).getCell(2).setCellValue(position.substring(0, position.length()-1));//桩号
        DecimalFormat standard = new DecimalFormat("0");
        sheet.getRow(tableNum*34+5).getCell(2).setCellValue(standard.format(Double.parseDouble(row.getSlgdz())));//C6压实度标准
        sheet.getRow(tableNum*34+6).getCell(6).setCellValue(Double.parseDouble(row.getA()));//H7 a
        sheet.getRow(tableNum*34+6).getCell(8).setCellValue(Double.parseDouble(row.getB()));//J7 b
        sheet.getRow(tableNum*34+6).getCell(10).setCellValue(Double.parseDouble(row.getC()));//L7 c
    }

    /**
     * 获取需要的数据
     * @param proname
     * @param htd
     * @param fbgc
     * @return
     */
    public List<JjgFbgcLjgcLjtsfysdSl> getdata(String proname, String htd, String fbgc){
        QueryWrapper<JjgFbgcLjgcLjtsfysdSl> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("xh","qyzhjwz");
        List<JjgFbgcLjgcLjtsfysdSl> data = jjgFbgcLjgcLjtsfysdSlMapper.selectList(wrapper);
        return data;
    }


    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) {
        return null;
    }

    /**
     * 导出模板文件
     * @param response
     * @throws IOException
     */
    @Override
    public void exportysdsl(HttpServletResponse response) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();// 创建一个Excel文件
        XSSFCellStyle columnHeadStyle = JjgFbgcCommonUtils.tsfCellStyle(workbook);
        XSSFSheet sheet = workbook.createSheet("实测数据");// 创建一个Excel的Sheet
        sheet.setColumnWidth(0,28*256);
        List<String> checklist = Arrays.asList("路基压实度_砂砾规定值","试验时间","桩号","a",
                "b","c","取样桩号及位置","试坑深度(cm)", "灌砂前筒+砂质量（g）","灌砂后筒+砂质量（g）",
                "锥体砂重","量砂的密度ρs","混合料的湿质量","盒号","盒+干质量","盒质量","5-38mm颗粒质量(g)","序号");
        for (int i=0;i< checklist.size();i++){
            XSSFRow row = sheet.createRow(i);// 创建第一行
            XSSFCell cell = row.createCell(0);// 创建第一行第一列
            cell.setCellValue(new XSSFRichTextString(checklist.get(i)));
            cell.setCellStyle(columnHeadStyle);
        }
        String filename = "路基土石方压实度_砂砾实测数据.xls";// 设置下载时客户端Excel的名称
        filename = new String((filename).getBytes("GBK"), "ISO8859_1");
        response.setContentType("application/octet-stream;charset=UTF-8");
        response.addHeader("Content-Disposition", "attachment;filename=" + filename);
        OutputStream ouputStream = response.getOutputStream();
        workbook.write(ouputStream);
        ouputStream.flush();
        ouputStream.close();

    }

    /**
     * 导入数据
     * @param file
     * @param commonInfoVo
     * @throws IOException
     * @throws ParseException
     */
    @Override
    public void importysdsl(MultipartFile file, CommonInfoVo commonInfoVo) throws IOException, ParseException {
        // 将文件流传过来，变成workbook对象
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
        JjgFbgcLjgcLjtsfysdSlVo jjgFbgcLjgcLjtsfysdSlVo = new JjgFbgcLjgcLjtsfysdSlVo();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        List<String> titlelist = new ArrayList();
        for (int n=0;n<1;n++){
            for (int m=0;m<rows;m++){
                XSSFRow row = sheet.getRow(m);
                XSSFCell cell = row.getCell(n);
                titlelist.add(cell.toString());
            }
        }
        List<String> checklist = Arrays.asList("路基压实度_砂砾规定值","试验时间","桩号","a",
                "b","c","取样桩号及位置","试坑深度(cm)", "灌砂前筒+砂质量（g）","灌砂后筒+砂质量（g）",
                "锥体砂重","量砂的密度ρs","混合料的湿质量","盒号","盒+干质量","盒质量","5-38mm颗粒质量(g)","序号");
        if(checklist.equals(titlelist)){
            for (int j = 1;j<columns;j++){//列
                Map<String,Object> map = new HashMap<>();
                Field[] fields = jjgFbgcLjgcLjtsfysdSlVo.getClass().getDeclaredFields();
                JjgFbgcLjgcLjtsfysdSl jjgFbgcLjgcLjtsfysdSl = new JjgFbgcLjgcLjtsfysdSl();
                for(int k=0;k<rows;k++){//行
                    //列是不变的 行增加
                    XSSFRow row = sheet.getRow(k);
                    XSSFCell cell = row.getCell(j);
                    switch (cell.getCellType()){
                        case XSSFCell.CELL_TYPE_STRING ://String
                            cell.setCellType(CellType.STRING);
                            map.put(fields[k].getName(),cell.getStringCellValue());//属性赋值
                            break;
                        case XSSFCell.CELL_TYPE_BOOLEAN ://bealean
                            cell.setCellType(CellType.STRING);
                            //map.put(fields[k].getName(),Boolean.valueOf(cell.getBooleanCellValue()).toString());//属性赋值
                            map.put(fields[k].getName(),String.valueOf(cell.getStringCellValue()));//属性赋值
                            break;
                        case XSSFCell.CELL_TYPE_NUMERIC ://number
                            //默认日期读取出来是数字，判断是否是日期格式的数字
                            if(DateUtil.isCellDateFormatted(cell)){
                                //读取的数字是日期，转换一下格式
                                DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                                Date date = cell.getDateCellValue();
                                map.put(fields[k].getName(),dateFormat.format(date));//属性赋值
                            }else {//不是日期直接赋值
                                cell.setCellType(CellType.STRING);
                                //map.put(fields[k].getName(),Double.valueOf(cell.getNumericCellValue()).toString());//属性赋值
                                map.put(fields[k].getName(),String.valueOf(cell.getStringCellValue()));//属性赋值
                            }
                            break;
                        case XSSFCell.CELL_TYPE_BLANK :
                            cell.setCellType(CellType.STRING);
                            map.put(fields[k].getName(),"");//属性赋值
                            break;
                        default:
                            System.out.println("未知类型------>"+cell);
                    }
                }
                jjgFbgcLjgcLjtsfysdSl.setSlgdz((String) map.get("slgdz"));
                jjgFbgcLjgcLjtsfysdSl.setSysj(simpleDateFormat.parse((String) map.get("sysj")));
                jjgFbgcLjgcLjtsfysdSl.setZh((String) map.get("zh"));
                jjgFbgcLjgcLjtsfysdSl.setA((String) map.get("a"));
                jjgFbgcLjgcLjtsfysdSl.setB((String) map.get("b"));
                jjgFbgcLjgcLjtsfysdSl.setC((String) map.get("c"));
                jjgFbgcLjgcLjtsfysdSl.setQyzhjwz((String) map.get("qyzhjwz"));
                jjgFbgcLjgcLjtsfysdSl.setSksd((String) map.get("sksd"));
                jjgFbgcLjgcLjtsfysdSl.setGsqtszl((String) map.get("gsqtszl"));
                jjgFbgcLjgcLjtsfysdSl.setGshtszl((String) map.get("gshtszl"));
                jjgFbgcLjgcLjtsfysdSl.setZtsz((String) map.get("ztsz"));
                jjgFbgcLjgcLjtsfysdSl.setLsdmd((String) map.get("lsdmd"));
                jjgFbgcLjgcLjtsfysdSl.setHhldszl((String) map.get("hhldszl"));
                jjgFbgcLjgcLjtsfysdSl.setHh((String) map.get("hh"));
                jjgFbgcLjgcLjtsfysdSl.setHgzl((String) map.get("hgzl"));
                jjgFbgcLjgcLjtsfysdSl.setHzl((String) map.get("hzl"));
                jjgFbgcLjgcLjtsfysdSl.setKlzl((String) map.get("klzl"));
                jjgFbgcLjgcLjtsfysdSl.setXh((String) map.get("xh"));
                jjgFbgcLjgcLjtsfysdSl.setProname(commonInfoVo.getProname());
                jjgFbgcLjgcLjtsfysdSl.setHtd(commonInfoVo.getHtd());
                jjgFbgcLjgcLjtsfysdSl.setFbgc(commonInfoVo.getFbgc());
                jjgFbgcLjgcLjtsfysdSl.setCreatetime(new Date());
                jjgFbgcLjgcLjtsfysdSlMapper.insert(jjgFbgcLjgcLjtsfysdSl);
            }
        }else {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }
}
