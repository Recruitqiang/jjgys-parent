package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLjgcHdjgcc;
import glgc.jjgys.model.project.JjgFbgcLjgcLjbp;
import glgc.jjgys.model.project.JjgFbgcLjgcLjcj;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.ljgc.JjgFbgcLjgcLjbpVo;
import glgc.jjgys.model.projectvo.ljgc.JjgFbgcLjgcLjcjVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcLjgcLjcjMapper;
import glgc.jjgys.system.service.JjgFbgcLjgcLjcjService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
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
import java.text.DecimalFormat;
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
public class JjgFbgcLjgcLjcjServiceImpl extends ServiceImpl<JjgFbgcLjgcLjcjMapper, JjgFbgcLjgcLjcj> implements JjgFbgcLjgcLjcjService {

    @Autowired
    private JjgFbgcLjgcLjcjMapper jjgFbgcLjgcLjcjMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    private static XSSFWorkbook wb = null;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获得数据
        QueryWrapper<JjgFbgcLjgcLjcj> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("zh","jczh");
        List<JjgFbgcLjgcLjcj> data = jjgFbgcLjgcLjcjMapper.selectList(wrapper);

        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"01路基压实度沉降.xlsx");
        if (data == null || data.size()==0){
            return;
        }else {
            //存放鉴定表的目录
            File fdir = new File(filepath+File.separator+proname+File.separator+htd);
            if(!fdir.exists()){
                //创建文件根目录
                fdir.mkdirs();
            }
            File directory = new File("src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String path =reportPath + File.separator+"路基沉降.xlsx";
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream in = new FileInputStream(f);
            wb = new XSSFWorkbook(in);
            List<String> numsList = jjgFbgcLjgcLjcjMapper.selectNums(proname,htd,fbgc);
            System.out.println(gettableNum(numsList));
            createTable(gettableNum(numsList));
            String sheetname = "路基沉降";
            if(DBtoExcel(data,proname,htd,fbgc,sheetname)){
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

    public boolean DBtoExcel(List<JjgFbgcLjgcLjcj> data,String proname,String htd,String fbgc,String sheetname) throws IOException {
        if(data.size() > 0){
            DecimalFormat nf = new DecimalFormat(".00");
            XSSFCellStyle cellstyle = JjgFbgcCommonUtils.dBtoExcelUtils(wb);
            XSSFSheet sheet = wb.getSheet(sheetname);
            String zh = data.get(0).getZh();
            String xuhao = data.get(0).getXh();
            HashMap<Integer, String> map = new HashMap<Integer, String>();
            String zhuanghao = zh+";\n";
            for(JjgFbgcLjgcLjcj row:data){
                if(xuhao.equals(row.getXh())){
                    //加上这一段空格"          "是为了表头的排版
                    if(!zhuanghao.contains(row.getZh())){
                        zhuanghao += "          "+row.getZh()+";\n";
                    }
                }
                else{
                    if(map.get(Integer.valueOf(xuhao)) == null){
                        map.put(Integer.valueOf(xuhao), zhuanghao.substring(0,zhuanghao.length()-1));
                    }
                    else{
                        map.put(Integer.valueOf(xuhao), map.get(Integer.valueOf(xuhao))+"\n"+zhuanghao.substring(0,zhuanghao.length()-1));
                    }
                    xuhao = row.getXh();
                    zhuanghao = row.getZh()+";\n";
                }
            }
            if(!"".equals(zhuanghao)){
                if(map.get(Integer.valueOf(xuhao)) == null){
                    map.put(Integer.valueOf(xuhao), zhuanghao.substring(0,zhuanghao.length()-1));
                }
                else{
                    map.put(Integer.valueOf(xuhao), map.get(Integer.valueOf(xuhao))+"\n"+zhuanghao.substring(0,zhuanghao.length()-1));
                }
            }
            int index = 0;
            int tableNum = 0;
            zh = data.get(0).getZh();
            xuhao = data.get(0).getXh();
            fillTitleCellData(sheet, tableNum,data.get(0), proname,htd,map.get(Integer.valueOf(xuhao)));
            for(JjgFbgcLjgcLjcj row:data){
                if(xuhao.equals(row.getXh())){// 序号相同 在同一表格
                    if(index == 23){//到23表示满了一页
                        calculateTotal(sheet, tableNum);
                        tableNum ++;
                        fillTitleCellData(sheet, tableNum,row,proname,htd,map.get(Integer.valueOf(xuhao)));
                        index = 0;
                    }
                    fillCommonCellData(sheet, tableNum, index, row, nf, cellstyle);
                    index ++;
                }
                else{
                    zh = row.getZh();
                    xuhao = row.getXh();
                    calculateTotal(sheet, tableNum);
                    tableNum ++;
                    index = 0;
                    fillTitleCellData(sheet, tableNum,row,proname,htd,map.get(Integer.valueOf(xuhao)));
                    fillCommonCellData(sheet, tableNum, index, row, nf, cellstyle);
                    index += 1;
                }
            }
            System.out.println("tableNum的值为：" + tableNum);
            if(tableNum >= 0){
                calculateTotal(sheet, tableNum);
            }
            /*
             * 计算沉降的总数据
             */
            sheet.getRow(6).createCell(8).setCellValue("总点数");
            sheet.getRow(6).createCell(9).setCellValue("合格点数");
            sheet.getRow(6).createCell(10).setCellValue("合格率");


            sheet.getRow(7).createCell(8).setCellFormula("SUM("
                    +sheet.getRow(31).getCell(8).getReference()+":"
                    +sheet.getRow(tableNum*33+31).getCell(8).getReference()+")");

            sheet.getRow(7).createCell(9).setCellFormula("SUM("
                    +sheet.getRow(31).getCell(9).getReference()+":"
                    +sheet.getRow(tableNum*33+31).getCell(9).getReference()+")");

            sheet.getRow(7).createCell(10).setCellFormula(
                    sheet.getRow(7).getCell(9).getReference()+"*100/"
                            +sheet.getRow(7).getCell(8).getReference());
            return true;
        }
        else{
            return false;
        }
    }

    public void fillCommonCellData(XSSFSheet sheet, int tableNum, int index, JjgFbgcLjgcLjcj row, DecimalFormat nf, XSSFCellStyle cellstyle) {
        sheet.getRow(tableNum*33+8+index%23).getCell(0).setCellValue(row.getJczh());
        sheet.getRow(tableNum*33+8+index%23).getCell(1).setCellValue(Double.valueOf(row.getNyds1()));
        sheet.getRow(tableNum*33+8+index%23).getCell(2).setCellValue(Double.valueOf(row.getNyds2()));

        sheet.getRow(tableNum*33+8+index%23).getCell(3).setCellFormula("ROUND(("
                +sheet.getRow(tableNum*33+8+index%23).getCell(2).getReference()+"-"
                +sheet.getRow(tableNum*33+8+index%23).getCell(1).getReference()+"),1)");
        sheet.getRow(tableNum*33+8+index%23).getCell(4).setCellFormula("IF("
                +sheet.getRow(tableNum*33+8+index%23).getCell(3).getReference()+"<="
                +Double.valueOf(row.getYxps())+",\"√\",\"\")");
        sheet.getRow(tableNum*33+8+index%23).getCell(5).setCellFormula("IF("
                +sheet.getRow(tableNum*33+8+index%23).getCell(3).getReference()+"<="
                +Double.valueOf(row.getYxps())+",\"\",\"×\")");
    }

    public void calculateTotal(XSSFSheet sheet, int tableNum) {
        sheet.getRow(tableNum*33+31).getCell(1).setCellFormula("COUNT("
                +sheet.getRow(tableNum*33+8).getCell(1).getReference()+":"
                +sheet.getRow(tableNum*33+30).getCell(1).getReference()+")");
        sheet.getRow(tableNum*33+31).getCell(3).setCellFormula("COUNTIF("
                +sheet.getRow(tableNum*33+8).getCell(4).getReference()+":"
                +sheet.getRow(tableNum*33+30).getCell(4).getReference()+",\"√\")");
        sheet.getRow(tableNum*33+31).getCell(6).setCellFormula(
                sheet.getRow(tableNum*33+31).getCell(3).getReference()+"*100/"
                        +sheet.getRow(tableNum*33+31).getCell(1).getReference());

        sheet.getRow(tableNum*33+31).createCell(8).setCellFormula(sheet.getRow(tableNum*33+31).getCell(1).getReference());
        sheet.getRow(tableNum*33+31).createCell(9).setCellFormula(sheet.getRow(tableNum*33+31).getCell(3).getReference());
    }

    public void fillTitleCellData(XSSFSheet sheet, int tableNum, JjgFbgcLjgcLjcj row,String proname,String htd,String position) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        sheet.getRow(tableNum * 33 + 2).getCell(1).setCellValue(proname);
        sheet.getRow(tableNum * 33 + 2).getCell(4).setCellValue(htd);
        sheet.getRow(tableNum * 33 + 3).getCell(1).setCellValue(row.getFbgc());
        sheet.getRow(tableNum*33+3).getCell(4).setCellValue(simpleDateFormat.format(row.getJcsj()));//检测日期

        sheet.getRow(tableNum*33+4).getCell(0).setCellValue("检查段落："+position);//工程部位
        sheet.getRow(tableNum*33+4).getCell(3).setCellValue("允许偏差："+Double.valueOf(row.getYxps()).intValue()+"mm");

    }


    public void createTable(int tableNum) throws IOException {
        for(int i = 1; i < tableNum; i++){
            //System.out.println("正在生成表格 -> "+i);
            RowCopy.copyRows(wb, "路基沉降", "路基沉降", 0, 32, i*33);
        }
        if(tableNum >1){
            wb.setPrintArea(wb.getSheetIndex("路基沉降"), 0, 6, 0, tableNum*33-1);
        }else
        {
            wb.setPrintArea(wb.getSheetIndex("路基沉降"), 0, 6, 0, 32);
        }

    }

    //判断要生成几个表格
    public int gettableNum(List<String> numlist){
        int tableNum = 0;
        for (String i: numlist){
            int dataNum=0;
            dataNum = Integer.parseInt(i);
            tableNum+= dataNum%23 == 0 ? dataNum/23 : dataNum/23+1;

        }
        return tableNum;
    }


    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        String title = "路基填筑沉降量鉴定表";
        String sheetname = "路基沉降";
        //获取鉴定表文件
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"01路基压实度沉降.xlsx");
        if(!f.exists()){
            return null;
        }else {
            List<Map<String,Object>> mapList = new ArrayList<>();//存放结果
            Map<String,Object> jgmap = new HashMap<>();
            DecimalFormat df = new DecimalFormat(".00");
            DecimalFormat decf = new DecimalFormat("0.##");
            //创建工作簿
            XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));
            //读取工作表
            XSSFSheet slSheet = xwb.getSheet(sheetname);
            slSheet.getRow(7).getCell(8).setCellType(CellType.STRING);
            slSheet.getRow(7).getCell(9).setCellType(CellType.STRING);
            slSheet.getRow(7).getCell(10).setCellType(CellType.STRING);
            XSSFCell bt = slSheet.getRow(0).getCell(0);
            String bt1 = bt.getStringCellValue();
            String xmname = slSheet.getRow(2).getCell(1).toString();//陕西高速
            String htdname = slSheet.getRow(2).getCell(4).toString();//LJ-1
            String hd = slSheet.getRow(3).getCell(1).toString();//涵洞
            if(slSheet != null){
                if(proname.equals(xmname) && title.equals(bt1) && htd.equals(htdname) && fbgc.equals(hd)){
                    jgmap.put("总点数",decf.format(Double.valueOf(slSheet.getRow(7).getCell(8).getStringCellValue())));
                    jgmap.put("合格点数",decf.format(Double.valueOf(slSheet.getRow(7).getCell(9).getStringCellValue())));
                    jgmap.put("合格率",df.format(Double.valueOf(slSheet.getRow(7).getCell(10).getStringCellValue())));
                    mapList.add(jgmap);
                }else {
                    return null;
                }
            }
            return mapList;
        }

    }

    @Override
    public void exportljcj(HttpServletResponse response) {
        String fileName = "路基沉降实测数据";
        String sheetName = "路基沉降实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcLjgcLjcjVo()).finish();

    }

    @Override
    public void importljcj(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcLjgcLjcjVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcLjgcLjcjVo>(JjgFbgcLjgcLjcjVo.class) {
                                @Override
                                public void handle(List<JjgFbgcLjgcLjcjVo> dataList) {
                                    for(JjgFbgcLjgcLjcjVo ljcjVo: dataList)
                                    {
                                        JjgFbgcLjgcLjcj jjgFbgcLjgcLjcj = new JjgFbgcLjgcLjcj();
                                        BeanUtils.copyProperties(ljcjVo,jjgFbgcLjgcLjcj);
                                        jjgFbgcLjgcLjcj.setCreatetime(new Date());
                                        jjgFbgcLjgcLjcj.setProname(commonInfoVo.getProname());
                                        jjgFbgcLjgcLjcj.setHtd(commonInfoVo.getHtd());
                                        jjgFbgcLjgcLjcj.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcLjgcLjcjMapper.insert(jjgFbgcLjgcLjcj);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }


    }
}
