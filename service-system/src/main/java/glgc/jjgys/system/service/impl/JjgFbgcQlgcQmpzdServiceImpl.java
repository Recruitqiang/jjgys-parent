package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcQlgcQmgzsd;
import glgc.jjgys.model.project.JjgFbgcQlgcQmpzd;
import glgc.jjgys.model.project.JjgFbgcSdgcLmgzsdsgpsf;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.qlgc.JjgFbgcQlgcQmpzdVo;
import glgc.jjgys.model.projectvo.sdgc.JjgFbgcSdgcLmgzsdsgpsfVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcQlgcQmpzdMapper;
import glgc.jjgys.system.service.JjgFbgcQlgcQmpzdService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.poi.ss.usermodel.CellType;
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
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-03-20
 */
@Service
public class JjgFbgcQlgcQmpzdServiceImpl extends ServiceImpl<JjgFbgcQlgcQmpzdMapper, JjgFbgcQlgcQmpzd> implements JjgFbgcQlgcQmpzdService {

    @Autowired
    private JjgFbgcQlgcQmpzdMapper jjgFbgcQlgcQmpzdMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        List<Map<String,Object>> qlmclist = jjgFbgcQlgcQmpzdMapper.selectqlmc(proname,htd,fbgc);
        if (qlmclist.size()>0){
            for (Map<String, Object> m : qlmclist)
            {
                for (String k : m.keySet()){
                    String qlmc = m.get(k).toString();
                    DBtoExcelql(proname,htd,fbgc,qlmc,qlmclist.size());
                }
            }
        }

    }

    /**
     *
     * @param proname
     * @param htd
     * @param fbgc
     * @param qlmc
     * @param qlsize
     * @throws IOException
     */
    private void DBtoExcelql(String proname, String htd, String fbgc, String qlmc, int qlsize) throws IOException {
        XSSFWorkbook wb = null;
        QueryWrapper<JjgFbgcQlgcQmpzd> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.like("qm",qlmc);
        wrapper.orderByAsc("zh");
        List<JjgFbgcQlgcQmpzd> data = jjgFbgcQlgcQmpzdMapper.selectList(wrapper);
        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"34桥面平整度3米直尺法-"+qlmc+".xlsx");
        if (data == null || data.size()==0){
            return;
        }else {
            File fdir = new File(filepath + File.separator + proname + File.separator + htd);
            if (!fdir.exists()) {
                //创建文件根目录
                fdir.mkdirs();
            }
            File directory = new File("src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String path = reportPath + File.separator + "平整度3米直尺法.xlsx";
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            createTable(gettableNum(qlsize),wb);
            if(DBtoExcel(data,wb,qlmc)){
                calculateTheSheet(wb.getSheet("平整度"));

                for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                    JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
                }

                JjgFbgcCommonUtils.deleteEmptySheets(wb);
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
    private void calculateTheSheet(XSSFSheet sheet) {
        XSSFRow row = null;
        boolean flag = false;
        String name = "";
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        //System.out.println("sheet.getPhysicalNumberOfRows() = "+ sheet.getPhysicalNumberOfRows());
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if(flag && !"".equals(row.getCell(0).toString()) && name.equals(row.getCell(0).toString())){
                rowend = sheet.getRow(i+8);
                calculateConclution(sheet, i);
                i+=8;
            }
            if(flag && !"".equals(row.getCell(0).toString()) && !name.equals(row.getCell(0).toString())){
                rowstart.createCell(6).setCellFormula("COUNT("
                        +rowstart.getCell(3).getReference()+":"
                        +rowend.getCell(3).getReference()+")");//G=COUNT(D6:D59)总点数
                rowstart.createCell(7).setCellFormula("COUNTIF("
                        +rowstart.getCell(4).getReference()+":"
                        +rowend.getCell(4).getReference()+",\"√\")");//H=COUNTIF(E6:E59,"√")合格点数
                rowstart.createCell(8).setCellFormula(
                        rowstart.getCell(7).getReference()+"/"
                                +rowstart.getCell(6).getReference()+"*100");//合格率
                rowstart = sheet.getRow(i);
                name = rowstart.getCell(0).toString();
                rowend = sheet.getRow(i+8);
                calculateConclution(sheet, i);
                i+=8;
            }
            if ("桩号".equals(row.getCell(0).toString())) {
                flag = true;
                i += 1;
                rowstart = sheet.getRow(i + 1);
                rowend = sheet.getRow(i+9);
                name = rowstart.getCell(0).toString();
                calculateConclution(sheet, i+1);
                i+=9;
            }
        }
        rowend = sheet.getRow(sheet.getPhysicalNumberOfRows()-1);
        rowstart.createCell(6).setCellFormula("COUNT("
                +rowstart.getCell(3).getReference()+":"
                +rowend.getCell(3).getReference()+")");//G=COUNT(D6:D59)总点数
        rowstart.createCell(7).setCellFormula("COUNTIF("
                +rowstart.getCell(4).getReference()+":"
                +rowend.getCell(4).getReference()+",\"√\")");//H=COUNTIF(E6:E59,"√")合格点数
        rowstart.createCell(8).setCellFormula(
                rowstart.getCell(7).getReference()+"/"
                        +rowstart.getCell(6).getReference()+"*100");//合格率

        sheet.getRow(2).getCell(7).setCellFormula("COUNT("
                +sheet.getRow(5).getCell(3).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).getCell(3).getReference()+")");//G=COUNT(D6:D59)总点数
        sheet.getRow(3).getCell(7).setCellFormula("COUNTIF("
                +sheet.getRow(5).getCell(4).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).getCell(4).getReference()+",\"√\")");//H=COUNTIF(E6:E59,"√")合格点数
        sheet.getRow(4).getCell(7).setCellFormula(
                sheet.getRow(3).getCell(7).getReference()+"/"
                        +sheet.getRow(2).getCell(7).getReference()+"*100");//合格率

    }

    /**
     *
     * @param sheet
     * @param index
     */
    private void calculateConclution(XSSFSheet sheet, int index) {
        for(int i=0; i<9;i++){
            sheet.getRow(index+i).getCell(4).setCellFormula("IF(AND("
                    +sheet.getRow(index+i).getCell(3).getReference()+"<="
                    +sheet.getRow(index+i).getCell(2).getReference()+","
                    +sheet.getRow(index+i).getCell(2).getReference()+">0),\"√\",\"\")");//E6=IF(AND(D6<=C6,C6>0),"√","")
            sheet.getRow(index+i).getCell(5).setCellFormula("IF(AND("
                    +sheet.getRow(index+i).getCell(3).getReference()+">"
                    +sheet.getRow(index+i).getCell(2).getReference()+","
                    +sheet.getRow(index+i).getCell(2).getReference()+">0),\"×\",\"\")");//F6=IF(AND(D6>C6,C6>0),"×","")
        }
    }

    /**
     * 写入数据
     * @param data
     * @param wb
     * @param qlmc
     * @return
     */
    private boolean DBtoExcel(List<JjgFbgcQlgcQmpzd> data, XSSFWorkbook wb,String qlmc) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        XSSFSheet sheet = wb.getSheet("平整度");
        sheet.getRow(1).getCell(1).setCellValue(data.get(0).getProname());
        sheet.getRow(1).getCell(4).setCellValue(data.get(0).getHtd());
        sheet.getRow(2).getCell(1).setCellValue(qlmc);
        sheet.getRow(2).getCell(4).setCellValue(simpleDateFormat.format(data.get(0).getJcsj()));
        int index = 5;
        for(int i=0;i<data.size();i++){
            sheet.getRow(index).getCell(0).setCellValue(data.get(i).getQm()+data.get(i).getZh());
            sheet.getRow(index).getCell(1).setCellValue(data.get(i).getWz());
            sheet.getRow(index).getCell(2).setCellValue(Integer.parseInt(data.get(i).getSjz()));
            sheet.getRow(index+1).getCell(2).setCellValue(Integer.parseInt(data.get(i+1).getSjz()));
            sheet.getRow(index+2).getCell(2).setCellValue(Integer.parseInt(data.get(i+2).getSjz()));
            sheet.getRow(index).getCell(3).setCellValue(Double.parseDouble(data.get(i).getScz()));
            sheet.getRow(index+1).getCell(3).setCellValue(Double.parseDouble(data.get(i+1).getScz()));
            sheet.getRow(index+2).getCell(3).setCellValue(Double.parseDouble(data.get(i+2).getScz()));

            sheet.getRow(index+3).getCell(1).setCellValue(data.get(i+3).getWz());
            sheet.getRow(index+3).getCell(2).setCellValue(Integer.parseInt(data.get(i+3).getSjz()));
            sheet.getRow(index+1+3).getCell(2).setCellValue(Integer.parseInt(data.get(i+1+3).getSjz()));
            sheet.getRow(index+2+3).getCell(2).setCellValue(Integer.parseInt(data.get(i+2+3).getSjz()));
            sheet.getRow(index+3).getCell(3).setCellValue(Double.parseDouble(data.get(i+3).getScz()));
            sheet.getRow(index+1+3).getCell(3).setCellValue(Double.parseDouble(data.get(i+1+3).getScz()));
            sheet.getRow(index+2+3).getCell(3).setCellValue(Double.parseDouble(data.get(i+2+3).getScz()));

            sheet.getRow(index+6).getCell(1).setCellValue(data.get(i+6).getWz());
            sheet.getRow(index+6).getCell(2).setCellValue(Integer.parseInt(data.get(i+6).getSjz()));
            sheet.getRow(index+1+6).getCell(2).setCellValue(Integer.parseInt(data.get(i+1+6).getSjz()));
            sheet.getRow(index+2+6).getCell(2).setCellValue(Integer.parseInt(data.get(i+2+6).getSjz()));
            sheet.getRow(index+6).getCell(3).setCellValue(Double.parseDouble(data.get(i+6).getScz()));
            sheet.getRow(index+1+6).getCell(3).setCellValue(Double.parseDouble(data.get(i+1+6).getScz()));
            sheet.getRow(index+2+6).getCell(3).setCellValue(Double.parseDouble(data.get(i+2+6).getScz()));

            index += 9;
            i += 8;
        }
        return true;
    }

    /**
     *
     * @param tableNum
     * @param wb
     */
    private void createTable(int tableNum, XSSFWorkbook wb) {
        for(int i = 1; i < tableNum; i++){
            RowCopy.copyRows(wb, "平整度", "平整度", 5, 31, i*32);
        }
        if(tableNum >= 1){
            wb.setPrintArea(wb.getSheetIndex("平整度"), 0, 5, 0, tableNum*27+4);
        }
    }

    /**
     *
     * @param size
     * @return
     */
    private int gettableNum(int size) {
        return size%3 == 0 ? size/3 : size/3+1;
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        String title = "平整度质量鉴定表";
        String sheetname = "平整度";

        List<Map<String, Object>> mapList = new ArrayList<>();

        List<Map<String,Object>> qlmclist = jjgFbgcQlgcQmpzdMapper.selectqlmc(proname,htd,fbgc);
        if (qlmclist.size()>0){
            for (Map<String, Object> m : qlmclist) {
                for (String k : m.keySet()){
                    String qlmc = m.get(k).toString();
                    Map<String, Object> lookqljdb = lookqljdb(proname, htd, qlmc, sheetname,title);
                    mapList.add(lookqljdb);
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
     * @param qlmc
     * @param sheetname
     * @param title
     * @return
     * @throws IOException
     */
    private Map<String, Object> lookqljdb(String proname, String htd, String qlmc, String sheetname, String title) throws IOException {
        DecimalFormat df = new DecimalFormat(".00");
        DecimalFormat decf = new DecimalFormat("0.##");
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "34桥面平整度3米直尺法-"+qlmc+".xlsx");
        if (!f.exists()) {
            return null;
        } else {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
            XSSFSheet slSheet = wb.getSheet(sheetname);
            XSSFCell bt = slSheet.getRow(0).getCell(0);//标题
            XSSFCell xmname = slSheet.getRow(1).getCell(1);//项目名
            XSSFCell htdname = slSheet.getRow(1).getCell(4);//合同段名
            Map<String, Object> jgmap = new HashMap<>();
            if (proname.equals(xmname.toString()) && title.equals(bt.toString()) && htd.equals(htdname.toString())) {

                slSheet.getRow(2).getCell(7).setCellType(CellType.STRING);//总点数
                slSheet.getRow(3).getCell(7).setCellType(CellType.STRING);//合格点数
                slSheet.getRow(4).getCell(7).setCellType(CellType.STRING);//合格率
                double zds = Double.valueOf(slSheet.getRow(2).getCell(7).getStringCellValue());
                double hgds = Double.valueOf(slSheet.getRow(3).getCell(7).getStringCellValue());
                double hgl = Double.valueOf(slSheet.getRow(4).getCell(7).getStringCellValue());
                String zdsz1 = decf.format(zds);
                String hgdsz1 = decf.format(hgds);
                String hglz1 = df.format(hgl);
                jgmap.put("检测项目", qlmc);
                jgmap.put("检测点数", zdsz1);
                jgmap.put("合格点数", hgdsz1);
                jgmap.put("合格率", hglz1);
            }
            return jgmap;
        }
    }

    @Override
    public void exportQmpzd(HttpServletResponse response) {
        String fileName = "09桥面平整度三米直尺法实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcQlgcQmpzdVo()).finish();

    }

    @Override
    public void importQmpzd(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcQlgcQmpzdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcQlgcQmpzdVo>(JjgFbgcQlgcQmpzdVo.class) {
                                @Override
                                public void handle(List<JjgFbgcQlgcQmpzdVo> dataList) {
                                    for(JjgFbgcQlgcQmpzdVo qlgcQmpzdVo: dataList)
                                    {
                                        JjgFbgcQlgcQmpzd qlgcQmpzd = new JjgFbgcQlgcQmpzd();
                                        BeanUtils.copyProperties(qlgcQmpzdVo,qlgcQmpzd);
                                        qlgcQmpzd.setCreatetime(new Date());
                                        qlgcQmpzd.setProname(commonInfoVo.getProname());
                                        qlgcQmpzd.setHtd(commonInfoVo.getHtd());
                                        qlgcQmpzd.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcQlgcQmpzdMapper.insert(qlgcQmpzd);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }
}
