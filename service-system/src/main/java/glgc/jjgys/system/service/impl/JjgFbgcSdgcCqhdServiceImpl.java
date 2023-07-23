package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcSdgcCqhd;
import glgc.jjgys.model.project.JjgFbgcSdgcCqtqd;
import glgc.jjgys.model.project.JjgFbgcSdgcSdhp;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.sdgc.JjgFbgcSdgcCqhdVo;
import glgc.jjgys.model.projectvo.sdgc.JjgFbgcSdgcCqtqdVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcSdgcCqhdMapper;
import glgc.jjgys.system.service.JjgFbgcSdgcCqhdService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.commons.lang3.StringUtils;
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
 * @since 2023-03-27
 */
@Service
public class JjgFbgcSdgcCqhdServiceImpl extends ServiceImpl<JjgFbgcSdgcCqhdMapper, JjgFbgcSdgcCqhd> implements JjgFbgcSdgcCqhdService {

    @Autowired
    private JjgFbgcSdgcCqhdMapper jjgFbgcSdgcCqhdMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;


    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        List<Map<String,Object>> sdmclist = jjgFbgcSdgcCqhdMapper.selectsdmc(proname,htd,fbgc);
        if (sdmclist.size()>0){
            for (Map<String, Object> m : sdmclist)
            {
                for (String k : m.keySet()){
                    String sdmc = m.get(k).toString();
                    List<Map<String,Object>> cd = jjgFbgcSdgcCqhdMapper.selectcd(proname,htd,fbgc,sdmc);
                    if (cd.get(0).get("zgy1") != null && !"".equals(cd.get(0).get("zgy1")) && cd.get(0).get("zgy2") !=null && !"".equals(cd.get(0).get("zgy2")) && cd.get(0).get("zgy3") !=null && !"".equals(cd.get(0).get("zgy3"))){
                        DBtoExcelsd(4,proname,htd,fbgc,sdmc);
                    }else if (cd.get(0).get("zgy1") != null && !"".equals(cd.get(0).get("zgy1")) && cd.get(0).get("zgy2") !=null && !"".equals(cd.get(0).get("zgy2"))){
                        DBtoExcelsd(3,proname,htd,fbgc,sdmc);
                    }else if (cd.get(0).get("zgy1") != null && !"".equals(cd.get(0).get("zgy1"))){
                        DBtoExcelsd(2,proname,htd,fbgc,sdmc);
                    }

                }
            }
        }
    }

    private void DBtoExcelsd(int cd,String proname, String htd, String fbgc, String sdmc) throws IOException {
        XSSFWorkbook wb = null;
        QueryWrapper<JjgFbgcSdgcCqhd> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.like("sdmc",sdmc);
        wrapper.orderByAsc("sdmc","zh");
        List<JjgFbgcSdgcCqhd> data = jjgFbgcSdgcCqhdMapper.selectList(wrapper);
        String sheetname = cd+"车道"+data.get(0).getWz();
        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"39隧道衬砌厚度-"+sdmc+".xlsx");
        //健壮性判断如果没有数据返回"请导入数据"
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
            String name = "隧道衬砌厚度.xlsx";
            String path = reportPath + File.separator + name;
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);

            if (cd == 2){
                createTable2(gettableNum(data.size()),wb,sheetname);
                DBtoExcel2(data,wb,sheetname,sdmc);

            }else if (cd ==3) {
                createTable3(gettableNum(data.size()),wb,sheetname);
                DBtoExcel3(data,wb,sheetname,sdmc);

            }else if (cd ==4){
                createTable4(gettable4(data.size()),wb,sheetname);
                DBtoExcel4(data,wb,sheetname,sdmc);

            }
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
            }
            JjgFbgcCommonUtils.deleteEmptySheets(wb);
            FileOutputStream fileOut = new FileOutputStream(f);
            wb.write(fileOut);
            fileOut.flush();
            fileOut.close();
            out.close();
            wb.close();
        }

    }



    private int gettableNum(int size) {
        return size%31 ==0 ? size/31 : size/31+1;
    }

    private void DBtoExcel4(List<JjgFbgcSdgcCqhd> data, XSSFWorkbook wb, String sheetname, String sdmc) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        if (data.size()>0){
            XSSFSheet sheet = wb.getSheet(sheetname);
            XSSFRow titlerow = sheet.getRow(1);
            titlerow.getCell(0).setCellValue("项目名称："+data.get(0).getProname());
            titlerow.getCell(18).setCellValue("合同段："+data.get(0).getHtd());
            titlerow = sheet.getRow(2);
            titlerow.getCell(0).setCellValue("单位工程名称："+sdmc);
            titlerow = sheet.getRow(3);
            titlerow.getCell(18).setCellValue("检测时间："+simpleDateFormat.format(data.get(0).getJcsj()));

            //填写数据
            int curRow = 7;
            for(JjgFbgcSdgcCqhd row : data){
                titlerow = sheet.getRow(curRow);
                //桩号
                titlerow.getCell(0).setCellValue(row.getZh());
                //位置
                titlerow.getCell(1).setCellValue(row.getWz());
                //设计厚度
                titlerow.getCell(2).setCellValue(Double.valueOf(String.format("%.0f", Double.valueOf(row.getSjhd()))).intValue());

                //左拱腰1
                titlerow.getCell(3).setCellValue(Double.valueOf(String.format("%.0f", Double.valueOf(row.getZgy1()))).intValue());
                titlerow.getCell(4).setCellFormula("IF("+titlerow.getCell(3).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+">=0,\"√\",\"\")");
                titlerow.getCell(5).setCellFormula("IF("+titlerow.getCell(3).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+"<0,\"×\",\"\")");
                //左拱腰2
                titlerow.getCell(6).setCellValue(Double.valueOf(String.format("%.0f", Double.valueOf(row.getZgy2()))).intValue());
                titlerow.getCell(7).setCellFormula("IF("+titlerow.getCell(6).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+">=0,\"√\",\"\")");
                titlerow.getCell(8).setCellFormula("IF("+titlerow.getCell(6).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+"<0,\"×\",\"\")");
                //左拱腰3
                titlerow.getCell(9).setCellValue(Double.valueOf(String.format("%.0f", Double.valueOf(row.getZgy3()))).intValue());
                titlerow.getCell(10).setCellFormula("IF("+titlerow.getCell(9).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+">=0,\"√\",\"\")");
                titlerow.getCell(11).setCellFormula("IF("+titlerow.getCell(9).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+"<0,\"×\",\"\")");
                //右拱腰1
                titlerow.getCell(12).setCellValue(Double.valueOf(String.format("%.0f", Double.valueOf(row.getYgy1()))).intValue());
                titlerow.getCell(13).setCellFormula("IF("+titlerow.getCell(12).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+">=0,\"√\",\"\")");
                titlerow.getCell(14).setCellFormula("IF("+titlerow.getCell(12).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+"<0,\"×\",\"\")");
                //右拱腰2
                titlerow.getCell(15).setCellValue(Double.valueOf(String.format("%.0f", Double.valueOf(row.getYgy2()))).intValue());
                titlerow.getCell(16).setCellFormula("IF("+titlerow.getCell(15).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+">=0,\"√\",\"\")");
                titlerow.getCell(17).setCellFormula("IF("+titlerow.getCell(15).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+"<0,\"×\",\"\")");
                //右拱腰3
                titlerow.getCell(18).setCellValue(Double.valueOf(String.format("%.0f", Double.valueOf(row.getYgy3()))).intValue());
                titlerow.getCell(19).setCellFormula("IF("+titlerow.getCell(18).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+">=0,\"√\",\"\")");
                titlerow.getCell(20).setCellFormula("IF("+titlerow.getCell(18).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+"<0,\"×\",\"\")");
                //拱顶
                titlerow.getCell(21).setCellValue(Double.valueOf(String.format("%.0f", Double.valueOf(row.getGd()))).intValue());
                titlerow.getCell(22).setCellFormula("IF("+titlerow.getCell(21).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+">=0,\"√\",\"\")");
                titlerow.getCell(23).setCellFormula("IF("+titlerow.getCell(21).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+"<0,\"×\",\"\")");
                curRow ++;
            }
            XSSFRow row = sheet.getRow(sheet.getLastRowNum());
            curRow = sheet.getLastRowNum();
            row.getCell(4).setCellFormula("COUNT(D8:V"+curRow+")");
            row.getCell(10).setCellFormula("COUNTIF(E8:F"+curRow+",\"√\")+COUNTIF(H8:I"+curRow+",\"√\")"
                    + "+COUNTIF(K8:L"+curRow+",\"√\")+COUNTIF(N8:O"+curRow+",\"√\")+COUNTIF(Q8:R"+curRow+",\"√\")"
                    + "+COUNTIF(T8:U"+curRow+",\"√\")+COUNTIF(W8:X"+curRow+",\"√\")");
            row.getCell(15).setCellFormula(row.getCell(10).getReference()+"*100/"+row.getCell(4).getReference());
        }
    }

    private void DBtoExcel3(List<JjgFbgcSdgcCqhd> data, XSSFWorkbook wb, String sheetname, String sdmc) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        if (data.size()>0){
            XSSFSheet sheet = wb.getSheet(sheetname);
            XSSFRow titlerow = sheet.getRow(1);
            titlerow.getCell(0).setCellValue("项目名称："+data.get(0).getProname());
            titlerow.getCell(12).setCellValue("合同段："+data.get(0).getHtd());
            titlerow = sheet.getRow(2);
            titlerow.getCell(0).setCellValue("单位工程名称："+sdmc);
            titlerow = sheet.getRow(3);
            titlerow.getCell(12).setCellValue("检测时间："+simpleDateFormat.format(data.get(0).getJcsj()));

            //填写数据
            int curRow = 7;
            for(JjgFbgcSdgcCqhd row : data){
                titlerow = sheet.getRow(curRow);
                //桩号
                titlerow.getCell(0).setCellValue(row.getZh());
                //位置
                titlerow.getCell(1).setCellValue(row.getWz());
                //设计厚度
                titlerow.getCell(2).setCellValue(Double.valueOf(String.format("%.0f", Double.valueOf(row.getSjhd()))).intValue());
                //左拱腰1
                titlerow.getCell(3).setCellValue(Double.valueOf(String.format("%.0f", Double.valueOf(row.getZgy1()))).intValue());
                titlerow.getCell(4).setCellFormula("IF("+titlerow.getCell(3).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+">=0,\"√\",\"\")");
                titlerow.getCell(5).setCellFormula("IF("+titlerow.getCell(3).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+"<0,\"×\",\"\")");
                //左拱腰2
                titlerow.getCell(6).setCellValue(Double.valueOf(String.format("%.0f", Double.valueOf(row.getZgy2()))).intValue());
                titlerow.getCell(7).setCellFormula("IF("+titlerow.getCell(6).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+">=0,\"√\",\"\")");
                titlerow.getCell(8).setCellFormula("IF("+titlerow.getCell(6).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+"<0,\"×\",\"\")");
                //右拱腰1
                titlerow.getCell(9).setCellValue(Double.valueOf(String.format("%.0f", Double.valueOf(row.getYgy1()))).intValue());
                titlerow.getCell(10).setCellFormula("IF("+titlerow.getCell(9).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+">=0,\"√\",\"\")");
                titlerow.getCell(11).setCellFormula("IF("+titlerow.getCell(9).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+"<0,\"×\",\"\")");
                //右拱腰2
                titlerow.getCell(12).setCellValue(Double.valueOf(String.format("%.0f", Double.valueOf(row.getYgy2()))).intValue());
                titlerow.getCell(13).setCellFormula("IF("+titlerow.getCell(12).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+">=0,\"√\",\"\")");
                titlerow.getCell(14).setCellFormula("IF("+titlerow.getCell(12).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+"<0,\"×\",\"\")");
                //拱顶
                titlerow.getCell(15).setCellValue(Double.valueOf(String.format("%.0f", Double.valueOf(row.getGd()))).intValue());
                titlerow.getCell(16).setCellFormula("IF("+titlerow.getCell(15).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+">=0,\"√\",\"\")");
                titlerow.getCell(17).setCellFormula("IF("+titlerow.getCell(15).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+"<0,\"×\",\"\")");
                curRow ++;
            }
            XSSFRow row = sheet.getRow(sheet.getLastRowNum());
            curRow = sheet.getLastRowNum();
            row.getCell(4).setCellFormula("COUNT(D8:R"+curRow+")");
            row.getCell(10).setCellFormula("COUNTIF(E8:F"+curRow+",\"√\")+COUNTIF(H8:I"+curRow+",\"√\")"
                    + "+COUNTIF(K8:L"+curRow+",\"√\")+COUNTIF(N8:O"+curRow+",\"√\")+COUNTIF(Q8:R"+curRow+",\"√\")");
            row.getCell(15).setCellFormula(row.getCell(10).getReference()+"*100/"+row.getCell(4).getReference());

        }

    }

    private boolean DBtoExcel2(List<JjgFbgcSdgcCqhd> data, XSSFWorkbook wb, String sheetname, String sdmc) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        if (data.size()>0){
            XSSFSheet sheet = wb.getSheet(sheetname);
            //首先填写表头
            XSSFRow titlerow = sheet.getRow(1);
            titlerow.getCell(0).setCellValue("项目名称："+data.get(0).getProname());
            titlerow.getCell(8).setCellValue("合同段："+data.get(0).getHtd());
            titlerow = sheet.getRow(2);
            titlerow.getCell(0).setCellValue("单位工程名称："+sdmc);
            titlerow = sheet.getRow(3);
            titlerow.getCell(8).setCellValue("检测时间："+simpleDateFormat.format(data.get(0).getJcsj()));
            //填写数据
            int curRow = 7;
            for(JjgFbgcSdgcCqhd row : data){
                titlerow = sheet.getRow(curRow);
                //桩号
                titlerow.getCell(0).setCellValue(row.getZh());
                //位置
                titlerow.getCell(1).setCellValue(row.getWz());
                //设计厚度
                titlerow.getCell(2).setCellValue(Double.valueOf(String.format("%.0f", Double.valueOf(row.getSjhd()))).intValue());
                //左拱腰1
                titlerow.getCell(3).setCellValue(Double.valueOf(String.format("%.0f", Double.valueOf(row.getZgy1()))).intValue());
                titlerow.getCell(4).setCellFormula("IF("+titlerow.getCell(3).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+">=0,\"√\",\"\")");
                titlerow.getCell(5).setCellFormula("IF("+titlerow.getCell(3).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+"<0,\"×\",\"\")");
                //右拱腰1
                titlerow.getCell(6).setCellValue(Double.valueOf(String.format("%.0f", Double.valueOf(row.getYgy1()))).intValue());
                titlerow.getCell(7).setCellFormula("IF("+titlerow.getCell(6).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+">=0,\"√\",\"\")");
                titlerow.getCell(8).setCellFormula("IF("+titlerow.getCell(6).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+"<0,\"×\",\"\")");
                //拱顶
                titlerow.getCell(9).setCellValue(Double.valueOf(String.format("%.0f", Double.valueOf(row.getGd()))).intValue());
                titlerow.getCell(10).setCellFormula("IF("+titlerow.getCell(9).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+">=0,\"√\",\"\")");
                titlerow.getCell(11).setCellFormula("IF("+titlerow.getCell(9).getReference()
                        +"-$"+titlerow.getCell(2).getReference()+"<0,\"×\",\"\")");
                curRow ++;
            }
            XSSFRow row = sheet.getRow(sheet.getLastRowNum());
            curRow = sheet.getLastRowNum();

            row.getCell(3).setCellFormula("COUNT(D8:L"+curRow+")");
            row.getCell(7).setCellFormula("COUNTIF(E8:F"+curRow+",\"√\")"
                    + "+COUNTIF(H8:I"+curRow+",\"√\")+COUNTIF(K8:L"+curRow+",\"√\")");
            row.getCell(10).setCellFormula(row.getCell(7).getReference()+"*100/"+row.getCell(3).getReference());
            return true;
        }else{
            return false;
        }
    }


    private void createTable3(int tableNum, XSSFWorkbook wb, String sheetname) {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            RowCopy.copyRows(wb, sheetname, sheetname, 7, 37, 7 + i * 31);
        }

        RowCopy.copyRows(wb, "source", sheetname, 0, 0, record * 31+6);


        if(record >= 1)
            wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 17, 0, record * 31 + 6);
    }


    private void createTable2(int tableNum,XSSFWorkbook wb, String sheetname) {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            RowCopy.copyRows(wb, sheetname, sheetname, 7, 37, 7 + i * 31);
        }
        RowCopy.copyRows(wb, "source1", sheetname, 0, 0, record * 31+6);
        if(record >= 1)
            wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 11, 0, record * 31 + 6);
    }


    private void createTable4(int tableNum, XSSFWorkbook wb, String sheetname) {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            RowCopy.copyRows(wb, sheetname, sheetname, 7, 23, 7 + i * 17);
        }

        RowCopy.copyRows(wb, "source", sheetname, 2, 2, record * 17 + 6);

        if(record >= 1)
            wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 23, 0, record * 17 + 6);
    }

    private int gettable4(int size) {
        return size%17 ==0 ? size/17 : size/17+1;

    }



    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        List<Map<String,Object>> mapList = new ArrayList<>();
        List<Map<String, Object>> selectsdmc = selectsdmc2(proname, htd);
        for (int i=0;i<selectsdmc.size();i++) {
            String sdmc = selectsdmc.get(i).get("sdmc").toString();
            String cdnum = "";
            Map<String,Object> jgmap = new HashMap<>();
            File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "39隧道衬砌厚度-" + sdmc + ".xlsx");
            if (!f.exists()) {
                return null;
            } else {
                List<Map<String,Object>> cd = jjgFbgcSdgcCqhdMapper.selectcd2(proname,htd,sdmc);
                if (cd.get(0).get("zgy1") != null && !"".equals(cd.get(0).get("zgy1")) && cd.get(0).get("zgy2") !=null && !"".equals(cd.get(0).get("zgy2")) && cd.get(0).get("zgy3") !=null && !"".equals(cd.get(0).get("zgy3"))){
                    cdnum = "4车道";
                }else if (cd.get(0).get("zgy1") != null && !"".equals(cd.get(0).get("zgy1")) && cd.get(0).get("zgy2") !=null && !"".equals(cd.get(0).get("zgy2"))){
                    cdnum = "3车道";
                }else if (cd.get(0).get("zgy1") != null && !"".equals(cd.get(0).get("zgy1"))){
                    cdnum = "2车道";
                }
                List<Map<String,Object>> wz = jjgFbgcSdgcCqhdMapper.selectwz2(proname,htd,sdmc);
                String wzname = wz.get(0).get("wz").toString();
                String sheetname = cdnum+wzname;
                XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));
                XSSFSheet sheet = xwb.getSheet(sheetname);
                XSSFCell xmname = sheet.getRow(1).getCell(0);//项目名
                XSSFCell htdname = sheet.getRow(1).getCell(8);//合同段名
                String name = "项目名称："+proname;
                String he = "合同段："+htd;
                if (name.equals(xmname.toString()) && he.equals(htdname.toString())) {
                    if (cdnum.equals("2车道")){
                        int lastRowNum = sheet.getLastRowNum();
                        sheet.getRow(lastRowNum).getCell(3).setCellType(CellType.STRING);
                        sheet.getRow(lastRowNum).getCell(7).setCellType(CellType.STRING);
                        sheet.getRow(lastRowNum).getCell(10).setCellType(CellType.STRING);
                        jgmap.put("检测项目",sdmc);
                        jgmap.put("检测总点数",decf.format(Double.valueOf(sheet.getRow(lastRowNum).getCell(3).getStringCellValue())));
                        jgmap.put("合格点数",decf.format(Double.valueOf(sheet.getRow(lastRowNum).getCell(7).getStringCellValue())));
                        jgmap.put("合格率",df.format(Double.valueOf(sheet.getRow(lastRowNum).getCell(10).getStringCellValue())));
                    }else if(cdnum.equals("3车道") || cdnum.equals("4车道") ) {
                        int lastRowNum = sheet.getLastRowNum();
                        sheet.getRow(lastRowNum).getCell(4).setCellType(CellType.STRING);
                        sheet.getRow(lastRowNum).getCell(10).setCellType(CellType.STRING);
                        sheet.getRow(lastRowNum).getCell(15).setCellType(CellType.STRING);
                        jgmap.put("检测项目", sdmc);
                        jgmap.put("检测总点数", decf.format(Double.valueOf(sheet.getRow(lastRowNum).getCell(4).getStringCellValue())));
                        jgmap.put("合格点数", decf.format(Double.valueOf(sheet.getRow(lastRowNum).getCell(10).getStringCellValue())));
                        jgmap.put("合格率", df.format(Double.valueOf(sheet.getRow(lastRowNum).getCell(15).getStringCellValue())));
                    }
                    mapList.add(jgmap);
                }
            }
        }
        return mapList;
    }

    @Override
    public void exportsdcqhd(HttpServletResponse response) {
        String fileName = "02隧道衬砌厚度度实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcSdgcCqhdVo()).finish();


    }

    @Override
    public void importsdcqhd(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcSdgcCqhdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcSdgcCqhdVo>(JjgFbgcSdgcCqhdVo.class) {
                                @Override
                                public void handle(List<JjgFbgcSdgcCqhdVo> dataList) {
                                    for(JjgFbgcSdgcCqhdVo sdgcCqhdVo: dataList)
                                    {
                                        JjgFbgcSdgcCqhd fbgcSdgcCqhd = new JjgFbgcSdgcCqhd();
                                        BeanUtils.copyProperties(sdgcCqhdVo,fbgcSdgcCqhd);
                                        fbgcSdgcCqhd.setCreatetime(new Date());
                                        fbgcSdgcCqhd.setProname(commonInfoVo.getProname());
                                        fbgcSdgcCqhd.setHtd(commonInfoVo.getHtd());
                                        fbgcSdgcCqhd.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcSdgcCqhdMapper.insert(fbgcSdgcCqhd);
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
        List<Map<String,Object>> sdmclist = jjgFbgcSdgcCqhdMapper.selectsdmc(proname,htd,fbgc);
        return sdmclist;
    }

    @Override
    public List<Map<String, Object>> selectsjhd(CommonInfoVo commonInfoVo,String s) {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        List<Map<String, Object>> list = jjgFbgcSdgcCqhdMapper.selectsjhd(proname,htd,s);
        return list;

    }

    @Override
    public List<Map<String, Object>> lookjg(CommonInfoVo commonInfoVo, String value) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        List<Map<String,Object>> mapList = new ArrayList<>();
        List<Map<String, Object>> selectsdmc = selectsdmc2(proname, htd);
        //for (int i=0;i<selectsdmc.size();i++) {
            String sdmc = StringUtils.substringBetween(value, "-", ".");
            String cdnum = "";
            Map<String,Object> jgmap = new HashMap<>();
            File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + value);
            if (!f.exists()) {
                return null;
            } else {
                List<Map<String,Object>> cd = jjgFbgcSdgcCqhdMapper.selectcd2(proname,htd,sdmc);
                if (cd.get(0).get("zgy1") != null && !"".equals(cd.get(0).get("zgy1")) && cd.get(0).get("zgy2") !=null && !"".equals(cd.get(0).get("zgy2")) && cd.get(0).get("zgy3") !=null && !"".equals(cd.get(0).get("zgy3"))){
                    cdnum = "4车道";
                }else if (cd.get(0).get("zgy1") != null && !"".equals(cd.get(0).get("zgy1")) && cd.get(0).get("zgy2") !=null && !"".equals(cd.get(0).get("zgy2"))){
                    cdnum = "3车道";
                }else if (cd.get(0).get("zgy1") != null && !"".equals(cd.get(0).get("zgy1"))){
                    cdnum = "2车道";
                }
                List<Map<String,Object>> wz = jjgFbgcSdgcCqhdMapper.selectwz2(proname,htd,sdmc);
                String wzname = wz.get(0).get("wz").toString();
                String sheetname = cdnum+wzname;
                XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));
                XSSFSheet sheet = xwb.getSheet(sheetname);
                XSSFCell xmname = sheet.getRow(1).getCell(0);//项目名
                XSSFCell htdname = sheet.getRow(1).getCell(8);//合同段名
                String name = "项目名称："+proname;
                String he = "合同段："+htd;
                if (name.equals(xmname.toString()) && he.equals(htdname.toString())) {
                    if (cdnum.equals("2车道")){
                        int lastRowNum = sheet.getLastRowNum();
                        sheet.getRow(lastRowNum).getCell(3).setCellType(CellType.STRING);
                        sheet.getRow(lastRowNum).getCell(7).setCellType(CellType.STRING);
                        sheet.getRow(lastRowNum).getCell(10).setCellType(CellType.STRING);
                        jgmap.put("检测项目",sdmc);
                        jgmap.put("检测总点数",decf.format(Double.valueOf(sheet.getRow(lastRowNum).getCell(3).getStringCellValue())));
                        jgmap.put("合格点数",decf.format(Double.valueOf(sheet.getRow(lastRowNum).getCell(7).getStringCellValue())));
                        jgmap.put("合格率",df.format(Double.valueOf(sheet.getRow(lastRowNum).getCell(10).getStringCellValue())));
                    }else if(cdnum.equals("3车道") || cdnum.equals("4车道") ) {
                        int lastRowNum = sheet.getLastRowNum();
                        sheet.getRow(lastRowNum).getCell(4).setCellType(CellType.STRING);
                        sheet.getRow(lastRowNum).getCell(10).setCellType(CellType.STRING);
                        sheet.getRow(lastRowNum).getCell(15).setCellType(CellType.STRING);
                        jgmap.put("检测项目", sdmc);
                        jgmap.put("检测总点数", decf.format(Double.valueOf(sheet.getRow(lastRowNum).getCell(4).getStringCellValue())));
                        jgmap.put("合格点数", decf.format(Double.valueOf(sheet.getRow(lastRowNum).getCell(10).getStringCellValue())));
                        jgmap.put("合格率", df.format(Double.valueOf(sheet.getRow(lastRowNum).getCell(15).getStringCellValue())));
                    }
                    mapList.add(jgmap);
                }
            }
        //}
        return mapList;
    }

    @Override
    public int getds(CommonInfoVo commonInfoVo, String sdmc) {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        List<Map<String,Object>> list = jjgFbgcSdgcCqhdMapper.getds(proname,htd,sdmc);
        System.out.println(list);
        int num=0;
        for (Map<String, Object> map : list) {
            double sjhd = Double.valueOf(map.get("sjhd").toString());
            if (!map.get("zgy1").toString().equals("") && map.get("zgy1")!=null){
                double zgy = Double.valueOf(map.get("zgy1").toString());
                if (zgy < sjhd/2){
                    num += 1;
                }
            }
            if (map.get("zgy2")!=null && !map.get("zgy2").toString().equals("")){
                double zgy = Double.valueOf(map.get("zgy2").toString());
                if (zgy < sjhd/2){
                    num += 1;
                }
            }
            if (map.get("zgy3")!=null){
                double zgy = Double.valueOf(map.get("zgy3").toString());
                if (zgy < sjhd/2){
                    num += 1;
                }
            }

            if (map.get("ygy1")!=null){
                double zgy = Double.valueOf(map.get("ygy1").toString());
                if (zgy < sjhd/2){
                    num += 1;
                }
            }
            if (map.get("ygy2")!=null){
                double zgy = Double.valueOf(map.get("ygy2").toString());
                if (zgy < sjhd/2){
                    num += 1;
                }
            }
            if (map.get("ygy3")!=null){
                double zgy = Double.valueOf(map.get("ygy3").toString());
                if (zgy < sjhd/2){
                    num += 1;
                }
            }
        }
        return num;

    }

    public List<Map<String, Object>> selectsdmc2(String proname, String htd) {
        List<Map<String,Object>> sdmclist = jjgFbgcSdgcCqhdMapper.selectsdmc2(proname,htd);
        return sdmclist;
    }
}
