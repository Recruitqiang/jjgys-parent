package glgc.jjgys.system.service.impl;

import cn.hutool.core.io.resource.ClassPathResource;
import cn.hutool.core.io.resource.Resource;
import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLjgcZddmcc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.ljgc.JjgFbgcLjgcZddmccVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcLjgcZddmccMapper;
import glgc.jjgys.system.service.JjgFbgcLjgcZddmccService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.poi.ss.usermodel.CellType;
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
public class JjgFbgcLjgcZddmccServiceImpl extends ServiceImpl<JjgFbgcLjgcZddmccMapper, JjgFbgcLjgcZddmcc> implements JjgFbgcLjgcZddmccService {

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    private static XSSFWorkbook wb = null;

    @Autowired
    private JjgFbgcLjgcZddmccMapper jjgFbgcLjgcZddmccMapper;


    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获得数据
        QueryWrapper<JjgFbgcLjgcZddmcc> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("zhjwz","cjzh","bw");
        List<JjgFbgcLjgcZddmcc> data = jjgFbgcLjgcZddmccMapper.selectList(wrapper);

        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"11路基支挡断面尺寸.xlsx");
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
            String name = "支挡断面尺寸.xlsx";
            String path =reportPath+File.separator+name;
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream in = new FileInputStream(f);
            wb = new XSSFWorkbook(in);
            createTable(gettableNum(data.size()));
            String sheetname = "支挡断面尺寸";
            if(DBtoExcel(data,proname,htd,fbgc,sheetname)){
                calculateSheet(wb.getSheet("支挡断面尺寸"));
                for (int j = 0; j < wb.getNumberOfSheets(); j++) {   //表内公式  计算 显示结果
                    JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
                }
                FileOutputStream fileOut = new FileOutputStream(f);
                wb.write(fileOut);
                fileOut.flush();
                fileOut.close();

            }

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
                        +rowstart.getCell(4+1).getReference()+":"
                        +rowend.getCell(4+1).getReference()+")");//B32=COUNT(D6:D31) B=COUNT(E6:E68)
                row.getCell(3).setCellFormula("COUNTIF("
                        +rowstart.getCell(6+1).getReference()+":"
                        +rowend.getCell(6+1).getReference()+",\"√\")");//D32=COUNTIF(H6:H31,"√") D=COUNTIF(G6:G34,"√")+COUNTIF(G41:G68,"√")
                row.getCell(6).setCellFormula("ROUND("+
                        row.getCell(3).getReference()+"/"
                        +row.getCell(1).getReference()+"*100,1)");//H32=D32/B32*100 G=ROUND(D69/B69*100,1)
                break;
            }
            if(flag && !"".equals(row.getCell(4+1).toString())){
                if("不小于设计值".equals(row.getCell(5+1).toString())){
                    row.getCell(6+1).setCellFormula("IF("
                            +row.getCell(4+1).getReference()+">="
                            +row.getCell(3+1).getReference()+",\"√\",\"\")");//G=IF((E41+30>=D41)*(E41-30<=D41),"√","")
                    row.getCell(7+1).setCellFormula("IF("
                            +row.getCell(6+1).getReference()+"=\"\""
                            +",\"×\",\"\")");//H=IF((E41+30>=D41)*(E41-30<=D41),"","×")
                }
                else{
                    row.getCell(6+1).setCellFormula("IF(("
                            +row.getCell(4+1).getReference()+"+"
                            +JjgFbgcCommonUtils.getAllowError1(row.getCell(5+1).toString())+">="
                            +row.getCell(3+1).getReference()+")*("
                            +row.getCell(4+1).getReference()+"-"
                            +JjgFbgcCommonUtils.getAllowError2(row.getCell(5+1).toString())+"<="
                            +row.getCell(3+1).getReference()+"),\"√\",\"\")");//G=IF((E41+30>=D41)*(E41-30<=D41),"√","")

                    row.getCell(7+1).setCellFormula("IF(("
                            +row.getCell(4+1).getReference()+"+"
                            +JjgFbgcCommonUtils.getAllowError1(row.getCell(5+1).toString())+">="
                            +row.getCell(3+1).getReference()+")*("
                            +row.getCell(4+1).getReference()+"-"
                            +JjgFbgcCommonUtils.getAllowError2(row.getCell(5+1).toString())+"<="
                            +row.getCell(3+1).getReference()+"),\"\",\"×\")");//H=IF((E41+30>=D41)*(E41-30<=D41),"","×")
                }
            }
            if ("段落".equals(row.getCell(0).toString())) {
                rowstart = sheet.getRow(i+2);
                rowend = rowstart;
                i++;
                flag = true;
            }
        }
    }

    public boolean DBtoExcel(List<JjgFbgcLjgcZddmcc> data,String proname,String htd,String fbgc,String sheetname) throws IOException {
        XSSFCellStyle cellstyle = JjgFbgcCommonUtils.dBtoExcelUtils(wb);
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        String time = simpleDateFormat.format(data.get(0).getJcsj());
        //获取了最新的检测时间
        for(JjgFbgcLjgcZddmcc row:data){
            try {
                Date dt1 = simpleDateFormat.parse(time);
                Date dt2 = row.getJcsj();
                if(dt1.getTime() < dt2.getTime()){
                    time = simpleDateFormat.format(row.getJcsj());
                }
            } catch (ParseException e) {
                e.printStackTrace();
            }
        }
        XSSFSheet sheet = wb.getSheet(sheetname);

        String endpileNO = data.get(0).getZhjwz();//桩号及位置
        String pile = data.get(0).getCjzh();//抽检桩号
        String position = data.get(0).getBw();//部位
        String categroy = data.get(0).getLb();//类别

        int start = 5;
        int positionstart = 5;
        int categorystart = 5;

        int index = 0;
        int page = 0;
        String pileNO = "";
        ArrayList<String> compaction = new ArrayList<String>();
        pileNO = endpileNO;
        sheet.getRow(1).getCell(1).setCellValue(proname);
        sheet.getRow(1).getCell(6).setCellValue(htd);
        sheet.getRow(2).getCell(1).setCellValue(fbgc);
        sheet.getRow(2).getCell(6).setCellValue(time);
        for(JjgFbgcLjgcZddmcc row:data){
            fillCommonCellData(sheet,index, row);
            index++;
        }
        //  新合并单元格
        XSSFRow row = null;
        int S=5,E=5;//  判断  段落  开始结束row
        int zS=5,zE=5; //  判断  桩号  开始结束row
        int pS=5,pE=5; //  判断  部位  开始结束row
        int cS=5,cE=5; //  判断  类别  开始结束row
        boolean flag = false;
        String pileNo = "";
        int record = 0;
        for (int i = sheet.getFirstRowNum()+5; i < sheet.getPhysicalNumberOfRows(); i++) {
            for(int temp=S;temp < sheet.getPhysicalNumberOfRows();temp++)
            {
                if(sheet.getRow(temp).getCell(0).toString() == null || "".equals(sheet.getRow(temp).getCell(0).toString())){  //结束
                    System.out.println("TEMP空="+sheet.getRow(temp).getCell(0).toString());
                    break;
                }
                if(temp==33||(temp-33)%29==0) //分页处
                {
                    E=temp;
                    System.out.println("分页S="+S+"分页  E="+E);
                    break;
                }
                else if(!sheet.getRow(S).getCell(0).toString().equals(sheet.getRow(temp+1).getCell(0).toString()))  //与起始段落相同
                {
                    E=temp;
                    System.out.println("S="+S+" E="+E);
                    break;
                }
            }
            if(sheet.getRow(S).getCell(0).toString() == null || "".equals(sheet.getRow(S).getCell(0).toString())){  //结束
                System.out.println("S空="+sheet.getRow(S).getCell(0).toString());
                break;
            }
            if(E>5&&E-S>0) {  //至少两个 才能合并单元格
                sheet.addMergedRegion(new CellRangeAddress(S, E, 0, 0));
            }
            System.out.println("final S="+S+" E="+E);
            zS=S;
            zE=S;
            for(int temp = S;temp <= E;temp ++)    //合并 桩号
            {
                for(int ztemp = zS;ztemp <= E;ztemp ++)
                {
                    if(ztemp==33||(ztemp-33)%29==0) //分页处
                    {
                        zE=ztemp;
                        System.out.println("分页zS="+zS+" zE="+zE);
                        break;
                    }
                    else if(!sheet.getRow(zS).getCell(1).toString().equals(sheet.getRow(ztemp+1).getCell(1).toString()))
                    {
                        zE=ztemp;
                        System.out.println("zS="+zS+" zE="+zE);
                        break;
                    }
                }
                System.out.println("final zS="+zS+" zE="+zE);
                if(zE-zS>0) {  //至少两个 才能合并单元格
                    sheet.addMergedRegion(new CellRangeAddress(zS, zE, 1, 1));
                }
                pS=zS;
                pE=zE;
                for(int ttemp = zS;ttemp <= zE;ttemp ++)    //合并 部位
                {
                    for(int ptemp = pS;ptemp <= zE;ptemp ++)
                    {
                        if(ptemp==33||(ptemp-33)%29==0) //分页处
                        {
                            pE=ptemp;
                            System.out.println("分页pS="+zS+" pE="+zE);
                            break;
                        }
                        else if(!sheet.getRow(pS).getCell(2).toString().equals(sheet.getRow(ptemp+1).getCell(2).toString()))
                        {
                            pE=ptemp;
                            System.out.println("pS="+pS+" pE="+pE);
                            break;
                        }
                    }
                    System.out.println("final pS="+pS+" pE="+pE);
                    if(pE-pS>0) {  //至少两个 才能合并单元格
                        sheet.addMergedRegion(new CellRangeAddress(pS, pE, 2, 2));
                    }
                    pS=pE+1;//部位
                }
                cS=zS;//类别
                cE=zE;
                for(int ttemp = zS;ttemp <= zE;ttemp ++)    //合并 类别
                {
                    for(int ctemp = cS;ctemp <= zE;ctemp ++)
                    {
                        if(ctemp==33||(ctemp-33)%29==0) //分页处
                        {
                            cE=ctemp;
                            System.out.println("分页cS="+cS+" cE="+cE);
                            break;
                        }
                        else if(!sheet.getRow(cS).getCell(3).toString().equals(sheet.getRow(ctemp+1).getCell(3).toString()))
                        {
                            cE=ctemp;
                            System.out.println("cS="+cS+" cE="+cE);
                            break;
                        }
                    }
                    System.out.println("final cS="+cS+" cE="+cE);
                    if(cE-cS>0) {  //至少两个 才能合并单元格
                        sheet.addMergedRegion(new CellRangeAddress(cS, cE, 3, 3));
                    }
                    cS=cE+1;//部位
                }
                zS=zE+1;//桩号
            }
            S=E+1;//段落
        }
        return true;

    }

    public void fillCommonCellData(XSSFSheet sheet, int index, JjgFbgcLjgcZddmcc row) {
        sheet.getRow(index+5).getCell(0).setCellValue(row.getZhjwz());//部位
        sheet.getRow(index+5).getCell(1).setCellValue(row.getCjzh());//部位
        sheet.getRow(index+5).getCell(2).setCellValue(row.getBw());//桩号
        sheet.getRow(index+5).getCell(3).setCellValue(row.getLb());//类别
        sheet.getRow(index+5).getCell(4).setCellValue(Double.parseDouble(row.getSjz()));
        sheet.getRow(index+5).getCell(5).setCellValue(Double.parseDouble(row.getScz()));

        if(row.getYxwcz().equals(row.getYxwcf())){
            if("不小于设计值".equals(row.getYxwcf()) || "不小于设计值".equals(row.getYxwcz()) ){
                sheet.getRow(index+5).getCell(5+1).setCellValue("不小于设计值");
            }
            else{
                sheet.getRow(index+5).getCell(5+1).setCellValue("±"+row.getYxwcz());
            }
        }
        else{
            sheet.getRow(index+5).getCell(5+1).setCellValue("+"+row.getYxwcz()+",-"+row.getYxwcf());
        }

    }


    public int gettableNum(int size){
        return size%29 <= 27 ? size/29+1 : size/29+2;
    }

    public void createTable(int tableNum) throws IOException {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            if(i < record-1){
                RowCopy.copyRows(wb, "支挡断面尺寸", "支挡断面尺寸", 5, 33, (i - 1) * 29 + 34);
            }
            else{
                RowCopy.copyRows(wb, "支挡断面尺寸", "支挡断面尺寸", 5, 31, (i - 1) * 29 + 34);
            }
        }
        if(record == 1){
            wb.getSheet("支挡断面尺寸").shiftRows(33, 34, -1);
        }
        RowCopy.copyRows(wb, "source", "支挡断面尺寸", 0, 1,(record) * 29 + 3);
        wb.setPrintArea(wb.getSheetIndex("支挡断面尺寸"), 0, 8, 0,(record) * 29 + 4);
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String,Object>> mapList = new ArrayList<>();
        Map<String,Object> jgmap = new HashMap<>();
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        String title = "结构（断面）尺寸质量鉴定表";
        String sheetname="支挡断面尺寸";
        //获取鉴定表文件
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"11路基支挡断面尺寸.xlsx");
        if(!f.exists()){
            return null;
        }else {
            //创建工作簿
            XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));
            //读取工作表
            XSSFSheet slSheet = xwb.getSheet(sheetname);
            int lastRowNum = slSheet.getLastRowNum();
            slSheet.getRow(lastRowNum-1).getCell(1).setCellType(CellType.STRING);
            slSheet.getRow(lastRowNum-1).getCell(3).setCellType(CellType.STRING);
            slSheet.getRow(lastRowNum-1).getCell(6).setCellType(CellType.STRING);
            double zds= Double.valueOf(slSheet.getRow(lastRowNum-1).getCell(1).getStringCellValue());
            double hgds= Double.valueOf(slSheet.getRow(lastRowNum-1).getCell(3).getStringCellValue());
            double hgl= Double.valueOf(slSheet.getRow(lastRowNum-1).getCell(6).getStringCellValue());
            String zdsz = decf.format(zds);
            String hgdsz = decf.format(hgds);
            String hglz = df.format(hgl);
            jgmap.put("检测总点数",zdsz);
            jgmap.put("合格点数",hgdsz);
            jgmap.put("合格率",hglz);
            mapList.add(jgmap);
        }
        return mapList;
    }

    @Override
    public void exportzddmcc(HttpServletResponse response) {
        String fileName = "11支挡断面尺寸实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcLjgcZddmccVo()).finish();

    }

    @Override
    public void importzddmcc(MultipartFile file,CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcLjgcZddmccVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcLjgcZddmccVo>(JjgFbgcLjgcZddmccVo.class) {
                                @Override
                                public void handle(List<JjgFbgcLjgcZddmccVo> dataList) {
                                    for(JjgFbgcLjgcZddmccVo zddmccVo: dataList)
                                    {
                                        JjgFbgcLjgcZddmcc fbgcLjgcZddmcc = new JjgFbgcLjgcZddmcc();
                                        BeanUtils.copyProperties(zddmccVo,fbgcLjgcZddmcc);
                                        fbgcLjgcZddmcc.setCreatetime(new Date());
                                        fbgcLjgcZddmcc.setProname(commonInfoVo.getProname());
                                        fbgcLjgcZddmcc.setHtd(commonInfoVo.getHtd());
                                        fbgcLjgcZddmcc.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcLjgcZddmccMapper.insert(fbgcLjgcZddmcc);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }

    @Override
    public List<Map<String, Object>> selectyxps(String proname, String htd) {
        List<Map<String,Object>>  list = jjgFbgcLjgcZddmccMapper.selectyxps(proname,htd);
        return list;
    }

    @Override
    public Map<String, Object> selectchs(String proname, String htd) {
        Map<String, Object> map = jjgFbgcLjgcZddmccMapper.selectchs(proname,htd);
        return map;
    }
}
