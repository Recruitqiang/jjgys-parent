package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLjgcHdgqd;
import glgc.jjgys.model.project.JjgFbgcSdgcCqtqd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.ljgc.JjgFbgcLjgcHdgqdVo;
import glgc.jjgys.model.projectvo.sdgc.JjgFbgcSdgcCqtqdVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcSdgcCqtqdMapper;
import glgc.jjgys.system.service.JjgFbgcLjgcHdgqdService;
import glgc.jjgys.system.service.JjgFbgcSdgcCqtqdService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
 * @since 2023-03-26
 */
@Service
public class JjgFbgcSdgcCqtqdServiceImpl extends ServiceImpl<JjgFbgcSdgcCqtqdMapper, JjgFbgcSdgcCqtqd> implements JjgFbgcSdgcCqtqdService {

    @Autowired
    private JjgFbgcSdgcCqtqdMapper jjgFbgcSdgcCqtqdMapper;

    @Autowired
    private JjgFbgcLjgcHdgqdService jjgFbgcLjgcHdgqdService;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获取数据
        QueryWrapper<JjgFbgcSdgcCqtqd> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("sdmc","bw1");
        List<JjgFbgcSdgcCqtqd> data = jjgFbgcSdgcCqtqdMapper.selectList(wrapper);
        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"38隧道衬砌砼强度.xlsx");
        //健壮性判断如果没有数据返回"请导入数据"
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
            String name = "涵洞砼强度.xlsx";
            String path =reportPath + File.separator+name;
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            createTable(gettableNum(data.size()),wb);
            if(DBtoExcel(data,wb)){
                //设置公式,计算合格点数
                jjgFbgcLjgcHdgqdService.calculateSheet(wb.getSheet("原始数据"));
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

    private boolean DBtoExcel(List<JjgFbgcSdgcCqtqd> data,XSSFWorkbook wb) throws ParseException {
        String proname = data.get(0).getProname();
        String htd = data.get(0).getHtd();
        String fbgc = data.get(0).getFbgc();
        XSSFCellStyle cellstyle = JjgFbgcCommonUtils.dBtoExcelUtils(wb);
        //表格数据填写
        XSSFSheet sheet = wb.getSheet("原始数据");
        Date jcsj = data.get(0).getJcsj();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        String testtime = simpleDateFormat.format(jcsj);
        int index = 0;
        int tableNum = 0;
        //填写表头
        fillTitleCellData(sheet,tableNum,proname,htd,fbgc);
        //给每个table填写表头
        for(JjgFbgcSdgcCqtqd row:data){
            //比较检测时间，拿到最新的检测时间
            testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.getJcsj()));
            if(index/10 == 2){
                sheet.getRow(tableNum*25+2).getCell(29).setCellValue(testtime);
                testtime = simpleDateFormat.format(row.getJcsj());
                tableNum ++;
                fillTitleCellData(sheet,tableNum,proname,htd,fbgc);
                index = 0;
            }
            //填写中间下方的普通单元格
            fillCommonCellData(sheet, tableNum, index+5, row,cellstyle);
            index++;
        }
        sheet.getRow(tableNum*25+2).getCell(29).setCellValue(testtime);
        return true;
    }

    //填写表头
    private void fillTitleCellData(XSSFSheet sheet, int tableNum,String proname,String htd,String fbgc) {
        if(sheet.getRow(tableNum*25+1) == null || sheet.getRow(tableNum*25+1).getCell(2) == null){
            return;
        }
        sheet.getRow(tableNum*25+1).getCell(2).setCellValue(proname);
        sheet.getRow(tableNum*25+1).getCell(29).setCellValue(htd);
        sheet.getRow(tableNum*25+2).getCell(2).setCellValue(fbgc);
    }


    //填写中间下方的普通单元格
    private void fillCommonCellData(XSSFSheet sheet, int tableNum, int index,JjgFbgcSdgcCqtqd row, XSSFCellStyle cellstyle) {
        sheet.getRow(tableNum*25+index).getCell(0).setCellValue(row.getSdmc());
        sheet.getRow(tableNum*25+index).getCell(1).setCellValue(row.getBw1());
        sheet.getRow(tableNum*25+index).getCell(2).setCellValue(row.getBw2());

        if(!"".equals(row.getCdz1()))
            sheet.getRow(tableNum*25+index).getCell(3).setCellValue(Double.valueOf(row.getCdz1()).intValue());
        if(!"".equals(row.getZ2()))
            sheet.getRow(tableNum*25+index).getCell(4).setCellValue(Double.valueOf(row.getZ2()).intValue());
        if(!"".equals(row.getZ3()))
            sheet.getRow(tableNum*25+index).getCell(5).setCellValue(Double.valueOf(row.getZ3()).intValue());
        if(!"".equals(row.getZ4()))
            sheet.getRow(tableNum*25+index).getCell(6).setCellValue(Double.valueOf(row.getZ4()).intValue());
        if(!"".equals(row.getZ5()))
            sheet.getRow(tableNum*25+index).getCell(7).setCellValue(Double.valueOf(row.getZ5()).intValue());
        if(!"".equals(row.getZ6()))
            sheet.getRow(tableNum*25+index).getCell(8).setCellValue(Double.valueOf(row.getZ6()).intValue());
        if(!"".equals(row.getZ7()))
            sheet.getRow(tableNum*25+index).getCell(9).setCellValue(Double.valueOf(row.getZ7()).intValue());
        if(!"".equals(row.getZ8()))
            sheet.getRow(tableNum*25+index).getCell(10).setCellValue(Double.valueOf(row.getZ8()).intValue());
        if(!"".equals(row.getZ9()))
            sheet.getRow(tableNum*25+index).getCell(11).setCellValue(Double.valueOf(row.getZ9()).intValue());
        if(!"".equals(row.getZ10()))
            sheet.getRow(tableNum*25+index).getCell(12).setCellValue(Double.valueOf(row.getZ10()).intValue());
        if(!"".equals(row.getZ11()))
            sheet.getRow(tableNum*25+index).getCell(13).setCellValue(Double.valueOf(row.getZ11()).intValue());
        if(!"".equals(row.getZ12()))
            sheet.getRow(tableNum*25+index).getCell(14).setCellValue(Double.valueOf(row.getZ12()).intValue());
        if(!"".equals(row.getZ13()))
            sheet.getRow(tableNum*25+index).getCell(15).setCellValue(Double.valueOf(row.getZ13()).intValue());
        if(!"".equals(row.getZ14()))
            sheet.getRow(tableNum*25+index).getCell(16).setCellValue(Double.valueOf(row.getZ14()).intValue());
        if(!"".equals(row.getZ15()))
            sheet.getRow(tableNum*25+index).getCell(17).setCellValue(Double.valueOf(row.getZ15()).intValue());
        if(!"".equals(row.getZ16()))
            sheet.getRow(tableNum*25+index).getCell(18).setCellValue(Double.valueOf(row.getZ16()).intValue());

        if("水平".equals(row.getHtjd())){
            sheet.getRow(tableNum*25+index).getCell(20).setCellValue(row.getHtjd());//U回弹角度
        }
        else{
            sheet.getRow(tableNum*25+index).getCell(20).setCellValue(Double.valueOf(row.getHtjd()).intValue());//U回弹角度
        }
        sheet.getRow(tableNum*25+index).getCell(23).setCellValue(row.getJzm());//X浇筑面
        sheet.getRow(tableNum*25+index).getCell(26).setCellValue(Double.valueOf(row.getThsd()));//AA碳化深度
        sheet.getRow(tableNum*25+index).getCell(28).setCellValue(row.getSfbs());//AC是否泵送
        sheet.getRow(tableNum*25+index).getCell(31).setCellValue(Double.valueOf(row.getSjqd()).intValue());//AF设计强度
        sheet.getRow(tableNum*25+index).getCell(31).setCellStyle(cellstyle);

    }

    private void createTable(int tableNum,XSSFWorkbook wb) {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            RowCopy.copyRows(wb, "原始数据", "原始数据", 0, 24, i * 25);
        }
        wb.getSheet("原始数据").setColumnHidden(21, true);
        wb.getSheet("原始数据").setColumnHidden(22, true);
        wb.getSheet("原始数据").setColumnHidden(24, true);
        wb.getSheet("原始数据").setColumnHidden(25, true);
        wb.getSheet("原始数据").setColumnHidden(27, true);
        wb.getSheet("原始数据").setColumnHidden(30, true);
        if(record > 1)
            wb.setPrintArea(wb.getSheetIndex("原始数据"), 0, 33, 0, record * 25-1);
    }

    private int gettableNum(int size) {
        return size%20 ==0 ? size/20 : size/20+1;
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        String title = "混凝土强度质量鉴定表（回弹法）";
        String sheetname = "原始数据";
        //获取鉴定表文件
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"38隧道衬砌砼强度.xlsx");
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
            List<Map<String, Object>> mapList = JjgFbgcCommonUtils.gettqdjcjg(map);
            return mapList;
        }
    }

    @Override
    public void exportsdcqtqd(HttpServletResponse response) {
        String fileName = "01隧道衬砌砼强度实测数据";
        String sheetName = "原始数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcSdgcCqtqdVo()).finish();

    }

    @Override
    public void importsdcqtsd(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcSdgcCqtqdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcSdgcCqtqdVo>(JjgFbgcSdgcCqtqdVo.class) {
                                @Override
                                public void handle(List<JjgFbgcSdgcCqtqdVo> dataList) {
                                    for(JjgFbgcSdgcCqtqdVo sdgcCqtqdVo: dataList)
                                    {
                                        JjgFbgcSdgcCqtqd fbgcSdgcCqtqd = new JjgFbgcSdgcCqtqd();
                                        BeanUtils.copyProperties(sdgcCqtqdVo,fbgcSdgcCqtqd);
                                        fbgcSdgcCqtqd.setCreatetime(new Date());
                                        fbgcSdgcCqtqd.setProname(commonInfoVo.getProname());
                                        fbgcSdgcCqtqd.setHtd(commonInfoVo.getHtd());
                                        fbgcSdgcCqtqd.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcSdgcCqtqdMapper.insert(fbgcSdgcCqtqd);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }
}
