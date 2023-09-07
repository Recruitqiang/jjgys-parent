package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabx;
import glgc.jjgys.model.project.JjgFbgcLjgcHdgqd;
import glgc.jjgys.model.project.JjgFbgcLmgcLqlmysd;
import glgc.jjgys.model.projectvo.jagc.JjgFbgcJtaqssJabxVo;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.lmgc.JjgFbgcLmgcLqlmysdVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcLmgcLqlmysdMapper;
import glgc.jjgys.system.service.JjgFbgcLmgcLqlmysdService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
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
import java.text.Collator;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-04-10
 */
@Service
public class JjgFbgcLmgcLqlmysdServiceImpl extends ServiceImpl<JjgFbgcLmgcLqlmysdMapper, JjgFbgcLmgcLqlmysd> implements JjgFbgcLmgcLqlmysdService {

    @Autowired
    private JjgFbgcLmgcLqlmysdMapper jjgFbgcLmgcLqlmysdMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();

        //获取数据
        QueryWrapper<JjgFbgcLmgcLqlmysd> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        wrapper.like("htd",htd);
        wrapper.like("fbgc",fbgc);
        wrapper.orderByAsc("zh");
        List<JjgFbgcLmgcLqlmysd> data = jjgFbgcLmgcLqlmysdMapper.selectList(wrapper);

        String sffl = data.get(0).getSffl();
        //鉴定表要存放的路径
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"12沥青路面压实度.xlsx");
        if (data == null || data.size()==0){
            return;
        }else {
            File fdir = new File(filepath + File.separator + proname + File.separator + htd);
            if (!fdir.exists()) {
                //创建文件根目录
                fdir.mkdirs();
            }
            File directory = new File("service-system/src/main/resources");
            String reportPath = directory.getCanonicalPath();
            if (sffl.equals("否")) {
                String path = reportPath + "/static/沥青路面压实度.xlsx";
                Files.copy(Paths.get(path), new FileOutputStream(f));
                FileInputStream out = new FileInputStream(f);
                wb = new XSSFWorkbook(out);

                List<JjgFbgcLmgcLqlmysd> zxzfdata = new ArrayList<>();
                List<JjgFbgcLmgcLqlmysd> zxyfdata = new ArrayList<>();
                List<JjgFbgcLmgcLqlmysd> sdzfdata = new ArrayList<>();
                List<JjgFbgcLmgcLqlmysd> sdyfdata = new ArrayList<>();
                for(JjgFbgcLmgcLqlmysd ysd : data){
                    //沥青路面压实度左幅  将所有左幅的数据收集，包含路，桥，隧道，不包含匝道，连接线，连接线隧道，并且也不分左右幅
                    if (ysd.getZh().substring(0,1).equals("Z")){
                        zxzfdata.add(ysd);
                    }
                    if (ysd.getZh().substring(0,1).equals("Y")){
                        zxyfdata.add(ysd);
                    }
                    //隧道和桥的数据 是放在隧道左右幅的工作簿中，分左右幅
                    if (ysd.getLqs().contains("隧") && ysd.getZh().substring(0,1).equals("Z") || ysd.getLqs().contains("桥")){
                        sdzfdata.add(ysd);
                    }
                    if (ysd.getLqs().contains("隧") && ysd.getZh().substring(0,1).equals("Y") || ysd.getLqs().contains("桥")){
                        sdyfdata.add(ysd);
                    }
                }
                //匝道数据，是不分左右福的，
                List<JjgFbgcLmgcLqlmysd> zddata = jjgFbgcLmgcLqlmysdMapper.selectzd(proname,htd,fbgc);
                List<JjgFbgcLmgcLqlmysd> ljxdata = jjgFbgcLmgcLqlmysdMapper.selectljx(proname,htd,fbgc);
                List<JjgFbgcLmgcLqlmysd> ljxsddata = jjgFbgcLmgcLqlmysdMapper.selectljxsd(proname,htd,fbgc);

                //沥青路面压实度主线左幅
                lqlmysdzx(wb, zxzfdata, "沥青路面压实度左幅", "路面面层（主线左幅）");
                //沥青路面压实度主线右幅
                lqlmysdzx(wb, zxyfdata, "沥青路面压实度右幅", "路面面层（主线右幅）");
                //沥青路面压实度隧道左幅
                lqlmysdOther(wb, sdzfdata, "隧道左幅", "隧道路面");
                //沥青路面压实度隧道右幅
                lqlmysdOther(wb, sdyfdata, "隧道右幅", "隧道路面");
                //沥青路面压实度匝道
                lqlmysdOther(wb, zddata, "沥青路面压实度匝道", "匝道路面");
                //沥青路面压实度连接线
                lqlmysdOther(wb, ljxdata, "连接线", "连接线");
                //沥青路面压实度连接线隧道
                lqlmysdOther(wb, ljxsddata, "连接线隧道", "连接线隧道");
                for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                    if (shouldBeCalculate(wb.getSheetAt(j))) {
                        calculateSheet(wb.getSheetAt(j));
                        JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
                    }
                }
                getTunnelTotal(wb.getSheet("隧道右幅"));
                getTunnelTotal(wb.getSheet("隧道左幅"));
                JjgFbgcCommonUtils.updateFormula(wb, wb.getSheet("隧道右幅"));
                JjgFbgcCommonUtils.updateFormula(wb, wb.getSheet("隧道左幅"));
                //删除空表
                JjgFbgcCommonUtils.deleteEmptySheets(wb);

                FileOutputStream fileOut = new FileOutputStream(f);
                wb.write(fileOut);
                fileOut.flush();
                fileOut.close();
                out.close();

            } else if (sffl.equals("是")){
                String path = reportPath + "/static/沥青路面压实度-第二版.xlsx";
                Files.copy(Paths.get(path), new FileOutputStream(f));
                FileInputStream out = new FileInputStream(f);
                wb = new XSSFWorkbook(out);

                List<JjgFbgcLmgcLqlmysd> zxzfsmcdata = jjgFbgcLmgcLqlmysdMapper.selectzxsmc(proname,htd,fbgc,sffl,"上面层");
                List<JjgFbgcLmgcLqlmysd> zxyfsmcdata = jjgFbgcLmgcLqlmysdMapper.selectzxyfsmc(proname,htd,fbgc,sffl,"上面层");

                List<JjgFbgcLmgcLqlmysd> sdzfsmcdata = jjgFbgcLmgcLqlmysdMapper.selectsdzfsmc(proname,htd,fbgc,sffl,"上面层");
                List<JjgFbgcLmgcLqlmysd> sdyfsmcdata = jjgFbgcLmgcLqlmysdMapper.selectsdyfsmc(proname,htd,fbgc,sffl,"上面层");
                List<JjgFbgcLmgcLqlmysd> zxzfzxmcdata = jjgFbgcLmgcLqlmysdMapper.selectzxzxmc(proname,htd,fbgc,sffl);
                List<JjgFbgcLmgcLqlmysd> zxyfzxmcdata = jjgFbgcLmgcLqlmysdMapper.selectzxyfzxmc(proname,htd,fbgc,sffl);

                List<JjgFbgcLmgcLqlmysd> sdzfzxmcdata = jjgFbgcLmgcLqlmysdMapper.selectsdzfzxmc(proname,htd,fbgc,sffl);
                List<JjgFbgcLmgcLqlmysd> sdyfzxmcdata = jjgFbgcLmgcLqlmysdMapper.selectsdyfzxmc(proname,htd,fbgc,sffl);

                List<JjgFbgcLmgcLqlmysd> zdsmcdata = jjgFbgcLmgcLqlmysdMapper.selectzdsmc(proname,htd,fbgc);
                List<JjgFbgcLmgcLqlmysd> zdsxmcdata = jjgFbgcLmgcLqlmysdMapper.selectzdsxmc(proname,htd,fbgc);

                List<JjgFbgcLmgcLqlmysd> ljxsmcdata = jjgFbgcLmgcLqlmysdMapper.selectljxsmc(proname,htd,fbgc);
                List<JjgFbgcLmgcLqlmysd> ljxzxmcdata = jjgFbgcLmgcLqlmysdMapper.selectljxzxmc(proname,htd,fbgc);

                List<JjgFbgcLmgcLqlmysd> ljxsdsmcdata = jjgFbgcLmgcLqlmysdMapper.selectljxsdsmc(proname,htd,fbgc);
                List<JjgFbgcLmgcLqlmysd> ljxsdzxmcdata = jjgFbgcLmgcLqlmysdMapper.selectljxsdzxmc(proname,htd,fbgc);

                //沥青路面压实度主线左幅上面层
                lqlmysdzxsmc(wb,zxzfsmcdata,"沥青路面压实度左幅-上面层","路面面层（主线左幅）");

                //沥青路面压实度主线左幅中下面层
                lqlmysdzxzxmc(wb,zxzfzxmcdata,"沥青路面压实度左幅-中下面层","路面面层（主线左幅）");

                //沥青路面压实度主线右幅上面层
                lqlmysdzxsmc(wb,zxyfsmcdata,"沥青路面压实度右幅-上面层","路面面层（主线右幅）");

                //沥青路面压实度主线右幅中下面层
                lqlmysdzxzxmc(wb,zxyfzxmcdata,"沥青路面压实度右幅-中下面层","路面面层（主线右幅）");

                //沥青路面压实度隧道左幅上面层
                lqlmysdsdsmc(wb,sdzfsmcdata,"隧道左幅-上面层","隧道路面");

                //沥青路面压实度隧道左幅中下面层
                lqlmysdzxmc(wb,sdzfzxmcdata,"隧道左幅-中下面层","隧道路面");

                //沥青路面压实度隧道右幅上面层
                lqlmysdsdsmc(wb,sdyfsmcdata,"隧道右幅-上面层","隧道路面");

                //沥青路面压实度隧道右幅中下面层
                lqlmysdzxmc(wb,sdyfzxmcdata,"隧道右幅-中下面层","隧道路面");

                //沥青路面压实度匝道上面层
                lqlmysdsdsmc(wb,zdsmcdata,"沥青路面压实度匝道-上面层","匝道路面");

                //沥青路面压实度匝道中下面层
                lqlmysdzxmc(wb,zdsxmcdata,"沥青路面压实度匝道-中下面层","匝道路面");

                //沥青路面压实度连接线上面层
                lqlmysdsdsmc(wb,ljxsmcdata,"连接线-上面层","连接线");

                //沥青路面压实度连接线中下面层
                lqlmysdzxmc(wb,ljxzxmcdata,"连接线-中下面层","连接线");

                //沥青路面压实度连接线隧道上面层
                lqlmysdsdsmc(wb,ljxsdsmcdata,"连接线隧道-上面层","连接线隧道");

                //沥青路面压实度连接线隧道中下面层
                lqlmysdzxmc(wb,ljxsdzxmcdata,"连接线隧道-中下面层","连接线隧道");

                for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                    if (shouldBeCalculate(wb.getSheetAt(j))) {
                        calculateSheet(wb.getSheetAt(j));
                        JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
                    }
                }
                getTunnelTotal(wb.getSheet("隧道右幅-上面层"));
                getTunnelTotal(wb.getSheet("隧道右幅-中下面层"));
                getTunnelTotal(wb.getSheet("隧道左幅-上面层"));
                getTunnelTotal(wb.getSheet("隧道左幅-中下面层"));
                JjgFbgcCommonUtils.updateFormula(wb, wb.getSheet("隧道右幅-上面层"));
                JjgFbgcCommonUtils.updateFormula(wb, wb.getSheet("隧道右幅-中下面层"));
                JjgFbgcCommonUtils.updateFormula(wb, wb.getSheet("隧道左幅-上面层"));
                JjgFbgcCommonUtils.updateFormula(wb, wb.getSheet("隧道左幅-中下面层"));
                //删除空表
                JjgFbgcCommonUtils.deleteEmptySheets(wb);

                FileOutputStream fileOut = new FileOutputStream(f);
                wb.write(fileOut);
                fileOut.flush();
                fileOut.close();
                out.close();
            }

            wb.close();
        }
    }

    /**
     *
     * @param wb
     * @param data
     * @param sheetname
     * @param fbgcname
     */
    private void lqlmysdzxmc(XSSFWorkbook wb, List<JjgFbgcLmgcLqlmysd> data, String sheetname, String fbgcname) {
        if (data.size() > 0 && data != null){
            createZXZTable(wb,getZXZtableNum(data.size()),sheetname);
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
            XSSFSheet sheet = wb.getSheet(sheetname);
            String type = data.get(0).getLqs();
            String zh = data.get(0).getZh();
            sheet.getRow(1).getCell(2).setCellValue(data.get(0).getProname());
            sheet.getRow(1).getCell(9).setCellValue(data.get(0).getHtd());
            sheet.getRow(2).getCell(2).setCellValue(data.get(0).getFbgc());
            sheet.getRow(2).getCell(9).setCellValue(fbgcname);
            sheet.getRow(3).getCell(2).setCellValue(data.get(0).getLmlx());
            String date = simpleDateFormat.format(data.get(0).getJcsj());
            sheet.getRow(3).getCell(9).setCellValue(date);
            int index = 6;

            for(int i =0; i < data.size(); i++){
                if (zh.equals(data.get(i).getZh())){
                    boolean[] flag = new boolean[2];
                    for(int j = 0; j < 2; j++){
                        if(i+j < data.size()){
                            if("中面层".equals(data.get(i+j).getCw())){
                                sheet.getRow(index).getCell(2).setCellValue("中面层");
                                fillsdsmcCommonCellData(sheet, index+0, data.get(i+j));
                                flag[0] = true;
                            }
                            else if("下面层".equals(data.get(i+j).getCw())){
                                sheet.getRow(index+1).getCell(2).setCellValue("下面层");
                                fillsdsmcCommonCellData(sheet, index+1, data.get(i+j));
                                flag[1] = true;
                            }
                        }
                        else{
                            break;
                        }
                    }
                    for (int j = 0; j < flag.length; j++) {
                        if(!flag[j]){
                            fillCommonCell_Null_Data(sheet, index+j);
                        }
                    }
                    i++;
                    index+=2;
                    sheet.addMergedRegion(new CellRangeAddress(index-2, index-1, 0, 0));
                    sheet.getRow(index-2).getCell(0).setCellValue(zh);
                    sheet.addMergedRegion(new CellRangeAddress(index-2, index-1, 1, 1));
                    sheet.getRow(index-2).getCell(1).setCellValue(data.get(i).getQywz());
                }
                else{
                    type = data.get(i).getLqs();
                    zh = data.get(i).getZh();
                    i--;
                }
            }

        }



    }

    /**
     * 隧道上面层
     * @param wb
     * @param data
     * @param sheetname
     * @param fbgcname
     */
    private void lqlmysdsdsmc(XSSFWorkbook wb, List<JjgFbgcLmgcLqlmysd> data, String sheetname, String fbgcname) {
        if (data.size() > 0 && data != null){
            createZXZTable(wb,getZXZtableNum(data.size()),sheetname);
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
            XSSFSheet sheet = wb.getSheet(sheetname);
            String type = data.get(0).getLqs();
            String zh = data.get(0).getZh();
            sheet.getRow(1).getCell(2).setCellValue(data.get(0).getProname());
            sheet.getRow(1).getCell(9).setCellValue(data.get(0).getHtd());
            sheet.getRow(2).getCell(2).setCellValue(data.get(0).getFbgc());
            sheet.getRow(2).getCell(9).setCellValue(fbgcname);
            sheet.getRow(3).getCell(2).setCellValue(data.get(0).getLmlx());
            String date = simpleDateFormat.format(data.get(0).getJcsj());
            sheet.getRow(3).getCell(9).setCellValue(date);
            int index = 6;
            for(int i =0; i < data.size(); i++){
                if (zh.equals(data.get(i).getZh())){
                    if("上面层".equals(data.get(i).getCw())){
                        sheet.getRow(index).getCell(2).setCellValue("上面层");
                        fillsdsmcCommonCellData(sheet, index, data.get(i));
                    }
                    sheet.getRow(index).getCell(0).setCellValue(zh);
                    sheet.getRow(index).getCell(1).setCellValue(data.get(i).getQywz());
                    index++;
                }
                else{
                    type = data.get(i).getLqs();
                    zh = data.get(i).getZh();
                    i--;
                }
            }
        }

    }

    /**
     *
     * @param sheet
     * @param index
     * @param row
     */
    private void fillsdsmcCommonCellData(XSSFSheet sheet, int index, JjgFbgcLmgcLqlmysd row) {
        sheet.getRow(index).getCell(3).setCellValue(Double.parseDouble(row.getGzsjzl()));
        sheet.getRow(index).getCell(4).setCellValue(Double.parseDouble(row.getSjszzl()));
        sheet.getRow(index).getCell(5).setCellValue(Double.parseDouble(row.getSjbgzl()));
        sheet.getRow(index).getCell(7).setCellValue(Double.parseDouble(row.getSysbzmd()));
        sheet.getRow(index).getCell(8).setCellValue(Double.parseDouble(row.getZdllmd()));
        sheet.getRow(index).getCell(11).setCellValue(Double.parseDouble(row.getSysbzmdgdz()));
        sheet.getRow(index).getCell(12).setCellValue(Double.parseDouble(row.getZdllmdgdz()));
        sheet.getRow(sheet.getLastRowNum()-1).getCell(2).setCellValue(Double.parseDouble(row.getSysbzmdgdz()));
        sheet.getRow(sheet.getLastRowNum()-1).getCell(8).setCellValue(Double.parseDouble(row.getZdllmdgdz()));
    }


    /**
     * 分离
     * 沥青路面压实度主线左幅中下面层
     * @param wb
     * @param data
     * @param sheetname
     * @param fbgcname
     */
    private void lqlmysdzxzxmc(XSSFWorkbook wb, List<JjgFbgcLmgcLqlmysd> data, String sheetname, String fbgcname) {
        if (data.size() > 0 && data != null){
            createZXZTable(wb,getZXZtableNum(data.size()),sheetname);
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
            XSSFSheet sheet = wb.getSheet(sheetname);
            String type = data.get(0).getLqs();
            String zh = data.get(0).getZh();
            sheet.getRow(1).getCell(2).setCellValue(data.get(0).getProname());
            sheet.getRow(1).getCell(9).setCellValue(data.get(0).getHtd());
            sheet.getRow(2).getCell(2).setCellValue(data.get(0).getFbgc());
            sheet.getRow(2).getCell(9).setCellValue(fbgcname);
            sheet.getRow(3).getCell(2).setCellValue(data.get(0).getLmlx());
            String date = simpleDateFormat.format(data.get(0).getJcsj());
            sheet.getRow(3).getCell(9).setCellValue(date);
            int index = 6;

            for(int i =0; i < data.size(); i++){
                if (zh.equals(data.get(i).getZh())){
                    if(type.contains("隧道")){
                        i+=1;
                        index += 2;
                        sheet.addMergedRegion(new CellRangeAddress(index-2, index-1, 0, 1));
                        sheet.getRow(index-2).getCell(0).setCellValue(zh+type);
                        sheet.addMergedRegion(new CellRangeAddress(index-2, index-1, 2, 8));
                        sheet.getRow(index-2).getCell(2).setCellValue("见隧道路面压实度鉴定表");

                    }
                    else{
                        boolean[] flag = new boolean[2];
                        for(int j = 0; j < 2; j++){
                            if(i+j < data.size()){
                                if("中面层".equals(data.get(i+j).getCw())){
                                    sheet.getRow(index).getCell(2).setCellValue("中面层");
                                    fillzxsmcCommonCellData(sheet, index, data.get(i+j));
                                    flag[0] = true;
                                }else if("下面层".equals(data.get(i+j).getCw())){
                                    sheet.getRow(index+1).getCell(2).setCellValue("下面层");
                                    fillzxsmcCommonCellData(sheet, index+1, data.get(i+j));
                                    flag[1] = true;
                                }
                            }
                            else{
                                break;
                            }
                        }
                        for (int j = 0; j < flag.length; j++) {
                            if(!flag[j]){
                                fillCommonCell_Null_Data(sheet, index+j);
                            }
                        }
                        i++;
                        index+=2;
                        sheet.addMergedRegion(new CellRangeAddress(index-2, index-1, 0, 0));
                        sheet.getRow(index-2).getCell(0).setCellValue(zh);
                        sheet.addMergedRegion(new CellRangeAddress(index-2, index-1, 1, 1));
                        sheet.getRow(index-2).getCell(1).setCellValue(data.get(i).getQywz());
                    }
                }
                else{
                    type = data.get(i).getLqs();
                    zh = data.get(i).getZh();
                    i--;
                }

            }

        }

    }


    /**
     * 分离
     * 沥青路面压实度主线左幅上面层
     * @param wb
     * @param data
     * @param sheetname
     * @param fbgcname
     */
    private void lqlmysdzxsmc(XSSFWorkbook wb, List<JjgFbgcLmgcLqlmysd> data, String sheetname, String fbgcname) {
        if (data.size() > 0 && data != null){
            createZXZTable(wb,getZXZtableNum(data.size()),sheetname);
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
            XSSFSheet sheet = wb.getSheet(sheetname);
            String type = data.get(0).getLqs();
            String zh = data.get(0).getZh();
            sheet.getRow(1).getCell(2).setCellValue(data.get(0).getProname());
            sheet.getRow(1).getCell(9).setCellValue(data.get(0).getHtd());
            sheet.getRow(2).getCell(2).setCellValue(data.get(0).getFbgc());
            sheet.getRow(2).getCell(9).setCellValue(fbgcname);
            sheet.getRow(3).getCell(2).setCellValue(data.get(0).getLmlx());
            String date = simpleDateFormat.format(data.get(0).getJcsj());
            sheet.getRow(3).getCell(9).setCellValue(date);
            int index = 6;


            for(int i =0; i < data.size(); i++){

                if (zh.equals(data.get(i).getZh())){
                    if(type.contains("隧道") ){
                        //i++;
                        sheet.addMergedRegion(new CellRangeAddress(index, index, 0, 1));
                        sheet.getRow(index).getCell(0).setCellValue(zh+type);
                        sheet.addMergedRegion(new CellRangeAddress(index, index, 2, 8));
                        sheet.getRow(index).getCell(2).setCellValue("见隧道路面压实度鉴定表");
                        index ++;
                    }
                    else{
                        if("上面层".equals(data.get(i).getCw())){
                            sheet.getRow(index).getCell(2).setCellValue("上面层");
                            fillzxsmcCommonCellData(sheet, index, data.get(i));
                        }
                        sheet.getRow(index).getCell(0).setCellValue(zh);
                        sheet.getRow(index).getCell(1).setCellValue(data.get(i).getQywz());
                        index++;
                    }
                }
                else{
                    type = data.get(i).getLqs();
                    zh = data.get(i).getZh();
                    i--;
                }
            }
        }

    }

    /**
     * 分离
     * 上面层
     * @param sheet
     * @param index
     * @param row
     */
    private void fillzxsmcCommonCellData(XSSFSheet sheet, int index, JjgFbgcLmgcLqlmysd row) {
        sheet.getRow(index).getCell(3).setCellValue(Double.parseDouble(row.getGzsjzl()));
        sheet.getRow(index).getCell(4).setCellValue(Double.parseDouble(row.getSjszzl()));
        sheet.getRow(index).getCell(5).setCellValue(Double.parseDouble(row.getSjbgzl()));
        sheet.getRow(index).getCell(7).setCellValue(Double.parseDouble(row.getSysbzmd()));
        sheet.getRow(index).getCell(8).setCellValue(Double.parseDouble(row.getZdllmd()));
        sheet.getRow(index).getCell(11).setCellValue(Double.parseDouble(row.getSysbzmdgdz())-1);
        sheet.getRow(index).getCell(12).setCellValue(Double.parseDouble(row.getZdllmdgdz())-1);
        sheet.getRow(sheet.getLastRowNum()-1).getCell(2).setCellValue(Double.parseDouble(row.getSysbzmdgdz()));
        sheet.getRow(sheet.getLastRowNum()-1).getCell(8).setCellValue(Double.parseDouble(row.getZdllmdgdz()));

    }

    /**
     * 不分离上面层
     * 将每个隧道的数据汇总
     * @param sheet
     */
    private void getTunnelTotal(XSSFSheet sheet){
        XSSFRow row = null;
        boolean flag = false;
        XSSFRow startrow = null;
        XSSFRow endrow = null;
        String name = "";
        for (int i = sheet.getFirstRowNum(); i <= sheet
                .getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if(flag && row.getCell(0) != null && !"".equals(row.getCell(0).toString())){
                if(!name.equals(row.getCell(0).toString().replaceAll("[^\u4E00-\u9FA5]", "")) && !"".equals(name)){
                    endrow = sheet.getRow(i-1);
                    startrow.createCell(15).setCellFormula("COUNT("
                            +startrow.getCell(9).getReference()+":"
                            +endrow.getCell(9).getReference()+")");//=COUNT(J7:J21)
                    startrow.createCell(16).setCellFormula("COUNTIF("
                            +startrow.getCell(13).getReference()+":"
                            +endrow.getCell(13).getReference()+",\"√\")");//=COUNTIF(N7:N22,"√")
                    startrow.createCell(17).setCellFormula(startrow.getCell(16).getReference()+"*100/"
                            +startrow.getCell(15).getReference());
                    name = row.getCell(0).toString().replaceAll("[^\u4E00-\u9FA5]", "");
                    startrow = row;
                }
                if("".equals(name)){
                    name = row.getCell(0).toString().replaceAll("[^\u4E00-\u9FA5]", "");
                    startrow = row;
                }

            }
            if ("桩号".equals(row.getCell(0).toString())) {
                flag = true;
                i++;
            }
            if ("评定".equals(row.getCell(0).toString())) {
                break;
            }
        }
    }

    /**
     * 不分离上面层
     * 判断此sheet是否需要进行计算
     * @param sheet
     * @return
     */
    public boolean shouldBeCalculate(XSSFSheet sheet) {
        String title = null;
        title = sheet.getRow(0).getCell(0).getStringCellValue();
        if (title.endsWith("鉴定表")) {
            return true;
        }
        return false;
    }

    /**
     * 不分离上面层
     * @param sheet
     */
    private void calculateSheet(XSSFSheet sheet) {
        XSSFRow row = null;

        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }

            // 计算表格中间的数据
            fillBodyRow(row);

            // 完成表格尾部数据的计算及填写
            if ("评定".equals(row.getCell(0).toString())) {
                fillBottomConclusion(sheet, i);
                // 文件的所有计算都已完成，跳出循环，结束剩下2行的读取
                break;
            }
        }

        /*
         * 统计最大值最小值，生成报告表格的时候要用
         */
        row = sheet.getRow(sheet.getPhysicalNumberOfRows()-3);
        row.createCell(15).setCellValue("最大值");
        row.createCell(16).setCellFormula("ROUND("+
                "MAX("+sheet.getRow(6).getCell(9).getReference()+":"+
                sheet.getRow(sheet.getPhysicalNumberOfRows()-4).getCell(9).getReference()+"),1)");//=MAX(J7:J57)
        row.createCell(17).setCellFormula("ROUND("+"MAX("+sheet.getRow(6).getCell(10).getReference()+":"+
                sheet.getRow(sheet.getPhysicalNumberOfRows()-4).getCell(10).getReference()+"),1)");//=MAX(J7:J57)

        row = sheet.getRow(sheet.getPhysicalNumberOfRows()-2);
        row.createCell(15).setCellValue("最小值");
        row.createCell(16).setCellFormula("ROUND("+"MIN("+sheet.getRow(6).getCell(9).getReference()+":"+
                sheet.getRow(sheet.getPhysicalNumberOfRows()-4).getCell(9).getReference()+"),1)");//=MIN(J7:J57)
        row.createCell(17).setCellFormula("ROUND("+"MIN("+sheet.getRow(6).getCell(10).getReference()+":"+
                sheet.getRow(sheet.getPhysicalNumberOfRows()-4).getCell(10).getReference()+"),1)");//=MIN(J7:J57)

    }

    /**
     * 不分离上面层
     * 计算表格尾部汇总结果
     * @param sheet
     * @param i
     */
    public void fillBottomConclusion(XSSFSheet sheet, int i) {
        /*
         * 表格尾部有3行，由于数据间有联系，所以不能逐行计算 计算顺序为：平均值，监测点数->标准差->代表值->结论->合格点数->合格率
         */
        // 定义3个局部变量，表示表尾的1,2,3行
        XSSFRow row1, row2, row3;
        row1 = sheet.getRow(i);
        row2 = sheet.getRow(i + 1);
        row3 = sheet.getRow(i + 2);

        // 试验室标准密度控制->平均值
        row1.getCell(4).setCellType(CellType.NUMERIC);
        row1.getCell(4).setCellFormula(
                "AVERAGE(" + "J7" + ":" + "J" + row1.getRowNum() + ")");// AVERAGE(J7:J24)
        // 试验室标准密度控制->监测点数
        row1.getCell(6).setCellType(CellType.NUMERIC);
        row1.getCell(6).setCellFormula(
                "COUNT(" + "J7" + ":" + "J" + row1.getRowNum() + ")");// COUNT(J7:J24)

        // 最大理论密度控制->平均值
        row1.getCell(10).setCellType(CellType.NUMERIC);
        row1.getCell(10).setCellFormula(
                "AVERAGE(" + "K7" + ":" + "K" + row1.getRowNum() + ")");// AVERAGE(K7:K24)
        // 最大理论密度控制->监测点数
        row1.getCell(12).setCellType(CellType.NUMERIC);
        row1.getCell(12).setCellFormula(
                "COUNT(" + "K7" + ":" + "K" + row1.getRowNum() + ")");// COUNT(K7:K24)

        // 标准差
        row2.getCell(4).setCellType(CellType.NUMERIC);
        row2.getCell(4).setCellFormula(
                "STDEV(" + "J7" + ":" + "J" + (row2.getRowNum() - 1) + ")");// =STDEV(J7:J24)
        // 标准差
        row2.getCell(10).setCellType(CellType.NUMERIC);
        row2.getCell(10).setCellFormula(
                "STDEV(" + "K7" + ":" + "K" + (row2.getRowNum() - 1) + ")");// =STDEV(K7:K24)

        // 代表值
        row3.getCell(4).setCellType(CellType.NUMERIC);
        row3.getCell(4).setCellFormula(
                "ROUND(" + row1.getCell(4).getReference() + "-("
                        + row2.getCell(4).getReference() + "*VLOOKUP("
                        + row1.getCell(6).getReference()
                        + ",保证率系数!$A:$D,3)),1)");// ROUND(E40-(E41*VLOOKUP(G40,保证率系数!$A:$D,3,)),1)
        row3.getCell(10).setCellType(CellType.NUMERIC);
        row3.getCell(10).setCellFormula(
                "ROUND(" + row1.getCell(10).getReference() + "-("
                        + row2.getCell(10).getReference() + "*VLOOKUP("
                        + row1.getCell(12).getReference()
                        + ",保证率系数!$A:$D,3)),1)");// =ROUND(K40-(K41*VLOOKUP(M40,保证率系数!$A:$D,3,)),1)
        // 右下角结论
        row1.getCell(14).setCellType(CellType.NUMERIC);
        row1.getCell(14).setCellFormula(
                "IF((" + row3.getCell(4).getReference() + ">="
                        + row2.getCell(2).getReference() + ")*AND("
                        + row3.getCell(10).getReference() + ">="
                        + row2.getCell(8).getReference() + "),\"合格\",\"不合格\")");// IF((E42>=C41)*AND(K42>=I41),"合格","不合格")
        // 计算右侧结论
        fillRightConclusion(sheet);
        // 合格点数 =COUNTIF(N7:N24,"√")
        row2.getCell(6).setCellType(CellType.NUMERIC);
        row2.getCell(6).setCellFormula(
                "COUNTIF(" + "N7" + ":" + row1.getCell(13).getReference()
                        + ",\"√\")");
        row2.getCell(12).setCellFormula(
                "COUNTIF(" + "N7" + ":" + row1.getCell(13).getReference()
                        + ",\"√\")");
        // 合格率 =G41/G40*100
        row3.getCell(6).setCellFormula(
                row2.getCell(6).getReference() + "/"
                        + row1.getCell(6).getReference() + "*100");
        row3.getCell(12).setCellFormula(
                row2.getCell(12).getReference() + "/"
                        + row1.getCell(12).getReference() + "*100");
    }

    /**
     * 不分离上面层
     * =IF(O58="合格",IF((J13>=L13)*(K13>M13),"√",""),"")
     * =IF(J13="","",IF(N13="","×",""))
     * =IF((E60>=C59)*AND(K60>=I59),"合格","不合格")
     * 根据表格的一些汇总数据，完成右边的结论计算
     * @param sheet
     */
    public void fillRightConclusion(XSSFSheet sheet) {
        XSSFRow row = null;
        for (int i = sheet.getFirstRowNum(); i <= sheet
                .getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            // 判断前面的数据有值时，才进行结论的计算
            if (row.getCell(9).getCellType() == Cell.CELL_TYPE_FORMULA
                    && row.getCell(10).getCellType() == Cell.CELL_TYPE_FORMULA
                    && row.getCell(11).getCellType() == Cell.CELL_TYPE_NUMERIC
                    && row.getCell(12).getCellType() == Cell.CELL_TYPE_NUMERIC) {
                // 计算结论
                row.getCell(13).setCellFormula(
                        "IF(" + getLastCell(sheet).getReference()
                                + "=\"合格\",IF(("
                                + row.getCell(9).getReference() + ">="
                                + row.getCell(11).getReference() + ")*("
                                + row.getCell(10).getReference() + ">"
                                + row.getCell(12).getReference()
                                + "),\"√\",\"\"),\"\")");// =IF($O$40="合格",IF((J7>=L7)*(K7>M7),"√",""),"")
                row.getCell(14).setCellFormula(
                        "IF(" + row.getCell(9).getReference()
                                + "=\"\",\"\",IF("
                                + row.getCell(13).getReference()
                                + "=\"\",\"×\",\"\"))");// =IF(J13="","",IF(N13="","×",""))
            }
        }
    }

    /**
     * 不分离上面层
     * 得到最后一个单元格，计算右侧结论要使用
     * @param sheet
     * @return
     */
    public XSSFCell getLastCell(XSSFSheet sheet) {
        XSSFRow row = null;
        XSSFCell cell = null;
        row = sheet.getRow(sheet.getLastRowNum() - 2);
        cell = row.getCell(14);
        return cell;
    }

    /**
     * 不分离上面层
     *  计算表格中一行的数据
     * @param row
     */
    public void fillBodyRow(XSSFRow row) {
        if (row.getCell(3).getCellTypeEnum() == CellType.NUMERIC && row.getCell(4).getCellTypeEnum() == CellType.NUMERIC && row.getCell(5).getCellTypeEnum() == CellType.NUMERIC) {
            row.getCell(6).setCellType(CellType.NUMERIC);
            row.getCell(6).setCellFormula(
                    row.getCell(3).getReference() + "/" + "("
                            + row.getCell(5).getReference() + "-"
                            + row.getCell(4).getReference() + ")");// D7/(F7-E7)
        }
        if (row.getCell(7).getCellTypeEnum() == CellType.NUMERIC
                && row.getCell(8).getCellTypeEnum() == CellType.NUMERIC) {
            // 试验室
            row.getCell(9).setCellType(CellType.NUMERIC);
            row.getCell(9).setCellFormula("ROUND("+
                    row.getCell(6).getReference() + "/"
                    + row.getCell(7).getReference() + "*" + "100,1)");// G7/H7*100
            // 最大理论
            row.getCell(10).setCellType(CellType.NUMERIC);
            row.getCell(10).setCellFormula("ROUND("+
                    row.getCell(6).getReference() + "/"
                    + row.getCell(8).getReference() + "*" + "100,1)");// G7/I7*100
        }
    }


    /**
     * 不分离上面层
     *隧道 匝道 连接线 连接线隧道
     * @param wb
     * @param data
     * @param sheetname
     * @param fbgcname
     */
    private void lqlmysdOther(XSSFWorkbook wb, List<JjgFbgcLmgcLqlmysd> data, String sheetname, String fbgcname) {
        if (data.size() > 0 && data != null){
            createZXZTable(wb,getZXZtableNum(data.size()),sheetname);
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
            XSSFSheet sheet = wb.getSheet(sheetname);
            String type = data.get(0).getLqs();
            String zh = data.get(0).getZh();
            sheet.getRow(1).getCell(2).setCellValue(data.get(0).getProname());
            sheet.getRow(1).getCell(9).setCellValue(data.get(0).getHtd());
            sheet.getRow(2).getCell(2).setCellValue(data.get(0).getFbgc());
            sheet.getRow(2).getCell(9).setCellValue(fbgcname);
            sheet.getRow(3).getCell(2).setCellValue(data.get(0).getLmlx());
            String date = simpleDateFormat.format(data.get(0).getJcsj());
            sheet.getRow(3).getCell(9).setCellValue(date);
            int index = 6;
            for (int i = 0; i < data.size(); i++) {
                if (zh.equals(data.get(i).getZh())) {
                    int count = 0;
                    for (int j = 0; j < 3; j++) {
                        if (i + j < data.size()) {
                            if ("上面层".equals(data.get(i + j).getCw()) && data.get(i).getZh().equals(data.get(i + j).getZh())) {
                                sheet.getRow(index).getCell(2).setCellValue(data.get(i + j).getCw());
                                fillsdCommonCellData(sheet, index, data.get(i + j));
                                count++;
                            }
                            if ("中面层".equals(data.get(i + j).getCw()) && data.get(i).getZh().equals(data.get(i + j).getZh())) {
                                sheet.getRow(index + 1).getCell(2).setCellValue(data.get(i + j).getCw());
                                fillsdCommonCellData(sheet, index + 1, data.get(i + j));
                                count++;
                            }
                            if ("下面层".equals(data.get(i + j).getCw()) && data.get(i).getZh().equals(data.get(i + j).getZh())) {
                                sheet.getRow(index + 2).getCell(2).setCellValue(data.get(i + j).getCw());
                                fillsdCommonCellData(sheet, index + 2, data.get(i + j));
                                count++;
                            }
                        } else {
                            break;
                        }
                    }
                    i += 2;
                    index+=count;
                    if(count > 1){
                        sheet.addMergedRegion(new CellRangeAddress(index-count, index-1, 0, 0));
                    }
                    if (data.get(i).getLqs().equals("路") || data.get(i).getLqs().contains("隧") || data.get(i).getLqs().contains("连接线") || data.get(i).getLqs().equals("桥")){
                        sheet.getRow(index-count).getCell(0).setCellValue(zh+type);

                    }else {
                        sheet.getRow(index-count).getCell(0).setCellValue(zh);
                    }
                    if(count > 1){
                        sheet.addMergedRegion(new CellRangeAddress(index-count, index-1, 1, 1));
                    }
                    //sheet.getRow(index-count).getCell(1).setCellValue(data.get(i)[8]);
                } else {
                    type = data.get(i).getLqs();
                    zh = data.get(i).getZh();
                    i--;
                }
            }
        }
    }

    /**
     * 不分离上面层
     * 沥青路面压实度主线
     * @param wb
     * @param data
     * @param sheetname
     * @param fbgcname
     */
    private void lqlmysdzx(XSSFWorkbook wb,List<JjgFbgcLmgcLqlmysd> data,String sheetname,String fbgcname) {
        if (data.size() > 0 && data != null){
            createZXZTable(wb,getZXZtableNum(data.size()),sheetname);
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
            XSSFSheet sheet = wb.getSheet(sheetname);
            String type = data.get(0).getLqs();
            String zh = data.get(0).getZh();
            sheet.getRow(1).getCell(2).setCellValue(data.get(0).getProname());
            sheet.getRow(1).getCell(9).setCellValue(data.get(0).getHtd());
            sheet.getRow(2).getCell(2).setCellValue(data.get(0).getFbgc());
            sheet.getRow(2).getCell(9).setCellValue(fbgcname);
            sheet.getRow(3).getCell(2).setCellValue(data.get(0).getLmlx());
            String date = simpleDateFormat.format(data.get(0).getJcsj());
            sheet.getRow(3).getCell(9).setCellValue(date);

            int index = 6;

            for(int i =0; i < data.size(); i++){
                if (zh.equals(data.get(i).getZh())){
                    if(type.contains("隧") || type.contains("桥")){
                        i+=2;
                        index += 3;
                        sheet.addMergedRegion(new CellRangeAddress(index-3, index-1, 0, 1));
                        sheet.getRow(index-3).getCell(0).setCellValue(zh+type);
                        sheet.addMergedRegion(new CellRangeAddress(index-3, index-1, 2, 8));
                        sheet.getRow(index-3).getCell(2).setCellValue("见隧道路面压实度鉴定表");

                    }else {
                        for(int j = 0; j < 3; j++){
                            if(i+j < data.size()){
                                if("上面层".equals(data.get(i+j).getCw())){
                                    sheet.getRow(index).getCell(2).setCellValue(data.get(i+j).getCw());
                                    fillLmzfCommonCellData(sheet, index, data.get(i+j));
                                    //flag[0] = true;
                                }
                                if("中面层".equals(data.get(i+j).getCw())){
                                    sheet.getRow(index+1).getCell(2).setCellValue(data.get(i+j).getCw());
                                    fillLmzfCommonCellData(sheet, index+1, data.get(i+j));
                                    //flag[1] = true;
                                }
                                if("下面层".equals(data.get(i+j).getCw())){
                                    sheet.getRow(index+2).getCell(2).setCellValue(data.get(i+j).getCw());
                                    fillLmzfCommonCellData(sheet, index+2, data.get(i+j));
                                    //flag[2] = true;
                                }
                            }else {
                                break;
                            }
                        }
                        i+=2;
                        index+=3;
                        sheet.addMergedRegion(new CellRangeAddress(index-3, index-1, 0, 0));
                        sheet.getRow(index-3).getCell(0).setCellValue(zh);
                        sheet.addMergedRegion(new CellRangeAddress(index-3, index-1, 1, 1));
                    }
                }else {
                    type = data.get(i).getLqs();
                    zh = data.get(i).getZh();
                    i--;
                }
            }

        }

    }

    /**
     * 不分离上面层
     * 沥青路面压实度左幅的层位没有数据时，以-显示
     * @param sheet
     * @param index
     */
    private void fillCommonCell_Null_Data(XSSFSheet sheet, int index) {
        int[] num = new int[]{2,3,4,5,7,8,9,10,11,12};
        for (int i = 0; i < num.length; i++) {
            if(sheet.getRow(index).getCell(num[i]) == null){
                sheet.getRow(index).createCell(num[i]);
            }
            sheet.getRow(index).getCell(num[i]).setCellValue("-");
        }
    }


    /**
     * 不分离上面层
     * @param sheet
     * @param index
     * @param zfdata
     */
    private void fillsdCommonCellData(XSSFSheet sheet, int index, JjgFbgcLmgcLqlmysd zfdata) {
        sheet.getRow(index).getCell(0).setCellValue(zfdata.getZh());
        sheet.getRow(index).getCell(1).setCellValue(zfdata.getQywz());
        if(!"".equals(zfdata.getGzsjzl()) && zfdata.getGzsjzl() != null){
            sheet.getRow(index).getCell(3).setCellValue(Double.parseDouble(zfdata.getGzsjzl()));
        }
        if(!"".equals(zfdata.getSjszzl()) && zfdata.getSjszzl()!=null){
            sheet.getRow(index).getCell(4).setCellValue(Double.parseDouble(zfdata.getSjszzl()));
        }
        if(!"".equals(zfdata.getSjbgzl()) && zfdata.getSjbgzl()!=null){
            sheet.getRow(index).getCell(5).setCellValue(Double.parseDouble(zfdata.getSjbgzl()));
        }
        if(!"".equals(zfdata.getSysbzmd()) && zfdata.getSysbzmd()!=null){
            sheet.getRow(index).getCell(7).setCellValue(Double.parseDouble(zfdata.getSysbzmd()));
        }

        if(!"".equals(zfdata.getZdllmd()) && zfdata.getZdllmd()!=null){
            sheet.getRow(index).getCell(8).setCellValue(Double.parseDouble(zfdata.getZdllmd()));
        }

        if(!"".equals(zfdata.getSysbzmdgdz()) && zfdata.getSysbzmdgdz()!=null){
            sheet.getRow(index).getCell(11).setCellValue(Double.parseDouble(zfdata.getSysbzmdgdz()));
        }

        if(!"".equals(zfdata.getZdllmdgdz()) && zfdata.getZdllmdgdz()!=null){
            sheet.getRow(index).getCell(12).setCellValue(Double.parseDouble(zfdata.getZdllmdgdz()));
        }
        sheet.getRow(sheet.getLastRowNum()-1).getCell(2).setCellValue(Double.parseDouble(zfdata.getSysbzmdgdz()));
        sheet.getRow(sheet.getLastRowNum()-1).getCell(8).setCellValue(Double.parseDouble(zfdata.getZdllmdgdz()));
    }


    /**
     * 不分离上面层
     * 沥青路面压实度左幅数据
     * @param sheet
     * @param index
     * @param zfdata
     */
    private void fillLmzfCommonCellData(XSSFSheet sheet, int index, JjgFbgcLmgcLqlmysd zfdata) {
        sheet.getRow(index).getCell(0).setCellValue(zfdata.getZh());
        sheet.getRow(index).getCell(1).setCellValue(zfdata.getQywz());
        //sheet.getRow(index).getCell(2).setCellValue(zfdata.getCw());
        if(!"".equals(zfdata.getGzsjzl()) && zfdata.getGzsjzl() != null){
            sheet.getRow(index).getCell(3).setCellValue(Double.parseDouble(zfdata.getGzsjzl()));
        }else {
            sheet.getRow(index).getCell(3).setCellValue("-");
        }
        if(!"".equals(zfdata.getSjszzl()) && zfdata.getSjszzl()!=null){
            sheet.getRow(index).getCell(4).setCellValue(Double.parseDouble(zfdata.getSjszzl()));
        }else {
            sheet.getRow(index).getCell(4).setCellValue("-");
        }
        if(!"".equals(zfdata.getSjbgzl()) && zfdata.getSjbgzl()!=null){
            sheet.getRow(index).getCell(5).setCellValue(Double.parseDouble(zfdata.getSjbgzl()));
        }else {
            sheet.getRow(index).getCell(5).setCellValue("-");
        }
        if(!"".equals(zfdata.getSysbzmd()) && zfdata.getSysbzmd()!=null){
            sheet.getRow(index).getCell(7).setCellValue(Double.parseDouble(zfdata.getSysbzmd()));
        }else {
            sheet.getRow(index).getCell(7).setCellValue("-");
        }

        if(!"".equals(zfdata.getZdllmd()) && zfdata.getZdllmd()!=null){
            sheet.getRow(index).getCell(8).setCellValue(Double.parseDouble(zfdata.getZdllmd()));
        }else {
            sheet.getRow(index).getCell(8).setCellValue("-");
        }

        if(!"".equals(zfdata.getSysbzmdgdz()) && zfdata.getSysbzmdgdz()!=null){
            sheet.getRow(index).getCell(11).setCellValue(Double.parseDouble(zfdata.getSysbzmdgdz()));
        }else {
            sheet.getRow(index).getCell(11).setCellValue("-");
        }

        if(!"".equals(zfdata.getZdllmdgdz()) && zfdata.getZdllmdgdz()!=null){
            sheet.getRow(index).getCell(12).setCellValue(Double.parseDouble(zfdata.getZdllmdgdz()));
        }else {
            sheet.getRow(index).getCell(12).setCellValue("-");
        }
        sheet.getRow(sheet.getLastRowNum()-1).getCell(2).setCellValue(Double.parseDouble(zfdata.getSysbzmdgdz()));
        sheet.getRow(sheet.getLastRowNum()-1).getCell(8).setCellValue(Double.parseDouble(zfdata.getZdllmdgdz()));
    }


    /**
     *不分离上面层
     * @param wb
     * @param tableNum
     */
    private void createZXZTable(XSSFWorkbook wb,int tableNum,String sheetname) {
        if(tableNum == 0){
            return;
        }
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            if(i < record-1){
                RowCopy.copyRows(wb, "沥青路面压实度", sheetname, 6, 23, (i - 1) * 18 + 24);
            }
            else{
                RowCopy.copyRows(wb, "沥青路面压实度", sheetname, 6, 20, (i - 1) * 18 + 24);
            }
        }
        if(record == 1){
            wb.getSheet(sheetname).shiftRows(22, 24, -1);
        }
        RowCopy.copyRows(wb, "source", sheetname, 0, 3,(record) * 18 + 3);
        //wb.getSheet("沥青路面压实度左幅").getRow((record) * 18 + 4).getCell(2).setCellValue(lab);
        //wb.getSheet("沥青路面压实度左幅").getRow((record) * 18 + 4).getCell(8).setCellValue(max);
        wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 14, 0,(record) * 18 + 5);


    }

    /**
     *不分离上面层
     * @param size
     * @return
     */
    private int getZXZtableNum(int size) {
        return size%18 <= 15 ? size/18+1 : size/18+2;
    }


    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "12沥青路面压实度.xlsx");
        if (!f.exists()) {
            return null;
        } else {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
            List<Map<String, Object> > jgmap = new ArrayList<>();
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                if (!wb.isSheetHidden(wb.getSheetIndex(wb.getSheetAt(j)))) {
                    XSSFSheet slSheet = wb.getSheetAt(j);
                    XSSFCell xmname = slSheet.getRow(1).getCell(2);//项目名
                    XSSFCell htdname = slSheet.getRow(1).getCell(9);//合同段名
                    Map map = new HashMap();
                    if (proname.equals(xmname.toString()) && htd.equals(htdname.toString())) {
                        //获取到最后一行
                        int lastRowNum = slSheet.getLastRowNum();
                        slSheet.getRow(lastRowNum-2).getCell(6).setCellType(CellType.STRING);//总点数
                        slSheet.getRow(lastRowNum-1).getCell(6).setCellType(CellType.STRING);//合格点数
                        slSheet.getRow(lastRowNum).getCell(6).setCellType(CellType.STRING);//合格率
                        slSheet.getRow(lastRowNum-1).getCell(2).setCellType(CellType.STRING);//合格率
                        slSheet.getRow(lastRowNum-2).getCell(16).setCellType(CellType.STRING);
                        slSheet.getRow(lastRowNum-1).getCell(16).setCellType(CellType.STRING);
                        slSheet.getRow(lastRowNum).getCell(4).setCellType(CellType.STRING);
                        double zds = Double.valueOf(slSheet.getRow(lastRowNum-2).getCell(6).getStringCellValue());
                        double hgds = Double.valueOf(slSheet.getRow(lastRowNum-1).getCell(6).getStringCellValue());
                        double hgl = Double.valueOf(slSheet.getRow(lastRowNum).getCell(6).getStringCellValue());
                        double gdz = Double.valueOf(slSheet.getRow(lastRowNum-1).getCell(2).getStringCellValue());
                        String zdsz1 = decf.format(zds);
                        String hgdsz1 = decf.format(hgds);
                        String hglz1 = df.format(hgl);
                        String gdz1 = df.format(gdz);
                        map.put("路面类型", wb.getSheetName(j));
                        map.put("规定值", gdz1);
                        map.put("检测点数", zdsz1);
                        map.put("合格点数", hgdsz1);
                        map.put("合格率", hglz1);
                        map.put("最大值", slSheet.getRow(lastRowNum-2).getCell(16).getStringCellValue());
                        map.put("最小值", slSheet.getRow(lastRowNum-1).getCell(16).getStringCellValue());
                        map.put("代表值", slSheet.getRow(lastRowNum).getCell(4).getStringCellValue());
                    }
                    jgmap.add(map);

                }
            }
            return jgmap;
        }
    }

    /**
     * 导出模板文件
     * @param response
     */
    @Override
    public void exportlqlmysd(HttpServletResponse response) {
        String fileName = "01沥青路面压实度实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcLmgcLqlmysdVo()).finish();
    }

    /**
     * 导入数据
     * @param file
     * @param commonInfoVo
     */
    @Override
    public void importlqlmysd(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcLmgcLqlmysdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcLmgcLqlmysdVo>(JjgFbgcLmgcLqlmysdVo.class) {
                                @Override
                                public void handle(List<JjgFbgcLmgcLqlmysdVo> dataList) {
                                    for(JjgFbgcLmgcLqlmysdVo lqlmysdVo: dataList)
                                    {
                                        JjgFbgcLmgcLqlmysd fbgcLmgcLqlmysd = new JjgFbgcLmgcLqlmysd();
                                        BeanUtils.copyProperties(lqlmysdVo,fbgcLmgcLqlmysd);
                                        fbgcLmgcLqlmysd.setCreatetime(new Date());
                                        fbgcLmgcLqlmysd.setProname(commonInfoVo.getProname());
                                        fbgcLmgcLqlmysd.setHtd(commonInfoVo.getHtd());
                                        fbgcLmgcLqlmysd.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcLmgcLqlmysdMapper.insert(fbgcLmgcLqlmysd);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }


    }

    @Override
    public int selectnum(String proname, String htd) {
        int selectnum = jjgFbgcLmgcLqlmysdMapper.selectnum(proname, htd);
        return selectnum;
    }

}
