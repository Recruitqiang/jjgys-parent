package glgc.jjgys.system.service.impl;
import com.alibaba.excel.metadata.BaseRowModel;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.common.excel.ExcelWriterFactroy;
import glgc.jjgys.model.projectvo.ljgc.*;
import glgc.jjgys.system.service.*;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.JjgFbgcUtils;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import static java.util.Arrays.asList;

/**
 * <p>
 *  路基工程服务类
 * </p>
 *
 * @author lhy
 * @since 2023-07-01
 */
@Service
public class JjgFbgcLjgcServiceImpl implements JjgFbgcLjgcService {

    @Resource
    private  JjgFbgcLjgcHdgqdService jjgFbgcLjgcHdgqdService;
    @Resource
    private  JjgFbgcLjgcHdjgccService jjgFbgcLjgcHdjgccService;
    @Resource
    private  JjgFbgcLjgcPsdmccService jjgFbgcLjgcPsdmccService;
    @Resource
    private  JjgFbgcLjgcPspqhdService jjgFbgcLjgcPspqhdService;
    @Resource
    private  JjgFbgcLjgcLjtsfysdHtService jjgFbgcLjgcLjtsfysdHtService;
    @Resource
    private  JjgFbgcLjgcLjtsfysdSlService jjgFbgcLjgcLjtsfysdSlService;
    @Resource
    private  JjgFbgcLjgcLjcjService jjgFbgcLjgcLjcjService;
    @Resource
    private  JjgFbgcLjgcLjwcService jjgFbgcLjgcLjwcService;
    @Resource
    private  JjgFbgcLjgcLjwcLcfService jjgFbgcLjgcLjwcLcfService;
    @Resource
    private  JjgFbgcLjgcLjbpService jjgFbgcLjgcLjbpService;
    @Resource
    private  JjgFbgcLjgcXqjgccService jjgFbgcLjgcXqjgccService;
    @Resource
    private  JjgFbgcLjgcXqgqdService jjgFbgcLjgcXqgqdService;
    @Resource
    private  JjgFbgcLjgcZddmccService jjgFbgcLjgcZddmccService;
    @Resource
    private  JjgFbgcLjgcZdgqdService jjgFbgcLjgcZdgqdService;








    /**
     * 导出分部工程数据模板文件
     * @param response
     */
    @Override
    public void exportljgc(HttpServletResponse response,String workpath) {


        String filepath = workpath;

        try {
            List<String> filenames=asList("08涵洞砼强度实测数据","09涵洞结构尺寸实测数据","04排水断面尺寸实测数据","05排水铺砌厚度实测数据","01路基压实度沉降实测数据",
                    "02路基弯沉(贝克曼梁法)实测数据","02路基弯沉(落锤法)实测数据","03路基边坡实测数据","07小桥结构尺寸实测数据","06小桥砼强度实测数据",
                    "11支挡断面尺寸实测数据","10支挡砼强度实测数据");
            List<String> sheetnames=asList( "原始数据","实测数据","实测数据","实测数据","实测数据","实测数据","实测数据","实测数据","实测数据","实测数据","实测数据","实测数据");
            List<BaseRowModel> models=asList( new JjgFbgcLjgcHdgqdVo(),new JjgFbgcLjgcHdjgccVo(),new JjgFbgcLjgcPsdmccVo(),new JjgFbgcLjgcPspqhdVo(), new JjgFbgcLjgcLjcjVo(),
                    new JjgFbgcLjgcLjwcVo(), new JjgFbgcLjgcLjwcLcfVo(),new JjgFbgcLjgcLjbpVo(), new JjgFbgcLjgcXqjgccVo(),new JjgFbgcLjgcXqgqdVo(),new JjgFbgcLjgcZddmccVo(),
                    new JjgFbgcLjgcZdgqdVo());
            List<ExcelWriterFactroy> writer=new ArrayList<>();
            for(int i=0;i<filenames.size();i++){
                File file=new File(filepath+"/"+filenames.get(i)+".xlsx");

                if(!file.exists()){
                    ExcelUtil.saveLocal(filepath+"/"+filenames.get(i),models.get(i),sheetnames.get(i));
                    System.out.println("生成表格"+file.getName());
                }



            }
            //处理灰土沙砾
            XSSFWorkbook workbook = new XSSFWorkbook();// 创建一个Excel文件
            XSSFCellStyle columnHeadStyle = JjgFbgcCommonUtils.tsfCellStyle(workbook);

            XSSFSheet sheet = workbook.createSheet("实测数据");// 创建一个Excel的Sheet
            sheet.setColumnWidth(0,28*256);
            List<String> checklist = Arrays.asList("路基压实度_灰土规定值","试验时间","桩号","结构层次",
                    "结构类型","最大干密度","标准砂密度","取样桩号及位置","试坑深度(cm)",
                    "锥体及基板和表面间砂质量（g）","灌砂前筒+砂质量（g）","灌砂后筒+砂质量（g）",
                    "试样质量（g）","盒号","盒质量（g）","盒+湿试样质量（g）","盒+干试样质量（g）","序号");
            for (int i=0;i< checklist.size();i++){
                XSSFRow row = sheet.createRow(i);// 创建第一行
                XSSFCell cell = row.createCell(0);// 创建第一行第一列
                cell.setCellValue(new XSSFRichTextString(checklist.get(i)));
                cell.setCellStyle(columnHeadStyle);
            }

            String filepath1=workpath+"/01路基土石方压实度_灰土实测数据.xlsx";

            OutputStream outputStream = new FileOutputStream(filepath1);
            workbook.write(outputStream);
            outputStream.flush();
            outputStream.close();

            XSSFWorkbook workbook1 = new XSSFWorkbook();// 创建一个Excel文件
            XSSFCellStyle columnHeadStyle1 = JjgFbgcCommonUtils.tsfCellStyle(workbook1);
            XSSFSheet sheet1 = workbook1.createSheet("实测数据");// 创建一个Excel的Sheet
            sheet1.setColumnWidth(0,28*256);
            List<String> checklist1 = Arrays.asList("路基压实度_砂砾规定值","试验时间","桩号","a",
                    "b","c","取样桩号及位置","试坑深度(cm)", "灌砂前筒+砂质量（g）","灌砂后筒+砂质量（g）",
                    "锥体砂重","量砂的密度ρs","混合料的湿质量","盒号","盒+干质量","盒质量","5-38mm颗粒质量(g)","序号");
            for (int i=0;i< checklist1.size();i++){
                XSSFRow row = sheet1.createRow(i);// 创建第一行
                XSSFCell cell = row.createCell(0);// 创建第一行第一列
                cell.setCellValue(new XSSFRichTextString(checklist1.get(i)));
                cell.setCellStyle(columnHeadStyle1);
            }
            String filepath2 = workpath + "/01路基土石方压实度_砂砾实测数据.xlsx";

            OutputStream outputStream1 = new FileOutputStream(filepath2);
            workbook1.write(outputStream1);
            outputStream1.flush();
            outputStream1.close();


            JjgFbgcUtils.createDirectory("路基工程",filepath);
            JjgFbgcUtils.createDirectory("涵洞",filepath+"/路基工程");
            JjgFbgcUtils.createDirectory("排水工程",filepath+"/路基工程");
            JjgFbgcUtils.createDirectory("路基土石方",filepath+"/路基工程");
            JjgFbgcUtils.createDirectory("小桥工程",filepath+"/路基工程");
            JjgFbgcUtils.createDirectory("支挡工程",filepath+"/路基工程");



            List<File> list=new ArrayList<>();
            list.add(new File(filepath+"/08涵洞砼强度实测数据.xlsx"));
            list.add(new File(filepath+"/09涵洞结构尺寸实测数据.xlsx"));
            JjgFbgcUtils.addFile(list,filepath+"/路基工程/涵洞",filepath);
            list.clear();
            list.add(new File(filepath+"/04排水断面尺寸实测数据.xlsx"));
            list.add(new File(filepath+"/05排水铺砌厚度实测数据.xlsx"));
            JjgFbgcUtils.addFile(list,filepath+"/路基工程/排水工程",filepath);
            list.clear();
            list.add(new File(filepath+"/01路基土石方压实度_灰土实测数据.xlsx"));
            list.add(new File(filepath+"/01路基土石方压实度_砂砾实测数据.xlsx"));
            list.add(new File(filepath+"/01路基压实度沉降实测数据.xlsx"));
            list.add(new File(filepath+"/02路基弯沉(贝克曼梁法)实测数据.xlsx"));
            list.add(new File(filepath+"/02路基弯沉(落锤法)实测数据.xlsx"));
            list.add(new File(filepath+"/03路基边坡实测数据.xlsx"));
            JjgFbgcUtils.addFile(list,filepath+"/路基工程/路基土石方",filepath);
            list.clear();
            list.add(new File(filepath+"/07小桥结构尺寸实测数据.xlsx"));
            list.add(new File(filepath+"/06小桥砼强度实测数据.xlsx"));
            JjgFbgcUtils.addFile(list,filepath+"/路基工程/小桥工程",filepath);
            list.clear();
            list.add(new File(filepath+"/11支挡断面尺寸实测数据.xlsx"));
            list.add(new File(filepath+"/10支挡砼强度实测数据.xlsx"));
            JjgFbgcUtils.addFile(list,filepath+"/路基工程/支挡工程",filepath);
            File file=new File( filepath);
            File[] files=file.listFiles();
            for(File f:files){
                if(!f.isDirectory()){
                    f.delete();
                }
            }





        } catch (IOException e) {
            throw new RuntimeException(e);
        }


    }
    @Override
    public void importljgc( CommonInfoVo commonInfoVo,String workpath){

        String zippath=workpath;

        try {

            File unzipfile = new File(zippath);

            File[] filest = unzipfile.listFiles();
            File[] files=null;
            if(zippath.substring(zippath.length()-2).equals("暂存")){
                File f_t=filest[0];
                files = f_t.listFiles();
            }
            else{
                files = filest;
            }
            for(File f:files){
                File[] files1=f.listFiles();

                for(File f1:files1){
                    String name=f1.getName();

                    switch (name){
                        case "08涵洞砼强度实测数据.xlsx":
                            commonInfoVo.setFbgc("涵洞");
                            jjgFbgcLjgcHdgqdService.importhdgqd(JjgFbgcUtils.getMultipartFile(f1),commonInfoVo);
                            break;
                        case "09涵洞结构尺寸实测数据.xlsx":
                            commonInfoVo.setFbgc("涵洞");
                            jjgFbgcLjgcHdjgccService.importhdjgcc(JjgFbgcUtils.getMultipartFile(f1),commonInfoVo);
                            break;
                        case "04排水断面尺寸实测数据.xlsx":
                            commonInfoVo.setFbgc("排水工程");
                            jjgFbgcLjgcPsdmccService.importpsdmcc(JjgFbgcUtils.getMultipartFile(f1),commonInfoVo);
                            break;
                        case "05排水铺砌厚度实测数据.xlsx":
                            commonInfoVo.setFbgc("排水工程");
                            jjgFbgcLjgcPspqhdService.importpspqhd(JjgFbgcUtils.getMultipartFile(f1),commonInfoVo);
                            break;
                        case "01路基土石方压实度_灰土实测数据.xlsx":
                            commonInfoVo.setFbgc("路基土石方");
                            jjgFbgcLjgcLjtsfysdHtService.importysdht(JjgFbgcUtils.getMultipartFile(f1),commonInfoVo);
                            break;
                        case "01路基土石方压实度_砂砾实测数据.xlsx":
                            commonInfoVo.setFbgc("路基土石方");
                            jjgFbgcLjgcLjtsfysdSlService.importysdsl(JjgFbgcUtils.getMultipartFile(f1),commonInfoVo);
                            break;
                        case "01路基压实度沉降实测数据.xlsx":
                            commonInfoVo.setFbgc("路基土石方(石方路基)");
                            jjgFbgcLjgcLjcjService.importljcj(JjgFbgcUtils.getMultipartFile(f1),commonInfoVo);
                            break;
                        case "02路基弯沉(贝克曼梁法)实测数据.xlsx":
                            commonInfoVo.setFbgc("路基土石方");
                            jjgFbgcLjgcLjwcService.importljwc(JjgFbgcUtils.getMultipartFile(f1),commonInfoVo);
                            break;
                        case "02路基弯沉(落锤法)实测数据.xlsx":
                            commonInfoVo.setFbgc("路基土石方");
                            jjgFbgcLjgcLjwcLcfService.importljwclcf(JjgFbgcUtils.getMultipartFile(f1),commonInfoVo);
                            break;
                        case "03路基边坡实测数据.xlsx":
                            commonInfoVo.setFbgc("路基土石方");
                            jjgFbgcLjgcLjbpService.importljbp(JjgFbgcUtils.getMultipartFile(f1),commonInfoVo);
                            break;
                        case "07小桥结构尺寸实测数据.xlsx":
                            commonInfoVo.setFbgc("小桥工程");
                            jjgFbgcLjgcXqjgccService.importxqjgcc(JjgFbgcUtils.getMultipartFile(f1),commonInfoVo);
                            break;
                        case "06小桥砼强度实测数据.xlsx":
                            commonInfoVo.setFbgc("小桥工程");

                            MultipartFile m=JjgFbgcUtils.getMultipartFile(f1);
                            System.out.println("文件"+m.getOriginalFilename());

                            jjgFbgcLjgcXqgqdService.importxqgqd(m,commonInfoVo);
                            break;
                        case "11支挡断面尺寸实测数据.xlsx":
                            commonInfoVo.setFbgc("支挡工程");
                            jjgFbgcLjgcZddmccService.importzddmcc(JjgFbgcUtils.getMultipartFile(f1),commonInfoVo);
                            break;
                        case "10支挡砼强度实测数据.xlsx":
                            commonInfoVo.setFbgc("支挡工程");
                            jjgFbgcLjgcZdgqdService.importzdgqd(JjgFbgcUtils.getMultipartFile(f1),commonInfoVo);
                            break;
                    }
                }
            }

            JjgFbgcUtils.deleteDirAndFiles(unzipfile);




        }  catch (IOException e) {
            throw new RuntimeException(e);
        } catch (ParseException e) {
            throw new RuntimeException(e);
        }


    }
    @Override
    public  void generateJdb(CommonInfoVo commonInfoVo) throws Exception{
        commonInfoVo.setFbgc("涵洞");
        jjgFbgcLjgcHdgqdService.generateJdb(commonInfoVo);
        jjgFbgcLjgcHdjgccService.generateJdb(commonInfoVo);
        commonInfoVo.setFbgc("排水工程");
        jjgFbgcLjgcPspqhdService.generateJdb(commonInfoVo);
        jjgFbgcLjgcPsdmccService.generateJdb(commonInfoVo);
        commonInfoVo.setFbgc("路基土石方(石方路基)");
        jjgFbgcLjgcLjcjService.generateJdb(commonInfoVo);

        commonInfoVo.setFbgc("小桥");
        jjgFbgcLjgcXqjgccService.generateJdb(commonInfoVo);
        jjgFbgcLjgcXqgqdService.generateJdb(commonInfoVo);
        commonInfoVo.setFbgc("支挡工程");
        jjgFbgcLjgcZdgqdService.generateJdb(commonInfoVo);
        jjgFbgcLjgcZddmccService.generateJdb(commonInfoVo);
        commonInfoVo.setFbgc("路基土石方");

        jjgFbgcLjgcLjwcService.generateJdb(commonInfoVo);
        jjgFbgcLjgcLjwcLcfService.generateJdb(commonInfoVo);
        jjgFbgcLjgcLjbpService.generateJdb(commonInfoVo);
        jjgFbgcLjgcLjtsfysdHtService.generateJdb(commonInfoVo);






    }
    @Override
    public  void download(HttpServletResponse response,String filename,String workpath){


        JjgFbgcUtils.createDirectory("路基工程",workpath);
        JjgFbgcUtils.createDirectory("涵洞",workpath+"/路基工程");
        JjgFbgcUtils.createDirectory("排水工程",workpath+"/路基工程");
        JjgFbgcUtils.createDirectory("路基土石方",workpath+"/路基工程");
        JjgFbgcUtils.createDirectory("小桥工程",workpath+"/路基工程");
        JjgFbgcUtils.createDirectory("支挡工程",workpath+"/路基工程");



        List<File> list=new ArrayList<>();
        list.add(new File(workpath+"/08路基涵洞砼强度.xlsx"));
        list.add(new File(workpath+"/09路基涵洞结构尺寸.xlsx"));
        JjgFbgcUtils.addFile(list,workpath+"/路基工程/涵洞",workpath);
        list.clear();
        list.add(new File(workpath+"/04路基排水断面尺寸.xlsx"));
        list.add(new File(workpath+"/05路基排水铺砌厚度.xlsx"));
        JjgFbgcUtils.addFile(list,workpath+"/路基工程/排水工程",workpath);
        list.clear();
        list.add(new File(workpath+"/01路基土石方压实度.xlsx"));
        list.add(new File(workpath+"/01路基压实度沉降.xlsx"));
        list.add(new File(workpath+"/02路基弯沉(贝克曼梁法).xlsx"));
        list.add(new File(workpath+"/02路基弯沉(落锤法).xlsx"));
        list.add(new File(workpath+"/03路基边坡.xlsx"));
        JjgFbgcUtils.addFile(list,workpath+"/路基工程/路基土石方",workpath);
        list.clear();
        list.add(new File(workpath+"/07路基小桥结构尺寸.xlsx"));
        list.add(new File(workpath+"/06路基小桥砼强度.xlsx"));
        JjgFbgcUtils.addFile(list,workpath+"/路基工程/小桥工程",workpath);
        list.clear();
        list.add(new File(workpath+"/11路基支挡断面尺寸.xlsx"));
        list.add(new File(workpath+"/10路基支挡砼强度.xlsx"));
        JjgFbgcUtils.addFile(list,workpath+"/路基工程/支挡工程",workpath);

    }
}
