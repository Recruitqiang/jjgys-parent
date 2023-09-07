package glgc.jjgys.system.service.impl;

import com.alibaba.excel.metadata.BaseRowModel;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.projectvo.jagc.JjgFbgcJtaqssJabxVo;
import glgc.jjgys.model.projectvo.jagc.JjgFbgcJtaqssJabzVo;
import glgc.jjgys.model.projectvo.jagc.JjgFbgcJtaqssJathldmccVo;
import glgc.jjgys.model.projectvo.jagc.JjgFbgcJtaqssJathlqdVo;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.lmgc.*;
import glgc.jjgys.system.service.*;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.JjgFbgcUtils;
import net.lingala.zip4j.core.ZipFile;
import net.lingala.zip4j.exception.ZipException;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.security.core.parameters.P;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.text.ParseException;
import java.util.*;

import static java.util.Arrays.asList;

@Service
public class JjgFbgcJagcServiceImpl implements JjgFbgcJagcService {
    @Resource
    private JjgFbgcJtaqssJabxService jjgFbgcJtaqssJabxService;
    @Resource
    private JjgFbgcJtaqssJabzService jjgFbgcJtaqssJabzService;
    @Resource
    private JjgFbgcJtaqssJabxfhlService jjgFbgcJtaqssJabxfhlService;
    @Resource
    private JjgFbgcJtaqssJathlqdService jjgFbgcJtaqssJathlqdService;
    @Resource
    private  JjgFbgcJtaqssJathldmccService jjgFbgcJtaqssJathldmccService;
    @Override
    public void exportjagc(HttpServletResponse response,String workpath){
        String filepath = workpath;

        try {
            List<String> filenames=asList("02交安标线实测数据","01交安标志实测数据","04交安砼护栏强度实测数据","05交安砼护栏断面尺寸实测数据"
            );

            List<BaseRowModel> models=asList(new JjgFbgcJtaqssJabxVo(),new JjgFbgcJtaqssJabzVo(),new JjgFbgcJtaqssJathlqdVo(),new JjgFbgcJtaqssJathldmccVo());

            for(int i=0;i<filenames.size();i++){
                File file=new File(filepath+"/"+filenames.get(i)+".xlsx");

                if(!file.exists()){
                    ExcelUtil.saveLocal(filepath+"/"+filenames.get(i),models.get(i),"实测数据");
                    System.out.println("生成表格"+file.getName());
                }



            }
            XSSFWorkbook workbook = new XSSFWorkbook();// 创建一个Excel文件
            XSSFCellStyle columnHeadStyle = JjgFbgcCommonUtils.tsfCellStyle(workbook);

            XSSFSheet sheet = workbook.createSheet("实测数据");// 创建一个Excel的Sheet
            sheet.setColumnWidth(0,28*256);
            List<String> checklist = Arrays.asList("检测日期","位置及类型","基底厚度规定值(mm)","基底厚度实测值1(mm)",
                    "基底厚度实测值2(mm)","基底厚度实测值3(mm)","立柱壁厚规定值(mm)","立柱壁厚实测值1(mm)","立柱壁厚实测值2(mm)",
                    "立柱壁厚实测值3(mm)","中心高度规定值(mm)","中心高度允许偏差+(mm)",
                    "中心高度允许偏差-(mm)","中心高度实测值1(mm)","中心高度实测值2(mm)","中心高度实测值3(mm)","埋入深度规定值(mm)","埋入深度实测值(mm)");
            for (int i=0;i< checklist.size();i++){
                XSSFRow row = sheet.createRow(i);// 创建第一行
                XSSFCell cell = row.createCell(0);// 创建第一行第一列
                cell.setCellValue(new XSSFRichTextString(checklist.get(i)));
                cell.setCellStyle(columnHeadStyle);
            }
            String filepath1=workpath+"/03交安波形防护栏实测数据.xlsx";

            OutputStream outputStream = new FileOutputStream(filepath1);
            workbook.write(outputStream);
            outputStream.flush();
            outputStream.close();


            JjgFbgcUtils.createDirectory("交安工程",filepath);
            JjgFbgcUtils.createDirectory("标线",filepath+"/交安工程");
            JjgFbgcUtils.createDirectory("标志",filepath+"/交安工程");
            JjgFbgcUtils.createDirectory("防护栏",filepath+"/交安工程");





            List<File> list=new ArrayList<>();
            list.add(new File(filepath+"/02交安标线实测数据.xlsx"));


            JjgFbgcUtils.addFile(list,filepath+"/交安工程/标线",filepath);
            list.clear();
            list.add(new File(filepath+"/01交安标志实测数据.xlsx"));


            JjgFbgcUtils.addFile(list,filepath+"/交安工程/标志",filepath);
            list.clear();
            list.add(new File(filepath+"/03交安波形防护栏实测数据.xlsx"));
            list.add(new File(filepath+"/04交安砼护栏强度实测数据.xlsx"));
            list.add(new File(filepath+"/05交安砼护栏断面尺寸实测数据.xlsx"));


            JjgFbgcUtils.addFile(list,filepath+"/交安工程/防护栏",filepath);
            list.clear();
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
    public  void  importjagc( CommonInfoVo commonInfoVo,String workpath){
        String zippath=workpath;

        try {

            File unzipfile = new File(zippath );

            File[] filest = unzipfile.listFiles();
            File[] files=null;
            if(zippath.substring(zippath.length()-2).equals("暂存")){
                 File f_t=filest[0];
                 files = f_t.listFiles();
            }
            else{
                 files = filest;
            }

            for (File f : files) {
                File[] files1 = f.listFiles();

                for (File f1 : files1) {
                    String name = f1.getName();

                    switch (name) {
                        case "02交安标线实测数据.xlsx":
                            commonInfoVo.setFbgc("标线");
                            jjgFbgcJtaqssJabxService.importjabx(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "01交安标志实测数据.xlsx":
                            commonInfoVo.setFbgc("交通安全设施");
                            jjgFbgcJtaqssJabzService.importjabz(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "03交安波形防护栏实测数据.xlsx":
                            commonInfoVo.setFbgc("防护栏");

                            jjgFbgcJtaqssJabxfhlService.importjabxfhl(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);

                            break;
                        case "04交安砼护栏强度实测数据.xlsx":
                            commonInfoVo.setFbgc("砼防护栏");
                            jjgFbgcJtaqssJathlqdService.importjathlqd(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "05交安砼护栏断面尺寸实测数据.xlsx":
                            commonInfoVo.setFbgc("砼防护栏");
                            jjgFbgcJtaqssJathldmccService.importjathldmcc(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
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
    public  void generateJdb(CommonInfoVo commonInfoVo) throws IOException{
        try {
            commonInfoVo.setFbgc("标线");
            jjgFbgcJtaqssJabxService.generateJdb(commonInfoVo);
            commonInfoVo.setFbgc("交通安全设施");
            jjgFbgcJtaqssJabzService.generateJdb(commonInfoVo);
            commonInfoVo.setFbgc("防护栏");
            jjgFbgcJtaqssJabxfhlService.generateJdb(commonInfoVo);
            commonInfoVo.setFbgc("砼防护栏");
            jjgFbgcJtaqssJathlqdService.generateJdb(commonInfoVo);
            jjgFbgcJtaqssJathldmccService.generateJdb(commonInfoVo);

        } catch (ParseException e) {
            throw new RuntimeException(e);
        }catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
    @Override
    public  void download(HttpServletResponse response,String filename,String proname,String htd,String workpath){


        JjgFbgcUtils.createDirectory("交安工程",workpath);
        JjgFbgcUtils.createDirectory("标线",workpath+"/交安工程");
        JjgFbgcUtils.createDirectory("标志",workpath+"/交安工程");
        JjgFbgcUtils.createDirectory("防护栏",workpath+"/交安工程");





        List<File> list=new ArrayList<>();

        list.add(new File(workpath+"/57交安标线厚度.xlsx"));
        File file=new File(workpath+"/57交安标线白线逆反射系数.xlsx");
        if(file.exists()){
            list.add(file);
        }
        file=new File(workpath+"/57交安标线黄线逆反射系数.xlsx");
        if(file.exists()){
            list.add(file);
        }
        JjgFbgcUtils.addFile(list,workpath+"/交安工程/标线",workpath);
        list.clear();
        list.add(new File(workpath+"/56交安标志.xlsx"));
        JjgFbgcUtils.addFile(list,workpath+"/交安工程/标志",workpath);
        list.clear();
        list.add(new File(workpath+"/58交安钢防护栏.xlsx"));
        list.add(new File(workpath+"/59交安砼护栏强度.xlsx"));
        list.add(new File(workpath+"/60交安砼护栏断面尺寸.xlsx"));
        JjgFbgcUtils.addFile(list,workpath+"/交安工程/防护栏",workpath);
        list.clear();


    }

}
