package glgc.jjgys.system.service.impl;

import com.alibaba.excel.metadata.BaseRowModel;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.common.excel.ExcelWriterFactroy;
import glgc.jjgys.model.project.JjgFbgcQlgcSbTqd;
import glgc.jjgys.model.projectvo.ljgc.*;
import glgc.jjgys.model.projectvo.qlgc.*;
import glgc.jjgys.system.mapper.JjgFbgcQlgcQmgzsdMapper;
import glgc.jjgys.system.mapper.JjgFbgcQlgcQmhpMapper;
import glgc.jjgys.system.mapper.JjgFbgcQlgcQmpzdMapper;
import glgc.jjgys.system.service.*;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.JjgFbgcUtils;
import net.lingala.zip4j.core.ZipFile;
import net.lingala.zip4j.exception.ZipException;
import net.lingala.zip4j.model.FileHeader;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.text.ParseException;
import java.util.*;
import java.util.zip.ZipEntry;


import static java.util.Arrays.asList;

@Service
public class JjgFbgcQlgcServiceImpl implements JjgFbgcQlgcService{
    @Resource
    private JjgFbgcQlgcQmpzdMapper jjgFbgcQlgcQmpzdMapper;
    @Resource
    private JjgFbgcQlgcQmhpMapper jjgFbgcQlgcQmhpMapper;
    @Autowired
    private JjgFbgcQlgcQmgzsdMapper jjgFbgcQlgcQmgzsdMapper;
    @Resource
    private JjgFbgcQlgcQmpzdService jjgFbgcQlgcQmpzdService;
    @Resource
    private JjgFbgcQlgcQmhpService jjgFbgcQlgcQmhpService;
    @Resource
    private JjgFbgcQlgcQmgzsdService jjgFbgcQlgcQmgzsdService;
    @Resource
    private JjgFbgcQlgcSbTqdService jjgFbgcQlgcSbTqdService;
    @Resource
    private JjgFbgcQlgcSbJgccService jjgFbgcQlgcSbJgccService;
    @Resource
    private  JjgFbgcQlgcSbBhchdService jjgFbgcQlgcSbBhchdService;
    @Resource
    private  JjgFbgcQlgcXbTqdService jjgFbgcQlgcXbTqdService;
    @Resource
    private  JjgFbgcQlgcXbJgccService jjgFbgcQlgcXbJgccService;
    @Resource
    private  JjgFbgcQlgcXbBhchdService jjgFbgcQlgcXbBhchdService;
    @Resource
    private  JjgFbgcQlgcXbSzdService jjgFbgcQlgcXbSzdService;

    @Override
    public void exportqlgc(HttpServletResponse response,String workpath){
        String filepath = workpath;


            List<String> filenames=asList("09桥面平整度三米直尺法实测数据","桥面横坡实测数据","11桥面构造深度手工铺砂法实测数据","05桥梁上部砼强度实测数据","06桥梁上部结构尺寸实测数据",
                    "07桥梁上部保护层厚度实测数据","01桥梁下部墩台砼强度实测数据","02桥梁下部结构尺寸实测数据","03桥梁下部保护层厚度实测数据","04桥梁下部竖直度实测数据"
                    );
            List<String> sheetnames=asList( "实测数据","实测数据","实测数据","实测数据","实测数据","实测数据","实测数据","实测数据","实测数据","实测数据");
            List<BaseRowModel> models=asList(new JjgFbgcQlgcQmpzdVo(),new JjgFbgcQlgcQmhpVo(),new JjgFbgcQlgcQmgzsdVo(),new JjgFbgcQlgcSbTqdVo(),new JjgFbgcQlgcSbJgccVo(),
                    new JjgFbgcQlgcSbBhchdVo(),new JjgFbgcQlgcXbTqdVo(),new JjgFbgcQlgcXbJgccVo(),new JjgFbgcQlgcXbBhchdVo(),new JjgFbgcQlgcXbSzdVo());

            for(int i=0;i<filenames.size();i++){
                File file=new File(filepath+"/"+filenames.get(i)+".xlsx");



                if(!file.exists()){
                    ExcelUtil.saveLocal(filepath+"/"+filenames.get(i),models.get(i),sheetnames.get(i));
                    System.out.println("生成表格"+file.getName());
                }




            }



            JjgFbgcUtils.createDirectory("桥梁工程",filepath);
            JjgFbgcUtils.createDirectory("桥面",filepath+"/桥梁工程");
            JjgFbgcUtils.createDirectory("上部",filepath+"/桥梁工程");
            JjgFbgcUtils.createDirectory("下部",filepath+"/桥梁工程");




            List<File> list=new ArrayList<>();
            list.add(new File(filepath+"/09桥面平整度三米直尺法实测数据.xlsx"));
            list.add(new File(filepath+"/桥面横坡实测数据.xlsx"));
            list.add(new File(filepath+"/11桥面构造深度手工铺砂法实测数据.xlsx"));
            JjgFbgcUtils.addFile(list,filepath+"/桥梁工程/桥面",filepath);
            list.clear();
            list.add(new File(filepath+"/05桥梁上部砼强度实测数据.xlsx"));
            list.add(new File(filepath+"/06桥梁上部结构尺寸实测数据.xlsx"));
            list.add(new File(filepath+"/07桥梁上部保护层厚度实测数据.xlsx"));
            JjgFbgcUtils.addFile(list,filepath+"/桥梁工程/上部",filepath);
            list.clear();
            list.add(new File(filepath+"/01桥梁下部墩台砼强度实测数据.xlsx"));
            list.add(new File(filepath+"/02桥梁下部结构尺寸实测数据.xlsx"));
            list.add(new File(filepath+"/03桥梁下部保护层厚度实测数据.xlsx"));
            list.add(new File(filepath+"/04桥梁下部竖直度实测数据.xlsx"));
            JjgFbgcUtils.addFile(list,filepath+"/桥梁工程/下部",filepath);
            File file=new File( filepath);
            File[] files=file.listFiles();
            for(File f:files){
                if(!f.isDirectory()){
                    f.delete();
                }
            }












    }
    @Override
    public  void  importqlgc( CommonInfoVo commonInfoVo,String workpath){
        String zippath=workpath;



            File unzipfile = new File(zippath);

            File[] filest = unzipfile.listFiles();
            File[] files=null;
            if(zippath.substring(zippath.length()-2).equals("暂存")){
                File f_t=filest[0];
                System.out.println("sssss"+f_t.getName());
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
                        case "09桥面平整度三米直尺法实测数据.xlsx":
                            commonInfoVo.setFbgc("桥梁工程");
                            jjgFbgcQlgcQmpzdService.importQmpzd(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "桥面横坡实测数据.xlsx":
                            commonInfoVo.setFbgc("桥面系");
                            jjgFbgcQlgcQmhpService.importqmhp(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "11桥面构造深度手工铺砂法实测数据.xlsx":
                            commonInfoVo.setFbgc("桥梁工程");
                            jjgFbgcQlgcQmgzsdService.importqmgzsd(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "05桥梁上部砼强度实测数据.xlsx":
                            commonInfoVo.setFbgc("桥梁上部");
                            jjgFbgcQlgcSbTqdService.importqlsbtqd(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "06桥梁上部结构尺寸实测数据.xlsx":
                            commonInfoVo.setFbgc("桥梁上部");
                            jjgFbgcQlgcSbJgccService.importqlsbjgcc(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "07桥梁上部保护层厚度实测数据.xlsx":
                            commonInfoVo.setFbgc("桥梁上部");
                            jjgFbgcQlgcSbBhchdService.importbhchd(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "01桥梁下部墩台砼强度实测数据.xlsx":
                            commonInfoVo.setFbgc("桥梁下部");
                            jjgFbgcQlgcXbTqdService.importqlxbtqd(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "02桥梁下部结构尺寸实测数据.xlsx":
                            commonInfoVo.setFbgc("桥梁下部");
                            jjgFbgcQlgcXbJgccService.importqlxbjgcc(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "03桥梁下部保护层厚度实测数据.xlsx":
                            commonInfoVo.setFbgc("桥梁下部");
                            jjgFbgcQlgcXbBhchdService.importbhchd(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "04桥梁下部竖直度实测数据.xlsx":
                            commonInfoVo.setFbgc("桥梁下部");
                            jjgFbgcQlgcXbSzdService.importqlxbszd(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;

                    }
                }
            }

            JjgFbgcUtils.deleteDirAndFiles(unzipfile);





    }
    @Override
    public  void generateJdb(CommonInfoVo commonInfoVo) throws IOException {
        commonInfoVo.setFbgc("桥梁工程");
        jjgFbgcQlgcQmpzdService.generateJdb(commonInfoVo);
        jjgFbgcQlgcQmgzsdService.generateJdb(commonInfoVo);

        try {
            commonInfoVo.setFbgc("桥面系");
            jjgFbgcQlgcQmhpService.generateJdb(commonInfoVo);
            commonInfoVo.setFbgc("桥梁上部");
            jjgFbgcQlgcSbTqdService.generateJdb(commonInfoVo);
            jjgFbgcQlgcSbJgccService.generateJdb(commonInfoVo);
            jjgFbgcQlgcSbBhchdService.generateJdb(commonInfoVo);
            commonInfoVo.setFbgc("桥梁下部");
            jjgFbgcQlgcXbTqdService.generateJdb(commonInfoVo);
            jjgFbgcQlgcXbJgccService.generateJdb(commonInfoVo);
            jjgFbgcQlgcXbBhchdService.generateJdb(commonInfoVo);
            jjgFbgcQlgcXbSzdService.generateJdb(commonInfoVo);
        } catch (ParseException e) {
            throw new RuntimeException(e);
        }



    }

    @Override
    public  void download(HttpServletResponse response,String filename,String proname,String htd,String workpath){


        JjgFbgcUtils.createDirectory("桥梁工程",workpath);
        JjgFbgcUtils.createDirectory("桥面",workpath+"/桥梁工程");
        JjgFbgcUtils.createDirectory("上部",workpath+"/桥梁工程");
        JjgFbgcUtils.createDirectory("下部",workpath+"/桥梁工程");




        List<File> list=new ArrayList<>();
        List<Map<String,Object>> qlmclist = jjgFbgcQlgcQmpzdMapper.selectqlmc(proname,htd,"桥梁工程");
        if (qlmclist.size()>0){
            for (Map<String, Object> m : qlmclist)
            {
                for (String k : m.keySet()){
                    String qlmc = m.get(k).toString();
                    list.add(new File(workpath+"/34桥面平整度3米直尺法-"+qlmc+".xlsx"));
                }
            }
        }
        qlmclist.clear();
        qlmclist=jjgFbgcQlgcQmhpMapper.selectqlmc(proname,htd,"桥面系");
        if (qlmclist.size()>0){
            for (Map<String, Object> m : qlmclist)
            {
                for (String k : m.keySet()){
                    String qlmc = m.get(k).toString();
                    list.add(new File(workpath+"/35桥面横坡-"+qlmc+".xlsx"));
                }
            }
        }
        qlmclist.clear();
        qlmclist=jjgFbgcQlgcQmgzsdMapper.selectqlmc(proname,htd,"桥梁工程");
        if (qlmclist.size()>0){
            for (Map<String, Object> m : qlmclist)
            {
                for (String k : m.keySet()){
                    String qlmc = m.get(k).toString();
                    list.add(new File(workpath+"/37构造深度手工铺沙法-"+qlmc+".xlsx"));
                }
            }
        }

        JjgFbgcUtils.addFile(list,workpath+"/桥梁工程/桥面",workpath);
        list.clear();

        list.add(new File(workpath+"/29桥梁上部砼强度.xlsx"));
        list.add(new File(workpath+"/30桥梁上部主要结构尺寸.xlsx"));
        list.add(new File(workpath+"/31桥梁上部保护层厚度.xlsx"));
        JjgFbgcUtils.addFile(list,workpath+"/桥梁工程/上部",workpath);
        list.clear();
        list.add(new File(workpath+"/25桥梁下部墩台砼强度.xlsx"));
        list.add(new File(workpath+"/26桥梁下部主要结构尺寸.xlsx"));
        list.add(new File(workpath+"/27桥梁下部保护层厚度.xlsx"));
        list.add(new File(workpath+"/28桥梁下部墩台垂直度.xlsx"));

        JjgFbgcUtils.addFile(list,workpath+"/桥梁工程/下部",workpath);




    }


}
