package glgc.jjgys.system.service.impl;

import com.alibaba.excel.metadata.BaseRowModel;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.common.excel.ExcelWriterFactroy;
import glgc.jjgys.model.project.JjgFbgcSdgcHntlmqd;
import glgc.jjgys.model.project.JjgFbgcSdgcSdlqlmysd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.qlgc.*;
import glgc.jjgys.model.projectvo.sdgc.*;
import glgc.jjgys.system.mapper.*;
import glgc.jjgys.system.service.*;
import glgc.jjgys.system.utils.JjgFbgcUtils;
import net.lingala.zip4j.core.ZipFile;
import net.lingala.zip4j.exception.ZipException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static java.util.Arrays.asList;

@Service
public class JjgFbgcSdgcServiceImpl implements JjgFbgcSdgcService {
    @Autowired
    private JjgFbgcSdgcCqhdMapper jjgFbgcSdgcCqhdMapper;
    @Autowired
    private JjgFbgcSdgcSdlqlmysdMapper jjgFbgcSdgcSdlqlmysdMapper;
    @Autowired
    private JjgFbgcSdgcLmssxsMapper jjgFbgcSdgcLmssxsMapper;
    @Autowired
    private JjgFbgcSdgcHntlmqdMapper jjgFbgcSdgcHntlmqdMapper;
    @Autowired
    private JjgFbgcSdgcTlmxlbgcMapper jjgFbgcSdgcTlmxlbgcMapper;
    @Autowired
    private JjgFbgcSdgcLmgzsdsgpsfMapper jjgFbgcSdgcLmgzsdsgpsfMapper;
    @Autowired
    private JjgFbgcSdgcGssdlqlmhdzxfMapper jjgFbgcSdgcGssdlqlmhdzxfMapper;
    @Autowired
    private JjgFbgcSdgcSdhntlmhdzxfMapper jjgFbgcSdgcSdhntlmhdzxfMapper;
    @Autowired
    private JjgFbgcSdgcSdhpMapper jjgFbgcSdgcSdhpMapper;
    @Resource
    private JjgFbgcSdgcCqhdService jjgFbgcSdgcCqhdService;
    @Resource
    private JjgFbgcSdgcCqtqdService jjgFbgcSdgcCqtqdService;
    @Resource
    private JjgFbgcSdgcDmpzdService jjgFbgcSdgcDmpzdService;
    @Resource
    private JjgFbgcSdgcSdlqlmysdService jjgFbgcSdgcSdlqlmysdService;
    @Resource
    private  JjgFbgcSdgcLmssxsService jjgFbgcSdgcLmssxsService;
    @Resource
    private  JjgFbgcSdgcHntlmqdService jjgFbgcSdgcHntlmqdService;
    @Resource
    private  JjgFbgcSdgcTlmxlbgcService jjgFbgcSdgcTlmxlbgcService;
    @Resource
    private  JjgFbgcSdgcLmgzsdsgpsfService jjgFbgcSdgcLmgzsdsgpsfService;
    @Resource
    private  JjgFbgcSdgcGssdlqlmhdzxfService jjgFbgcSdgcGssdlqlmhdzxfService;
    @Resource
    private  JjgFbgcSdgcSdhntlmhdzxfService jjgFbgcSdgcSdhntlmhdzxfService;
    @Resource
    private  JjgFbgcSdgcSdhpService jjgFbgcSdgcSdhpService;
    @Resource
    private  JjgFbgcSdgcZtkdService jjgFbgcSdgcZtkdService;


    @Override
    public void exportsdgc(HttpServletResponse response,String workpath) {
        String filepath = workpath;


            List<String> filenames=asList("01隧道衬砌砼强度实测数据","02隧道衬砌厚度度实测数据","03隧道大面平整度实测数据","06隧道沥青路面压实度实测数据","06隧道沥青路面渗水系数实测数据",
                    "10隧道混凝土路面强度实测数据","11隧道砼路面相邻板高差实测数据","13隧道路面构造深度手工铺砂法实测数据","14高速隧道沥青路面厚度钻芯法实测数据","14隧道混凝土路面厚度钻芯法实测数据",
                    "15隧道横坡实测数据","04隧道总体宽度实测数据"
            );

            List<BaseRowModel> models=asList(new JjgFbgcSdgcCqtqdVo(),new JjgFbgcSdgcCqhdVo(),new JjgFbgcSdgcDmpzdVo(),new JjgFbgcSdgcSdlqlmysdVo(),new JjgFbgcSdgcLmssxsVo(),
                    new JjgFbgcSdgcHntlmqdVo(),new JjgFbgcSdgcTlmxlbgcVo(),new JjgFbgcSdgcLmgzsdsgpsfVo(),new JjgFbgcSdgcGssdlqlmhdzxfVo(),new JjgFbgcSdgcSdhntlmhdzxfVo(),
                    new JjgFbgcSdgcSdhpVo(),new JjgFbgcSdgcZtkdVo());

            for(int i=0;i<filenames.size();i++){
                File file=new File(filepath+"/"+filenames.get(i)+".xlsx");

                if(!file.exists()){
                    ExcelUtil.saveLocal(filepath+"/"+filenames.get(i),models.get(i),"实测数据");
                    System.out.println("生成表格"+file.getName());
                }



            }



            JjgFbgcUtils.createDirectory("隧道工程",filepath);
            JjgFbgcUtils.createDirectory("衬砌",filepath+"/隧道工程");
            JjgFbgcUtils.createDirectory("隧道路面",filepath+"/隧道工程");
            JjgFbgcUtils.createDirectory("总体",filepath+"/隧道工程");




            List<File> list=new ArrayList<>();
            list.add(new File(filepath+"/01隧道衬砌砼强度实测数据.xlsx"));
            list.add(new File(filepath+"/02隧道衬砌厚度度实测数据.xlsx"));
            list.add(new File(filepath+"/03隧道大面平整度实测数据.xlsx"));
            JjgFbgcUtils.addFile(list,filepath+"/隧道工程/衬砌",filepath);
            list.clear();
            list.add(new File(filepath+"/06隧道沥青路面压实度实测数据.xlsx"));
            list.add(new File(filepath+"/06隧道沥青路面渗水系数实测数据.xlsx"));
            list.add(new File(filepath+"/10隧道混凝土路面强度实测数据.xlsx"));
            list.add(new File(filepath+"/11隧道砼路面相邻板高差实测数据.xlsx"));
            list.add(new File(filepath+"/13隧道路面构造深度手工铺砂法实测数据.xlsx"));
            list.add(new File(filepath+"/14高速隧道沥青路面厚度钻芯法实测数据.xlsx"));
            list.add(new File(filepath+"/14隧道混凝土路面厚度钻芯法实测数据.xlsx"));
            list.add(new File(filepath+"/15隧道横坡实测数据.xlsx"));
            JjgFbgcUtils.addFile(list,filepath+"/隧道工程/隧道路面",filepath);
            list.clear();
            list.add(new File(filepath+"/04隧道总体宽度实测数据.xlsx"));

            JjgFbgcUtils.addFile(list,filepath+"/隧道工程/总体",filepath);
            File file=new File( filepath);
            File[] files=file.listFiles();
            for(File f:files){
                if(!f.isDirectory()){
                    f.delete();
                }
            }








    }

    @Override
    public void importsdgc( CommonInfoVo commonInfoVo,String workpath) {
        String zippath=workpath;



            File unzipfile = new File(zippath);

            File[] filest = unzipfile.listFiles();
            File[] files=null;

            if(zippath.substring(zippath.length()-2).equals("暂存")){
                File f_t=filest[0];
                System.out.println("ssss"+f_t.getName());
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
                        case "01隧道衬砌砼强度实测数据.xlsx":
                            commonInfoVo.setFbgc("衬砌");
                            jjgFbgcSdgcCqtqdService.importsdcqtsd(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "02隧道衬砌厚度度实测数据.xlsx":
                            commonInfoVo.setFbgc("衬砌");
                            jjgFbgcSdgcCqhdService.importsdcqhd(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "03隧道大面平整度实测数据.xlsx":
                            commonInfoVo.setFbgc("隧道工程");
                            jjgFbgcSdgcDmpzdService.importsddmpzd(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "06隧道沥青路面压实度实测数据.xlsx":
                            commonInfoVo.setFbgc("隧道工程");
                            jjgFbgcSdgcSdlqlmysdService.importsdlqlmysd(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "06隧道沥青路面渗水系数实测数据.xlsx":
                            commonInfoVo.setFbgc("隧道工程");
                            jjgFbgcSdgcLmssxsService.importSdlmssxs(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "10隧道混凝土路面强度实测数据.xlsx":
                            commonInfoVo.setFbgc("隧道工程");
                            jjgFbgcSdgcHntlmqdService.importsdhntlmqd(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "11隧道砼路面相邻板高差实测数据.xlsx":
                            commonInfoVo.setFbgc("隧道工程");
                            jjgFbgcSdgcTlmxlbgcService.importsdxlbgs(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "13隧道路面构造深度手工铺砂法实测数据.xlsx":
                            commonInfoVo.setFbgc("隧道工程");
                            jjgFbgcSdgcLmgzsdsgpsfService.importlmgzsdsgpsf(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "14高速隧道沥青路面厚度钻芯法实测数据.xlsx":
                            commonInfoVo.setFbgc("隧道工程");
                            jjgFbgcSdgcGssdlqlmhdzxfService.importsdzxf(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "14隧道混凝土路面厚度钻芯法实测数据.xlsx":
                            commonInfoVo.setFbgc("隧道工程");
                            jjgFbgcSdgcSdhntlmhdzxfService.importsdhntzxf(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "15隧道横坡实测数据.xlsx":
                            commonInfoVo.setFbgc("隧道工程");
                            jjgFbgcSdgcSdhpService.importsdhp(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "04隧道总体宽度实测数据.xlsx":
                            commonInfoVo.setFbgc("衬砌");
                            jjgFbgcSdgcZtkdService.importsdztkd(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                    }
                }
            }

            JjgFbgcUtils.deleteDirAndFiles(unzipfile);



    }

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException {


        try {
            commonInfoVo.setFbgc("衬砌");
            jjgFbgcSdgcCqtqdService.generateJdb(commonInfoVo);
            jjgFbgcSdgcCqhdService.generateJdb(commonInfoVo);
            jjgFbgcSdgcZtkdService.generateJdb(commonInfoVo);
            commonInfoVo.setFbgc("隧道工程");
            jjgFbgcSdgcDmpzdService.generateJdb(commonInfoVo);
            jjgFbgcSdgcSdlqlmysdService.generateJdb(commonInfoVo);
            jjgFbgcSdgcLmssxsService.generateJdb(commonInfoVo);
            jjgFbgcSdgcHntlmqdService.generateJdb(commonInfoVo);

            jjgFbgcSdgcTlmxlbgcService.generateJdb(commonInfoVo);
            jjgFbgcSdgcLmgzsdsgpsfService.generateJdb(commonInfoVo);
            jjgFbgcSdgcGssdlqlmhdzxfService.generateJdb(commonInfoVo);
            jjgFbgcSdgcSdhntlmhdzxfService.generateJdb(commonInfoVo);
            jjgFbgcSdgcSdhpService.generateJdb(commonInfoVo);
        } catch (ParseException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public void download(HttpServletResponse response, String filename, String proname, String htd, String workpath) {


        JjgFbgcUtils.createDirectory("隧道工程",workpath);
        JjgFbgcUtils.createDirectory("衬砌",workpath+"/隧道工程");
        JjgFbgcUtils.createDirectory("隧道路面",workpath+"/隧道工程");
        JjgFbgcUtils.createDirectory("总体",workpath+"/隧道工程");




        List<File> list=new ArrayList<>();
        List<Map<String,Object>> sdmclist = jjgFbgcSdgcCqhdMapper.selectsdmc(proname,htd,"衬砌");
        if (sdmclist.size()>0){
            for (Map<String, Object> m : sdmclist)
            {
                for (String k : m.keySet()){
                    String sdmc = m.get(k).toString();
                    list.add(new File(workpath+"/39隧道衬砌厚度-"+sdmc+".xlsx"));
                }
            }
        }
        sdmclist.clear();




        list.add(new File(workpath+"/38隧道衬砌砼强度.xlsx"));
        list.add(new File(workpath+"/40隧道大面平整度.xlsx"));
        JjgFbgcUtils.addFile(list,workpath+"/隧道工程/衬砌",workpath);
        list.clear();
        sdmclist = jjgFbgcSdgcSdlqlmysdMapper.selectsdmc(proname,htd,"隧道工程");
        if (sdmclist.size()>0){
            for (Map<String, Object> m : sdmclist)
            {
                for (String k : m.keySet()){
                    String sdmc = m.get(k).toString();
                    list.add(new File(workpath+"/43隧道沥青路面压实度-"+sdmc+".xlsx"));
                }
            }
        }
        sdmclist.clear();
        sdmclist = jjgFbgcSdgcLmssxsMapper.selectsdmc(proname,htd,"隧道工程");
        if (sdmclist.size()>0){
            for (Map<String, Object> m : sdmclist)
            {
                for (String k : m.keySet()){
                    String sdmc = m.get(k).toString();
                    list.add(new File(workpath+"/46隧道沥青路面渗水系数-"+sdmc+".xlsx"));
                }
            }
        }
        sdmclist.clear();
        sdmclist = jjgFbgcSdgcHntlmqdMapper.selectsdmc(proname,htd);
        if (sdmclist.size()>0){
            for (Map<String, Object> m : sdmclist)
            {
                for (String k : m.keySet()){
                    String sdmc = m.get(k).toString();
                    list.add(new File(workpath+"/47隧道混凝土路面强度-"+sdmc+".xlsx"));
                }
            }
        }
        sdmclist.clear();
        sdmclist = jjgFbgcSdgcTlmxlbgcMapper.selectsdmc(proname,htd);
        if (sdmclist.size()>0){
            for (Map<String, Object> m : sdmclist)
            {
                for (String k : m.keySet()){
                    String sdmc = m.get(k).toString();
                    list.add(new File(workpath+"/48隧道混凝土路面相邻板高差-"+sdmc+".xlsx"));
                }
            }
        }
        sdmclist.clear();
        sdmclist = jjgFbgcSdgcLmgzsdsgpsfMapper.selectsdmc(proname,htd,"隧道工程");
        if (sdmclist.size()>0){
            for (Map<String, Object> m : sdmclist)
            {
                for (String k : m.keySet()){
                    String sdmc = m.get(k).toString();
                    list.add(new File(workpath+"/51构造深度手工铺沙法-"+sdmc+".xlsx"));
                }
            }
        }
        sdmclist.clear();
        sdmclist = jjgFbgcSdgcGssdlqlmhdzxfMapper.selectsdmc(proname,htd,"隧道工程");
        if (sdmclist.size()>0){
            for (Map<String, Object> m : sdmclist)
            {
                for (String k : m.keySet()){
                    String sdmc = m.get(k).toString();
                    list.add(new File(workpath+"/53隧道沥青路面厚度-钻芯法-"+sdmc+".xlsx"));
                }
            }
        }
        sdmclist.clear();
        sdmclist = jjgFbgcSdgcSdhntlmhdzxfMapper.selectsdmc(proname,htd);
        if (sdmclist.size()>0){
            for (Map<String, Object> m : sdmclist)
            {
                for (String k : m.keySet()){
                    String sdmc = m.get(k).toString();
                    list.add(new File(workpath+"/54隧道混凝土路面厚度-钻芯法-"+sdmc+".xlsx"));
                }
            }
        }
        sdmclist.clear();
        sdmclist = jjgFbgcSdgcSdhpMapper.selectsdmc(proname,htd,"隧道工程");
        if (sdmclist.size()>0){
            for (Map<String, Object> m : sdmclist)
            {
                for (String k : m.keySet()){
                    String sdmc = m.get(k).toString();
                    list.add(new File(workpath+"/55隧道横坡-"+sdmc+".xlsx"));
                }
            }
        }
        sdmclist.clear();
        JjgFbgcUtils.addFile(list,workpath+"/隧道工程/隧道路面",workpath);
        list.clear();
        list.add(new File(workpath+"/41隧道总体宽度.xlsx"));
        ;

        JjgFbgcUtils.addFile(list,workpath+"/隧道工程/总体",workpath);




    }

}
