package glgc.jjgys.system.service.impl;

import com.alibaba.excel.metadata.BaseRowModel;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.lmgc.*;
import glgc.jjgys.model.projectvo.sdgc.*;
import glgc.jjgys.system.mapper.JjgFbgcLmgcLmhpMapper;
import glgc.jjgys.system.service.*;
import glgc.jjgys.system.utils.JjgFbgcUtils;
import net.lingala.zip4j.core.ZipFile;
import net.lingala.zip4j.exception.ZipException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.security.core.parameters.P;
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
import java.util.Objects;

import static java.util.Arrays.asList;

@Service
public class JjgFbgcLmgcServiceImpl implements JjgFbgcLmgcService {
    @Autowired
    private JjgFbgcLmgcLmhpMapper jjgFbgcLmgcLmhpMapper;
    @Resource
    private JjgFbgcLmgcLqlmysdService jjgFbgcLmgcLqlmysdService;
    @Resource
    private JjgFbgcLmgcLmwcService jjgFbgcLmgcLmwcService;
    @Resource
    private JjgFbgcLmgcLmwcLcfService jjgFbgcLmgcLmwcLcfService;
    @Resource
    private JjgFbgcLmgcLmssxsService jjgFbgcLmgcLmssxsService;
    @Resource
    private JjgFbgcLmgcHntlmqdService jjgFbgcLmgcHntlmqdService;
    @Resource
    private  JjgFbgcLmgcTlmxlbgcService jjgFbgcLmgcTlmxlbgcService;
    @Resource
    private JjgFbgcLmgcLmgzsdsgpsfService jjgFbgcLmgcLmgzsdsgpsfService;
    @Resource
    private JjgFbgcLmgcGslqlmhdzxfService jjgFbgcLmgcGslqlmhdzxfService;
    @Resource
    private JjgFbgcLmgcHntlmhdzxfService jjgFbgcLmgcHntlmhdzxfService;
    @Resource
    private  JjgFbgcLmgcLmhpService jjgFbgcLmgcLmhpService;
    @Override
    public void exportlmgc(HttpServletResponse response,String workpath){
        String filepath = workpath;


            List<String> filenames=asList("01沥青路面压实度实测数据","02路面弯沉(贝克曼梁法)实测数据","02路面弯沉(落锤法)实测数据","04沥青路面渗水系数实测数据","05混凝土路面强度实测数据",
                    "06砼路面相邻板高差实测数据","08路面构造深度手工铺砂法实测数据","09高速沥青路面厚度钻芯法实测数据","09混凝土路面厚度钻芯法实测数据","10路面横坡实测数据"
            );

            List<BaseRowModel> models=asList(new JjgFbgcLmgcLqlmysdVo(),new JjgFbgcLmgcLmwcVo(),new JjgFbgcLmgcLmwcLcfVo(),new JjgFbgcLmgcLmssxsVo(),new JjgFbgcLmgcHntlmqdVo(),
                    new JjgFbgcLmgcTlmxlbgcVo(),new JjgFbgcLmgcLmgzsdsgpsfVo(),new JjgFbgcLmgcGslqlmhdzxfVo(),new JjgFbgcLmgcHntlmhdzxfVo(),new JjgFbgcLmgcLmhpVo());

            for(int i=0;i<filenames.size();i++){
                File file=new File(filepath+"/"+filenames.get(i)+".xlsx");

                if(!file.exists()){
                    ExcelUtil.saveLocal(filepath+"/"+filenames.get(i),models.get(i),"实测数据");
                    System.out.println("生成表格"+file.getName());
                }



            }



            JjgFbgcUtils.createDirectory("路面工程",filepath);
            JjgFbgcUtils.createDirectory("路面面层",filepath+"/路面工程");





            List<File> list=new ArrayList<>();
            list.add(new File(filepath+"/01沥青路面压实度实测数据.xlsx"));
            list.add(new File(filepath+"/02路面弯沉(贝克曼梁法)实测数据.xlsx"));
            list.add(new File(filepath+"/02路面弯沉(落锤法)实测数据.xlsx"));

            list.add(new File(filepath+"/04沥青路面渗水系数实测数据.xlsx"));
            list.add(new File(filepath+"/05混凝土路面强度实测数据.xlsx"));
            list.add(new File(filepath+"/06砼路面相邻板高差实测数据.xlsx"));
            list.add(new File(filepath+"/08路面构造深度手工铺砂法实测数据.xlsx"));
            list.add(new File(filepath+"/09高速沥青路面厚度钻芯法实测数据.xlsx"));
            list.add(new File(filepath+"/09混凝土路面厚度钻芯法实测数据.xlsx"));
            list.add(new File(filepath+"/10路面横坡实测数据.xlsx"));

            JjgFbgcUtils.addFile(list,filepath+"/路面工程/路面面层",filepath);
            list.clear();
            File file=new File( filepath);
            File[] files=file.listFiles();
            for(File f:files){
                if(!f.isDirectory()){
                    f.delete();
                }
            }









    }
    @Override
    public  void importlmgc( CommonInfoVo commonInfoVo,String workpath) {
        String zippath=workpath;



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


            commonInfoVo.setFbgc("路面工程");

            for (File f : files) {
                File[] files1 = f.listFiles();

                for (File f1 : files1) {
                    String name = f1.getName();

                    switch (name) {
                        case "01沥青路面压实度实测数据.xlsx":

                            jjgFbgcLmgcLqlmysdService.importlqlmysd(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "02路面弯沉(贝克曼梁法)实测数据.xlsx":

                            jjgFbgcLmgcLmwcService.importlmwc(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "02路面弯沉(落锤法)实测数据.xlsx":

                            jjgFbgcLmgcLmwcLcfService.importlmwclcf(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "04沥青路面渗水系数实测数据.xlsx":

                            jjgFbgcLmgcLmssxsService.importLmssxs(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "05混凝土路面强度实测数据.xlsx":

                            jjgFbgcLmgcHntlmqdService.importhntlmqd(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "06砼路面相邻板高差实测数据.xlsx":

                            jjgFbgcLmgcTlmxlbgcService.importxlbgs(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "08路面构造深度手工铺砂法实测数据.xlsx":

                            jjgFbgcLmgcLmgzsdsgpsfService.importlmgzsdsgpsf(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "09高速沥青路面厚度钻芯法实测数据.xlsx":

                            jjgFbgcLmgcGslqlmhdzxfService.importLqlmhd(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "09混凝土路面厚度钻芯法实测数据.xlsx":

                            jjgFbgcLmgcHntlmhdzxfService.importHntlmhd(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;
                        case "10路面横坡实测数据.xlsx":

                            jjgFbgcLmgcLmhpService.importlmhp(JjgFbgcUtils.getMultipartFile(f1), commonInfoVo);
                            break;

                    }
                }
            }

            JjgFbgcUtils.deleteDirAndFiles(unzipfile);



    }
    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException{
        try {
            commonInfoVo.setFbgc("路面工程");
            jjgFbgcLmgcLqlmysdService.generateJdb(commonInfoVo);
            jjgFbgcLmgcLmwcService.generateJdb(commonInfoVo);
            jjgFbgcLmgcLmwcLcfService.generateJdb(commonInfoVo);

            jjgFbgcLmgcLmssxsService.generateJdb(commonInfoVo);
            jjgFbgcLmgcHntlmqdService.generateJdb(commonInfoVo);
            jjgFbgcLmgcTlmxlbgcService.generateJdb(commonInfoVo);
            jjgFbgcLmgcLmgzsdsgpsfService.generateJdb(commonInfoVo);

            jjgFbgcLmgcGslqlmhdzxfService.generateJdb(commonInfoVo);
            jjgFbgcLmgcHntlmhdzxfService.generateJdb(commonInfoVo);

            jjgFbgcLmgcLmhpService.generateJdb(commonInfoVo);


        } catch (ParseException e) {
            throw new RuntimeException(e);
        }catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
    @Override
    public void download(HttpServletResponse response,String filename,String proname,String htd,String workpath){


        JjgFbgcUtils.createDirectory("路面工程",workpath);
        JjgFbgcUtils.createDirectory("路面面层",workpath+"/路面工程");





        List<File> list=new ArrayList<>();

        list.add(new File(workpath+"/12沥青路面压实度.xlsx"));
        list.add(new File(workpath+"/13路面弯沉(贝克曼梁法).xlsx"));
        list.add(new File(workpath+"/13路面弯沉(落锤法).xlsx"));
        list.add(new File(workpath+"/15沥青路面渗水系数.xlsx"));
        list.add(new File(workpath+"/16混凝土路面强度.xlsx"));
        list.add(new File(workpath+"/17混凝土路面相邻板高差.xlsx"));
        list.add(new File(workpath+"/20构造深度手工铺沙法.xlsx"));
        list.add(new File(workpath+"/22沥青路面厚度-钻芯法.xlsx"));
        list.add(new File(workpath+"/23混凝土路面厚度-钻芯法.xlsx"));
        list.add(new File(workpath+"/24路面横坡.xlsx"));
        List<Map<String,String>> lmlx = jjgFbgcLmgcLmhpMapper.selectlx(proname,htd); //[{lxlx=主线}, {lxlx=岳口枢纽立交}, {lxlx=延壶路小桥连接线}, {lxlx=延长互通式立交}]
        if (lmlx.size()>0){
            for (Map<String, String> map : lmlx) {
                String lxlx = map.get("lxlx");
                if(!Objects.equals(lxlx, "主线")){
                    list.add(new File(workpath+"/24路面横坡-"+lxlx+".xlsx"));
                }

            }
        }

        JjgFbgcUtils.addFile(list,workpath+"/路面工程/路面面层",workpath);
        list.clear();
        ;






    }
}
