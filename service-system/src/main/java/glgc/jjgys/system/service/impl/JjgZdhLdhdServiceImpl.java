package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.*;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.zdh.JjgZdhLdhdVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.*;
import glgc.jjgys.system.service.JjgZdhLdhdService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
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
import java.util.stream.Collectors;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-07-01
 */
@Service
public class JjgZdhLdhdServiceImpl extends ServiceImpl<JjgZdhLdhdMapper, JjgZdhLdhd> implements JjgZdhLdhdService {

    @Autowired
    private JjgZdhLdhdMapper jjgZdhLdhdMapper;

    @Autowired
    private JjgLqsSdMapper jjgLqsSdMapper;

    @Autowired
    private JjgLqsQlMapper jjgLqsQlMapper;

    @Autowired
    private JjgHtdMapper jjgHtdMapper;

    @Autowired
    private JjgLqsHntlmzdMapper jjgLqsHntlmzdMapper;

    @Autowired
    private JjgLqsSfzMapper jjgLqsSfzMapper;

    @Autowired
    private JjgLqsLjxMapper jjgLqsLjxMapper;

    @Autowired
    private JjgLqsFhlmMapper jjgLqsFhlmMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();

        List<Map<String,Object>> lxlist = jjgZdhLdhdMapper.selectlx(proname,htd);
        for (Map<String, Object> map : lxlist) {
            String zx = map.get("lxbs").toString();
            int num = jjgZdhLdhdMapper.selectcdnum(proname,htd,zx);
            int n = 0;
            if (num == 1){
                n=2;
            }else {
                n=num/2;
            }
            handlezxData(proname,htd,zx,n,commonInfoVo);
        }

    }

    /**
     *
     * @param proname
     * @param htd
     * @param zx
     * @param cdsl
     */
    private void handlezxData(String proname, String htd, String zx, int cdsl, CommonInfoVo commonInfoVo) throws IOException, ParseException {
        String[] arr = null;
        if (cdsl == 2) {
            arr = new String[] {"左幅一车道", "左幅二车道", "右幅一车道", "右幅二车道"};
        } else if (cdsl == 3) {
            arr = new String[] {"左幅一车道", "左幅二车道", "左幅三车道", "右幅一车道", "右幅二车道", "右幅三车道"};
        } else if (cdsl == 4) {
            arr = new String[] {"左幅一车道", "左幅二车道", "左幅三车道", "左幅四车道","右幅一车道", "右幅二车道", "右幅三车道", "右幅四车道"};
        } else if (cdsl == 5) {
            arr = new String[] {"左幅一车道", "左幅二车道", "左幅三车道", "左幅四车道", "左幅五车道", "右幅一车道", "右幅二车道", "右幅三车道", "右幅四车道", "右幅五车道"};
        }
        StringBuilder sb = new StringBuilder();
        for (String str : arr) {
            sb.append("\"").append(str).append("\",");
        }
        String result = sb.substring(0, sb.length() - 1); // 去掉最后一个逗号
        if (zx.equals("主线")){

            List<Map<String,Object>> datazf = jjgZdhLdhdMapper.selectzfList(proname,htd,zx,result);
            List<Map<String,Object>> datayf = jjgZdhLdhdMapper.selectyfList(proname,htd,zx,result);

            QueryWrapper<JjgHtd> wrapperhtd = new QueryWrapper<>();
            wrapperhtd.like("proname",proname);
            wrapperhtd.like("name",htd);
            List<JjgHtd> htdList = jjgHtdMapper.selectList(wrapperhtd);


            List<JjgLqsSd> jjgLqsSdzf = new ArrayList<>();
            List<JjgLqsSd> jjgLqsSdyf = new ArrayList<>();

            List<JjgLqsQl> jjgLqsQlzf = new ArrayList<>();
            List<JjgLqsQl> jjgLqsQlyf = new ArrayList<>();


            for (JjgHtd jjgHtd : htdList) {
                String zhq = jjgHtd.getZhq();
                String zhz = jjgHtd.getZhz();
                String zy = jjgHtd.getZy();
                if (zy.equals("ZY")){
                    jjgLqsSdzf.addAll(jjgLqsSdMapper.selectsdzf(proname,zhq,zhz,"左幅"));
                    jjgLqsSdyf.addAll(jjgLqsSdMapper.selectsdyf(proname,zhq,zhz,"右幅"));

                    jjgLqsQlzf.addAll(jjgLqsQlMapper.selectqlzf(proname,zhq,zhz,"左幅"));
                    jjgLqsQlyf.addAll(jjgLqsQlMapper.selectqlyf(proname,zhq,zhz,"右幅"));

                }else if (zy.equals("Z")) {
                    jjgLqsSdzf.addAll(jjgLqsSdMapper.selectsdzf(proname,zhq,zhz,"左幅"));
                    jjgLqsQlzf.addAll(jjgLqsQlMapper.selectqlzf(proname,zhq,zhz,"左幅"));
                }else if (zy.equals("Y")) {
                    jjgLqsSdyf.addAll(jjgLqsSdMapper.selectsdyf(proname,zhq,zhz,"右幅"));
                    jjgLqsQlyf.addAll(jjgLqsQlMapper.selectqlyf(proname,zhq,zhz,"右幅"));
                }
            }

            QueryWrapper<JjgLqsFhlm> wrapperfhlmzf = new QueryWrapper<>();
            wrapperfhlmzf.like("proname",proname);
            wrapperfhlmzf.like("htd",htd);
            wrapperfhlmzf.like("lf","左幅");
            List<JjgLqsFhlm> jjgLqsFhlmzf = jjgLqsFhlmMapper.selectList(wrapperfhlmzf);

            QueryWrapper<JjgLqsFhlm> wrapperfhlmyf = new QueryWrapper<>();
            wrapperfhlmyf.like("proname",proname);
            wrapperfhlmyf.like("htd",htd);
            wrapperfhlmyf.like("lf","右幅");
            List<JjgLqsFhlm> jjgLqsFhlmyf = jjgLqsFhlmMapper.selectList(wrapperfhlmyf);


            List<Map<String,Object>> hpsdzfdata = new ArrayList<>();
            if (jjgLqsSdzf.size()>0){
                for (JjgLqsSd jjgLqsSd : jjgLqsSdzf) {
                    String zhq = String.valueOf(jjgLqsSd.getZhq());
                    String zhz = String.valueOf(jjgLqsSd.getZhz());
                    hpsdzfdata.addAll(jjgZdhLdhdMapper.selectSdZfData(proname,htd,zhq,zhz,result));
                }
            }
            List<Map<String,Object>> hpsdyfdata = new ArrayList<>();
            if (jjgLqsSdyf.size()>0){
                for (JjgLqsSd jjgLqsSd : jjgLqsSdyf) {
                    String zhq = String.valueOf(jjgLqsSd.getZhq());
                    String zhz = String.valueOf(jjgLqsSd.getZhz());
                    hpsdyfdata.addAll(jjgZdhLdhdMapper.selectSdyfData(proname,htd,zhq,zhz,result));
                }
            }
            List<Map<String,Object>> hpqlzfdata = new ArrayList<>();
            if (jjgLqsQlzf.size()>0){
                for (JjgLqsQl jjgLqsQl : jjgLqsQlzf) {
                    String zhq = String.valueOf(jjgLqsQl.getZhq());
                    String zhz = String.valueOf(jjgLqsQl.getZhz());
                    hpqlzfdata.addAll(jjgZdhLdhdMapper.selectQlZfData(proname,htd,zhq,zhz,result));
                }
            }
            List<Map<String,Object>> hpqlyfdata = new ArrayList<>();
            if (jjgLqsQlyf.size()>0){
                for (JjgLqsQl jjgLqsQl : jjgLqsQlyf) {
                    String zhq = String.valueOf(jjgLqsQl.getZhq());
                    String zhz = String.valueOf(jjgLqsQl.getZhz());
                    hpqlyfdata.addAll(jjgZdhLdhdMapper.selectQlYfData(proname,htd,zhq,zhz,result));
                }
            }

            List<Map<String,Object>> flmzfdata = new ArrayList<>();
            if (jjgLqsFhlmzf.size()>0){
                for (JjgLqsFhlm jjgLqsFhlm : jjgLqsFhlmzf) {
                    String zhq = String.valueOf(jjgLqsFhlm.getZhq());
                    String zhz = String.valueOf(jjgLqsFhlm.getZhz());
                    flmzfdata.addAll(jjgZdhLdhdMapper.seletcfhlmzfData(proname,htd,zhq,zhz));

                }

            }
            List<Map<String,Object>> flmyfdata = new ArrayList<>();

            if (jjgLqsFhlmyf.size()>0){
                for (JjgLqsFhlm jjgLqsFhlm : jjgLqsFhlmyf) {
                    String zhq = String.valueOf(jjgLqsFhlm.getZhq());
                    String zhz = String.valueOf(jjgLqsFhlm.getZhz());
                    flmyfdata.addAll(jjgZdhLdhdMapper.seletcfhlmyfData(proname,htd,zhq,zhz));

                }

            }

            //处理数据
            List<Map<String, Object>> sdzxList = montageld(hpsdzfdata);
            List<Map<String, Object>> sdyxList = montageld(hpsdyfdata);

            List<Map<String, Object>> qlzxList = montageld(hpqlzfdata);
            List<Map<String, Object>> qlyxList = montageld(hpqlyfdata);
            List<Map<String, Object>> flmzxList = montageld(flmzfdata);
            List<Map<String, Object>> flmyxList = montageld(flmyfdata);

            List<Map<String, Object>> lmzfList = montageld(datazf);
            List<Map<String, Object>> lmyfList = montageld(datayf);

            writeExcelData(proname,htd,lmzfList,lmyfList,sdzxList,sdyxList,qlzxList,qlyxList,flmzxList,flmyxList,cdsl,zx,commonInfoVo);
        }else if (zx.contains("连接线")){
            List<Map<String,Object>> dataljxzf = jjgZdhLdhdMapper.selectzfList(proname,htd,zx,result);
            List<Map<String,Object>> dataljxyf = jjgZdhLdhdMapper.selectyfList(proname,htd,zx,result);

            //连接线
            QueryWrapper<JjgLjx> wrapperljx = new QueryWrapper<>();
            wrapperljx.like("proname",proname);
            wrapperljx.like("sshtd",htd);
            List<JjgLjx> jjgLjxList = jjgLqsLjxMapper.selectList(wrapperljx);

            List<Map<String,Object>> sdldhd = new ArrayList<>();
            List<Map<String,Object>> qlldhd = new ArrayList<>();

            for (JjgLjx jjgLjx : jjgLjxList) {
                String zhq = jjgLjx.getZhq();
                String zhz = jjgLjx.getZhz();
                String bz = jjgLjx.getBz();
                String ljxlf = jjgLjx.getLf();
                String wz = jjgLjx.getLjxname();
                List<JjgLqsSd> jjgLqssd = jjgLqsSdMapper.selectsdList(proname,zhq,zhz,bz,wz,ljxlf);
                //有可能是单幅，有可能是左右幅都有
                for (JjgLqsSd jjgLqsSd : jjgLqssd) {
                    String lf = jjgLqsSd.getLf();

                    String sdzhq = String.valueOf((jjgLqsSd.getZhq()));
                    String sdzhz = String.valueOf((jjgLqsSd.getZhz()));
                    sdldhd.addAll(jjgZdhLdhdMapper.selectsdldhd(proname,bz,lf,zx,sdzhq,sdzhz));
                }

                List<JjgLqsQl> jjgLqsql = jjgLqsQlMapper.selectqlList(proname,zhq,zhz,bz,wz,ljxlf);
                for (JjgLqsQl jjgLqsQl : jjgLqsql) {
                    String lf = jjgLqsQl.getLf();

                    String qlzhq = String.valueOf(jjgLqsQl.getZhq());
                    String qlzhz = String.valueOf(jjgLqsQl.getZhz());
                    qlldhd.addAll(jjgZdhLdhdMapper.selectqlldhd(proname,bz,lf,zx,qlzhq,qlzhz));

                }
            }

            List<Map<String,Object>> zfqlldhd = new ArrayList<>();
            List<Map<String,Object>> yfqlldhd = new ArrayList<>();
            if (qlldhd.size()>0){
                for (int i = 0; i < qlldhd.size(); i++) {
                    if (qlldhd.get(i).get("cd").toString().contains("左幅")){
                        zfqlldhd.add(qlldhd.get(i));
                    }
                    if (qlldhd.get(i).get("cd").toString().contains("右幅")){
                        yfqlldhd.add(qlldhd.get(i));
                    }
                }
            }
            List<Map<String,Object>> zfsdldhd = new ArrayList<>();
            List<Map<String,Object>> yfsdldhd = new ArrayList<>();
            if (sdldhd.size()>0){
                for (Map<String, Object> sdmcx : sdldhd) {
                    if (sdmcx.get("cd").toString().contains("左幅")){
                        zfsdldhd.add(sdmcx);
                    }
                    if (sdmcx.get("cd").toString().contains("右幅")){
                        yfsdldhd.add(sdmcx);
                    }
                }
            }
            QueryWrapper<JjgLqsFhlm> wrapperfhlmzf = new QueryWrapper<>();
            wrapperfhlmzf.like("proname",proname);
            wrapperfhlmzf.like("name",htd);
            wrapperfhlmzf.like("lf","左幅");
            List<JjgLqsFhlm> jjgLqsFhlmzf = jjgLqsFhlmMapper.selectList(wrapperfhlmzf);

            QueryWrapper<JjgLqsFhlm> wrapperfhlmyf = new QueryWrapper<>();
            wrapperfhlmyf.like("proname",proname);
            wrapperfhlmyf.like("name",htd);
            wrapperfhlmyf.like("lf","右幅");
            List<JjgLqsFhlm> jjgLqsFhlmyf = jjgLqsFhlmMapper.selectList(wrapperfhlmyf);

            List<Map<String,Object>> flmzfdata = new ArrayList<>();
            if (jjgLqsFhlmzf.size()>0){
                for (JjgLqsFhlm jjgLqsFhlm : jjgLqsFhlmzf) {
                    String zhq = String.valueOf(jjgLqsFhlm.getZhq());
                    String zhz = String.valueOf(jjgLqsFhlm.getZhz());
                    flmzfdata.addAll(jjgZdhLdhdMapper.seletcfhlmzfData(proname,htd,zhq,zhz));

                }

            }
            List<Map<String,Object>> flmyfdata = new ArrayList<>();

            if (jjgLqsFhlmyf.size()>0){
                for (JjgLqsFhlm jjgLqsFhlm : jjgLqsFhlmyf) {
                    String zhq = String.valueOf(jjgLqsFhlm.getZhq());
                    String zhz = String.valueOf(jjgLqsFhlm.getZhz());
                    flmyfdata.addAll(jjgZdhLdhdMapper.seletcfhlmyfData(proname,htd,zhq,zhz));

                }

            }

            List<Map<String, Object>> qlzfsj = montageld(zfqlldhd);
            List<Map<String, Object>> qlyfsj = montageld(yfqlldhd);

            List<Map<String, Object>> sdzfsj = montageld(zfsdldhd);
            List<Map<String, Object>> sdyfsj = montageld(yfsdldhd);

            List<Map<String, Object>> allzfsj = montageld(dataljxzf);
            List<Map<String, Object>> allyfsj = montageld(dataljxyf);

            writeExcelData(proname,htd,allzfsj,allyfsj,sdzfsj,sdyfsj,qlzfsj,qlyfsj,flmzfdata,flmyfdata,cdsl,zx,commonInfoVo);


        } else {
            //匝道
            List<Map<String,Object>> datazdzf = jjgZdhLdhdMapper.selectzfList(proname,htd,zx,result);
            List<Map<String,Object>> datazdyf = jjgZdhLdhdMapper.selectyfList(proname,htd,zx,result);

            QueryWrapper<JjgLqsHntlmzd> wrapperzd = new QueryWrapper<>();
            wrapperzd.like("proname",proname);
            wrapperzd.like("wz",zx);
            List<JjgLqsHntlmzd> zdList = jjgLqsHntlmzdMapper.selectList(wrapperzd);

            List<Map<String,Object>> sdldhd = new ArrayList<>();//左右幅或者单幅，当前匝道下全部的隧道数据
            List<Map<String,Object>> qlldhd = new ArrayList<>();

            //匝道的所有数据
            for (JjgLqsHntlmzd jjgLqsHntlmzd : zdList) {
                String zhq = jjgLqsHntlmzd.getZhq();
                String zhz = jjgLqsHntlmzd.getZhz();
                String bz = jjgLqsHntlmzd.getZdlx();
                String wz = jjgLqsHntlmzd.getWz();
                String zdlf = jjgLqsHntlmzd.getLf();
                List<JjgLqsSd> jjgLqssd = jjgLqsSdMapper.selectsdList(proname,zhq,zhz,bz,wz,zdlf);
                //有可能是单幅，有可能是左右幅都有
                for (JjgLqsSd jjgLqsSd : jjgLqssd) {
                    String lf = jjgLqsSd.getLf();
                    String sdzhq = String.valueOf((jjgLqsSd.getZhq()));
                    String sdzhz = String.valueOf((jjgLqsSd.getZhz()));

                    sdldhd.addAll(jjgZdhLdhdMapper.selectsdldhd(proname,bz,lf,zx,sdzhq,sdzhz));
                }

                List<JjgLqsQl> jjgLqsql = jjgLqsQlMapper.selectqlList(proname,zhq,zhz,bz,wz,zdlf);
                for (JjgLqsQl jjgLqsQl : jjgLqsql) {
                    String lf = jjgLqsQl.getLf();
                    String qlzhq = String.valueOf(jjgLqsQl.getZhq());
                    String qlzhz = String.valueOf(jjgLqsQl.getZhz());
                    qlldhd.addAll(jjgZdhLdhdMapper.selectqlldhd(proname,bz,lf, zx, qlzhq, qlzhz));
                }
            }
            QueryWrapper<JjgLqsFhlm> wrapperfhlmzf = new QueryWrapper<>();
            wrapperfhlmzf.like("proname",proname);
            wrapperfhlmzf.like("name",htd);
            wrapperfhlmzf.like("lf","左幅");
            List<JjgLqsFhlm> jjgLqsFhlmzf = jjgLqsFhlmMapper.selectList(wrapperfhlmzf);

            QueryWrapper<JjgLqsFhlm> wrapperfhlmyf = new QueryWrapper<>();
            wrapperfhlmyf.like("proname",proname);
            wrapperfhlmyf.like("name",htd);
            wrapperfhlmyf.like("lf","右幅");
            List<JjgLqsFhlm> jjgLqsFhlmyf = jjgLqsFhlmMapper.selectList(wrapperfhlmyf);

            List<Map<String,Object>> flmzfdata = new ArrayList<>();
            if (jjgLqsFhlmzf.size()>0){
                for (JjgLqsFhlm jjgLqsFhlm : jjgLqsFhlmzf) {
                    String zhq = String.valueOf(jjgLqsFhlm.getZhq());
                    String zhz = String.valueOf(jjgLqsFhlm.getZhz());
                    flmzfdata.addAll(jjgZdhLdhdMapper.seletcfhlmzfData(proname,htd,zhq,zhz));

                }

            }
            List<Map<String,Object>> flmyfdata = new ArrayList<>();

            if (jjgLqsFhlmyf.size()>0){
                for (JjgLqsFhlm jjgLqsFhlm : jjgLqsFhlmyf) {
                    String zhq = String.valueOf(jjgLqsFhlm.getZhq());
                    String zhz = String.valueOf(jjgLqsFhlm.getZhz());
                    flmyfdata.addAll(jjgZdhLdhdMapper.seletcfhlmyfData(proname,htd,zhq,zhz));

                }

            }

            List<Map<String, Object>> sdmontageld = montageldzd(sdldhd);
            List<Map<String, Object>> qlmontageld = montageldzd(qlldhd);

            List<Map<String, Object>> fhlmzmontageld = montageldzd(flmzfdata);
            List<Map<String, Object>> fhlmymontageld = montageldzd(flmyfdata);

            List<Map<String, Object>> zdzfmontageld = montageldzd(datazdzf);
            List<Map<String, Object>> zdyfmontageld = montageldzd(datazdyf);

            zdwriteExcelData(proname,htd,zdzfmontageld,zdyfmontageld,sdmontageld,qlmontageld,fhlmzmontageld,fhlmymontageld,cdsl,commonInfoVo,zx);

        }

    }

    /**
     *
     * @param proname
     * @param htd
     * @param zdzfList
     * @param zdyfList
     * @param sdList
     * @param qlList
     * @param fhlmz
     * @param fhlmy
     * @param cdsl
     * @param zx
     * @throws IOException
     * @throws ParseException
     */
    private void zdwriteExcelData(String proname, String htd, List<Map<String, Object>> zdzfList, List<Map<String, Object>> zdyfList, List<Map<String, Object>> sdList, List<Map<String, Object>> qlList, List<Map<String, Object>> fhlmz, List<Map<String, Object>> fhlmy, int cdsl, CommonInfoVo commonInfoVo, String zx) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        String fname = "16路面雷达厚度-"+zx+".xlsx";

        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+fname);
        File fdir = new File(filepath + File.separator + proname + File.separator + htd);
        if (!fdir.exists()) {
            //创建文件根目录
            fdir.mkdirs();
        }
        File directory = new File("src/main/resources/static");
        String reportPath = directory.getCanonicalPath();
        String filename = "";

        if (cdsl == 5){
            filename = "雷达厚度-5车道.xlsx";
        }else if (cdsl == 4){
            filename = "雷达厚度-4车道.xlsx";
        }else if (cdsl == 3){
            filename = "雷达厚度-3车道.xlsx";
        }else if (cdsl == 2){
            filename = "雷达厚度-2车道.xlsx";
        }


        String path = reportPath + File.separator + filename;
        Files.copy(Paths.get(path), new FileOutputStream(f));
        FileInputStream out = new FileInputStream(f);
        wb = new XSSFWorkbook(out);

        List<Map<String,Object>> sdqllmData = new ArrayList<>();

        if (sdList.size() >0 && !sdList.isEmpty()){
            for (Map<String, Object> map : sdList) {
                Map<String,Object> map1 = new HashMap<>();
                map1.put("zh",map.get("zh").toString());
                map1.put("name",map.get("name").toString());
                map1.put("createTime",map.get("createTime").toString());
                map1.put("cd",map.get("cd").toString());
                if (map.get("zdbs")!=null){
                    map1.put("zdbs",map.get("zdbs").toString());
                }
                sdqllmData.add(map1);
            }
            List<Map<String, Object>> mapslist = mergedList(sdList,cdsl);
            DBtoExcel(proname,htd,mapslist,wb,"匝道隧道",cdsl,commonInfoVo.getSdsjz(),zx);
        }

        if (qlList.size() >0 && !qlList.isEmpty()){
            for (Map<String, Object> map : qlList) {
                Map<String,Object> map1 = new HashMap<>();
                map1.put("zh",map.get("zh").toString());
                map1.put("name",map.get("name").toString());
                map1.put("createTime",map.get("createTime").toString());
                map1.put("cd",map.get("cd").toString());
                if (map.get("zdbs")!=null){
                    map1.put("zdbs",map.get("zdbs").toString());
                }
                sdqllmData.add(map1);
            }
            List<Map<String, Object>> mapslist = mergedList(qlList,cdsl);
            DBtoExcel(proname,htd,mapslist,wb,"匝道桥",cdsl,commonInfoVo.getQlsjz(),zx);
        }


        if (fhlmz.size()>0 && !fhlmz.isEmpty()){
            for (Map<String, Object> map : fhlmz) {
                Map<String,Object> map1 = new HashMap<>();
                map1.put("zh",map.get("zh").toString());
                map1.put("name",map.get("name").toString());
                map1.put("cd","左幅");
                if (map.get("zdbs")!=null){
                    map1.put("zdbs",map.get("zdbs").toString());
                }
                sdqllmData.add(map1);
            }
            DBtoExcel(proname,htd,fhlmz,wb,"复合路面左幅",cdsl,commonInfoVo.getFhlmsjz(),zx);
        }

        if (fhlmy.size()>0 && !fhlmy.isEmpty()){
            for (Map<String, Object> map : fhlmy) {
                Map<String,Object> map1 = new HashMap<>();
                map1.put("zh",map.get("zh").toString());
                map1.put("name",map.get("name").toString());
                map1.put("cd","左幅");
                if (map.get("zdbs")!=null){
                    map1.put("zdbs",map.get("zdbs").toString());
                }
                sdqllmData.add(map1);
            }
            DBtoExcel(proname,htd,fhlmy,wb,"复合路面右幅",cdsl,commonInfoVo.getFhlmsjz(),zx);
        }

        if (zdzfList.size()>0 && !zdzfList.isEmpty()){
            DBtoExcelZd(proname,htd,zdzfList,sdqllmData,wb,"匝道左幅",cdsl,commonInfoVo.getSjz(),zx);
        }

        if (zdyfList.size()>0 && !zdyfList.isEmpty()){
            DBtoExcelZd(proname,htd,zdyfList,sdqllmData,wb,"匝道右幅",cdsl,commonInfoVo.getSjz(),zx);

        }

        String[] arr = {"匝道桥","匝道隧道","匝道左幅","匝道右幅","左幅","右幅","桥","隧道","复合路面左幅","复合路面右幅"};
        for (int i = 0; i < arr.length; i++) {

            if (shouldBeCalculate(wb.getSheet(arr[i]))) {
                if (arr[i].equals("匝道左幅") || arr[i].equals("匝道右幅")){
                    calculateBridgeAndTunnelSheet(wb,wb.getSheet(arr[i]),cdsl,commonInfoVo.getSjz());
                }else if (arr[i].equals("匝道桥")) {
                    calculateBridgeAndTunnelSheet(wb,wb.getSheet(arr[i]),cdsl,commonInfoVo.getQlsjz());
                }else if (arr[i].equals("匝道隧道")) {
                    calculateBridgeAndTunnelSheet(wb,wb.getSheet(arr[i]),cdsl,commonInfoVo.getSdsjz());
                }else if (arr[i].contains("复合路面")) {
                    calculateBridgeAndTunnelSheet(wb,wb.getSheet(arr[i]),cdsl,commonInfoVo.getFhlmsjz());
                }
            }else {
                wb.setSheetHidden(wb.getSheetIndex(arr[i]),true);
            }
        }

        FileOutputStream fileOut = new FileOutputStream(f);
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
        out.close();
        wb.close();

    }


    /**
     *
     * @param proname
     * @param htd
     * @param lmzfList
     * @param lmyfList
     * @param sdzxList
     * @param sdyxList
     * @param qlzxList
     * @param qlyxList
     * @param flmzxList
     * @param flmyxList
     * @param cdsl
     * @param zx
     */
    private void writeExcelData(String proname, String htd, List<Map<String, Object>> lmzfList, List<Map<String, Object>> lmyfList, List<Map<String, Object>> sdzxList, List<Map<String, Object>> sdyxList, List<Map<String, Object>> qlzxList, List<Map<String, Object>> qlyxList, List<Map<String, Object>> flmzxList, List<Map<String, Object>> flmyxList, int cdsl, String zx,CommonInfoVo commonInfoVo) throws IOException, ParseException {
        XSSFWorkbook wb = null;

        String fname="";
        if (zx.equals("主线")){
            fname = "16路面雷达厚度.xlsx";
        }else {
            fname = "16路面雷达厚度-"+zx+".xlsx";
        }

        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+fname);
        File fdir = new File(filepath + File.separator + proname + File.separator + htd);
        if (!fdir.exists()) {
            //创建文件根目录
            fdir.mkdirs();
        }
        File directory = new File("src/main/resources/static");
        String reportPath = directory.getCanonicalPath();
        String filename = "";

        if (cdsl == 5){
            filename = "雷达厚度-5车道.xlsx";
        }else if (cdsl == 4){
            filename = "雷达厚度-4车道.xlsx";
        }else if (cdsl == 3){
            filename = "雷达厚度-3车道.xlsx";
        }else if (cdsl == 2){
            filename = "雷达厚度-2车道.xlsx";
        }


        String path = reportPath + File.separator + filename;
        Files.copy(Paths.get(path), new FileOutputStream(f));
        FileInputStream out = new FileInputStream(f);
        wb = new XSSFWorkbook(out);

        List<Map<String,Object>> sdqllmData = new ArrayList<>();

        sdzxList.addAll(sdyxList);
        if (sdzxList.size() >0 && !sdzxList.isEmpty()){
            for (Map<String, Object> map : sdzxList) {
                Map<String,Object> map1 = new HashMap<>();
                map1.put("zh",map.get("zh").toString());
                map1.put("name",map.get("name").toString());
                map1.put("createTime",map.get("createTime").toString());
                map1.put("cd",map.get("cd").toString());
                if (map.get("zdbs")!=null){
                    map1.put("zdbs",map.get("zdbs").toString());
                }
                sdqllmData.add(map1);
            }
            List<Map<String, Object>> mapslist = mergedList(sdzxList,cdsl);
            DBtoExcel(proname,htd,mapslist,wb,"隧道",cdsl, commonInfoVo.getSdsjz(),zx);
        }

        qlzxList.addAll(qlyxList);
        if (qlzxList.size() >0 && !qlzxList.isEmpty()){
            for (Map<String, Object> map : qlzxList) {
                Map<String,Object> map1 = new HashMap<>();
                map1.put("zh",map.get("zh").toString());
                map1.put("name",map.get("name").toString());
                map1.put("createTime",map.get("createTime").toString());
                map1.put("cd",map.get("cd").toString());
                if (map.get("zdbs")!=null){
                    map1.put("zdbs",map.get("zdbs").toString());
                }
                sdqllmData.add(map1);
            }
            List<Map<String, Object>> mapslist = mergedList(qlzxList,cdsl);
            DBtoExcel(proname,htd,mapslist,wb,"桥",cdsl,commonInfoVo.getQlsjz(),zx);
        }

        if (flmzxList.size()>0 && !flmzxList.isEmpty()){
            for (Map<String, Object> map : flmzxList) {
                Map<String,Object> map1 = new HashMap<>();
                map1.put("zh",map.get("zh").toString());
                map1.put("name",map.get("name").toString());
                map1.put("cd",map.get("cd").toString());
                if (map.get("zdbs")!=null){
                    map1.put("zdbs",map.get("zdbs").toString());
                }
                sdqllmData.add(map1);
            }
            List<Map<String, Object>> mapslist = mergedList(flmzxList,cdsl);
            DBtoExcel(proname,htd,mapslist,wb,"复合路面左幅",cdsl,commonInfoVo.getFhlmsjz(),zx);
        }

        if (flmyxList.size()>0 && !flmyxList.isEmpty()){
            for (Map<String, Object> map : flmyxList) {
                Map<String,Object> map1 = new HashMap<>();
                map1.put("zh",map.get("zh").toString());
                map1.put("name",map.get("name").toString());
                map1.put("cd",map.get("cd").toString());
                if (map.get("zdbs")!=null){
                    map1.put("zdbs",map.get("zdbs").toString());
                }
                sdqllmData.add(map1);
            }
            List<Map<String, Object>> mapslist = mergedList(flmyxList,cdsl);
            DBtoExcel(proname,htd,mapslist,wb,"复合路面右幅",cdsl,commonInfoVo.getFhlmsjz(),zx);
        }

        if (lmzfList.size()>0 && !lmzfList.isEmpty()){
            DBtoExcelLM(proname,htd,lmzfList,sdqllmData,wb,"左幅",cdsl,commonInfoVo.getSjz(),zx);
        }

        if (lmyfList.size()>0 && !lmyfList.isEmpty()){
            DBtoExcelLM(proname,htd,lmyfList,sdqllmData,wb,"右幅",cdsl,commonInfoVo.getSjz(),zx);
        }

        String[] arr = {"匝道桥","匝道隧道","匝道左幅","匝道右幅","左幅","右幅","桥","隧道","复合路面左幅","复合路面右幅"};
        for (int i = 0; i < arr.length; i++) {
            if (shouldBeCalculate(wb.getSheet(arr[i]))) {
                if (arr[i].equals("左幅") || arr[i].equals("右幅")){
                    calculatePavementSheet(wb,wb.getSheet(arr[i]),cdsl);
                }else if (arr[i].equals("桥")) {
                    calculateBridgeAndTunnelSheet(wb,wb.getSheet(arr[i]),cdsl,commonInfoVo.getQlsjz());
                }else if (arr[i].equals("隧道")) {
                    calculateBridgeAndTunnelSheet(wb,wb.getSheet(arr[i]),cdsl,commonInfoVo.getSdsjz());
                }else if (arr[i].contains("复合路面")) {
                    calculateBridgeAndTunnelSheet(wb,wb.getSheet(arr[i]),cdsl,commonInfoVo.getFhlmsjz());
                }

            }else {
                wb.setSheetHidden(wb.getSheetIndex(arr[i]),true);
            }
        }

        FileOutputStream fileOut = new FileOutputStream(f);
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
        out.close();
        wb.close();

    }

    /**
     *桥梁隧道
     * @param sheet
     */
    private void calculateBridgeAndTunnelSheet(XSSFWorkbook wb,XSSFSheet sheet,int cdsl,String sjz){
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        boolean flag = false;
        FormulaEvaluator e = new XSSFFormulaEvaluator(wb);
        String name = "";
        int mm = 0;
        if (cdsl == 2){
            mm = 41;

        }else if (cdsl ==3 ){
            mm = 41;
        }else if (cdsl == 4){
            mm = 22;

        }else if (cdsl == 5){
            mm = 26;

        }
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }

            if(flag){
                if(name.equals(row.getCell(0).toString())){
                    rowend = sheet.getRow(i+41);
                }else {
                    fillBridgeAndTunnelTotal(wb,sheet, rowstart, rowend,cdsl,sjz);
                    rowstart = sheet.getRow(i);
                    rowend = sheet.getRow(i+mm);
                    name = rowstart.getCell(0).toString();
                }
                /*rowstart = sheet.getRow(i);
                rowend = sheet.getRow(i+41);
                fillBridgeAndTunnelTotal(wb,sheet, rowstart, rowend,cdsl);
                i += 41;*/
                i += mm;
            }
            if ("工程名称".equals(row.getCell(0).toString())) {
                i+=1;
                rowstart = sheet.getRow(i+1);
                rowend = sheet.getRow(i+(mm+1));
                name = rowstart.getCell(0).toString();
                fillBridgeAndTunnelTotal(wb,sheet, rowstart, rowend,cdsl,sjz);
                i+= (mm+1);
                flag = true;
            }
        }
        //fillBridgeAndTunnelTotal(wb,sheet, rowstart, rowend,cdsl,sjz);

        double value= 0 ;

        sheet.getRow(4).createCell(cdsl*3+7).setCellFormula("SUM("
                +sheet.getRow(5).createCell(cdsl*3+7).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(cdsl*3+7).getReference()+")");

        sheet.getRow(4).createCell(cdsl*3+8).setCellFormula("SUM("
                +sheet.getRow(5).createCell(cdsl*3+8).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(cdsl*3+8).getReference()+")");

        sheet.getRow(4).createCell(cdsl*3+9).setCellFormula(
                sheet.getRow(4).getCell(cdsl*3+8).getReference()+"*100/"+
                        sheet.getRow(4).getCell(cdsl*3+7).getReference());
        //最大值，最小值
        sheet.getRow(4).createCell(cdsl*3+4).setCellFormula("MAX("
                +sheet.getRow(5).createCell(cdsl*3+4).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(cdsl*3+4).getReference()+")");
        value = e.evaluate(sheet.getRow(4).getCell(cdsl*3+4)).getNumberValue();
        sheet.getRow(4).getCell(cdsl*3+4).setCellFormula(null);
        sheet.getRow(4).getCell(cdsl*3+4).setCellValue(value);

        sheet.getRow(4).createCell(cdsl*3+5).setCellFormula("MIN("
                +sheet.getRow(5).createCell(cdsl*3+5).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(cdsl*3+5).getReference()+")");
        value = e.evaluate(sheet.getRow(4).getCell(cdsl*3+5)).getNumberValue();
        sheet.getRow(4).getCell(cdsl*3+5).setCellFormula(null);
        sheet.getRow(4).getCell(cdsl*3+5).setCellValue(value);

        sheet.getRow(5).createCell(cdsl*3+7).setCellFormula(sheet.getRow(cdsl*2+1).getCell(cdsl*2+4).getReference());
        value = e.evaluate(sheet.getRow(5).getCell(cdsl*3+7)).getNumberValue();
        sheet.getRow(5).getCell(cdsl*3+7).setCellFormula(null);
        sheet.getRow(5).getCell(cdsl*3+7).setCellValue(value);

        sheet.getRow(5).createCell(cdsl*3+8).setCellFormula(sheet.getRow(cdsl*2+2).getCell(cdsl*2+4).getReference());
        value = e.evaluate(sheet.getRow(5).getCell(cdsl*3+8)).getNumberValue();
        sheet.getRow(5).getCell(cdsl*3+8).setCellFormula(null);
        sheet.getRow(5).getCell(cdsl*3+8).setCellValue(value);

        int hgds = 0;
        for (int i = 5; i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null || row.getCell(cdsl*3+8) == null || "".equals(row.getCell(cdsl*3+8).toString())) {
                continue;
            }

            hgds += Double.valueOf(e.evaluate(row.getCell(cdsl*3+8)).getNumberValue()).intValue();
        }

        sheet.getRow(4).getCell(cdsl*3+8).setCellFormula(null);
        sheet.getRow(4).getCell(cdsl*3+8).setCellValue(hgds);
    }

    /**
     *
     * @param wb
     * @param sheet
     * @param rowstart
     * @param rowend
     */
    private void fillBridgeAndTunnelTotal(XSSFWorkbook wb,XSSFSheet sheet, XSSFRow rowstart,XSSFRow rowend,int cdsl,String sjz) {
        sheet.getRow(rowstart.getRowNum()).getCell(cdsl*2+2).setCellValue("代表值允许偏差（cm）");
        sheet.getRow(rowstart.getRowNum()+1).getCell(cdsl*2+2).setCellValue("合格值允许偏差（cm）");
        sheet.getRow(rowstart.getRowNum()+2).getCell(cdsl*2+2).setCellValue("总点数");
        sheet.getRow(rowstart.getRowNum()+3).getCell(cdsl*2+2).setCellValue("合格数");
        sheet.getRow(rowstart.getRowNum()+4).getCell(cdsl*2+2).setCellValue("设计值(cm)");
        sheet.getRow(rowstart.getRowNum()+5).getCell(cdsl*2+2).setCellValue("平均值(cm)");
        sheet.getRow(rowstart.getRowNum()+6).getCell(cdsl*2+2).setCellValue("代表值(cm)");
        sheet.getRow(rowstart.getRowNum()+7).getCell(cdsl*2+2).setCellValue("均方差");
        sheet.getRow(rowstart.getRowNum()+8).getCell(cdsl*2+2).setCellValue("合格率");
        sheet.getRow(rowstart.getRowNum()+9).getCell(cdsl*2+2).setCellValue("评定结果");

        sheet.getRow(rowstart.getRowNum()).getCell(cdsl*2+4).setCellFormula("-0.05*"+
                sheet.getRow(rowstart.getRowNum()+4).getCell(cdsl*2+4).getReference());//代表值允许偏差（cm）=-0.08*I39
        sheet.getRow(rowstart.getRowNum()+1).getCell(cdsl*2+4).setCellFormula("-0.1*"+
                sheet.getRow(rowstart.getRowNum()+4).getCell(cdsl*2+4).getReference());//合格值允许偏差（cm）=-0.1*I39

        //总点数
        /**cdsl*2+1
         * 2c 5
         * 3c 7
         * 4c 9
         * 5c 11
         */
        System.out.println(rowstart.getCell(2).getReference()+":"+rowend.getCell(cdsl*2+1).getReference());
        sheet.getRow(rowstart.getRowNum()+2).getCell(cdsl*2+4).setCellFormula("COUNT("
                +rowstart.getCell(2).getReference()+":"
                +rowend.getCell(cdsl*2+1).getReference()+")");//I37=COUNT(C6:F44)

        XSSFFormulaEvaluator e=new XSSFFormulaEvaluator(wb);

        Cell cell = sheet.getRow(rowstart.getRowNum()+4).getCell(cdsl*2+4);
        double data = 0;
        if (cell.getCellTypeEnum() == CellType.NUMERIC) {
            data = cell.getNumericCellValue() + e.evaluate(sheet.getRow(rowstart.getRowNum()+1).getCell(cdsl*2+4)).getNumberValue();
        }

        //合格数
        //=IF(K15="合格",(K8-SUM(COUNTIF(C6:H89,"<9.0"))),"")
        sheet.getRow(rowstart.getRowNum()+3).getCell(cdsl*2+4).setCellFormula("IF("
                +sheet.getRow(rowstart.getRowNum()+9).getCell(cdsl*2+4).getReference()+"=\"合格\",("
                +sheet.getRow(rowstart.getRowNum()+2).getCell(cdsl*2+4).getReference()+"-SUM(COUNTIF("
                +rowstart.getCell(2).getReference()+":"
                +rowend.getCell(cdsl*2+1).getReference()+",\"<"+ data +"\"))),\"-\")");

        sheet.getRow(rowstart.getRowNum()+4).getCell(cdsl*2+4).setCellValue(Double.parseDouble(sjz));

        //平均值
        sheet.getRow(rowstart.getRowNum()+5).getCell(cdsl*2+4).setCellFormula("IFERROR(AVERAGE("
                +rowstart.getCell(2).getReference()+":"
                +rowend.getCell(cdsl*2+1).getReference()+"),\"-\")");//I40=AVERAGE(B6:C44,E6:F44,H6:I27)

        //均方差
        sheet.getRow(rowstart.getRowNum()+7).getCell(cdsl*2+4).setCellFormula("IFERROR(STDEV("
                +rowstart.getCell(2).getReference()+":"
                +rowend.getCell(cdsl*2+1).getReference()+"),\"-\")");

        //代表值(cm)
        sheet.getRow(rowstart.getRowNum()+6).getCell(cdsl*2+4).setCellFormula("IFERROR(IF("
                +sheet.getRow(rowstart.getRowNum()+2).getCell(cdsl*2+4).getReference()+">100,("
                +sheet.getRow(rowstart.getRowNum()+5).getCell(cdsl*2+4).getReference()+"-"
                +sheet.getRow(rowstart.getRowNum()+7).getCell(cdsl*2+4).getReference()+"*1.6449/SQRT("
                +sheet.getRow(rowstart.getRowNum()+2).getCell(cdsl*2+4).getReference()+")),("
                +sheet.getRow(rowstart.getRowNum()+5).getCell(cdsl*2+4).getReference()+"-VLOOKUP("
                +sheet.getRow(rowstart.getRowNum()+2).getCell(cdsl*2+4).getReference()+",Sheet1!A:D,"+(3)+",FALSE)*"
                +sheet.getRow(rowstart.getRowNum()+7).getCell(cdsl*2+4).getReference()+")),\"-\")");
        //I41=IF(I37>100,(I40-I42*1.6449/SQRT(I37)),(I40-VLOOKUP(I37,Sheet1!A:D,3,FALSE)*I42))

        //将代表值放到右边的最大值最小值位置方便在报告中
        sheet.getRow(rowstart.getRowNum()+2).createCell(cdsl*3+4).setCellFormula(sheet.getRow(rowstart.getRowNum()+6).getCell(cdsl*2+4).getReference());
        sheet.getRow(rowstart.getRowNum()+2).createCell(cdsl*3+5).setCellFormula(sheet.getRow(rowstart.getRowNum()+6).getCell(cdsl*2+4).getReference());

        double totalNum = e.evaluate(sheet.getRow(rowstart.getRowNum()+2).getCell(cdsl*2+4)).getNumberValue();//总点数不为0的时候才计算最大值最小值
        if(totalNum != 0){
            //评定结果
            sheet.getRow(rowstart.getRowNum()+9).getCell(cdsl*2+4).setCellFormula("IFERROR(IF("+
                    sheet.getRow(rowstart.getRowNum()+6).getCell(cdsl*2+4).getReference()+">="+
                    sheet.getRow(rowstart.getRowNum()+4).getCell(cdsl*2+4).getReference()+"+"+
                    sheet.getRow(rowstart.getRowNum()).getCell(cdsl*2+4).getReference()+",\"合格\",\"不合格\"),\"-\")");
            //评定结果=IF(I41>=I39+I35,"合格","不合格")
        }else{
            sheet.getRow(rowstart.getRowNum()+9).getCell(cdsl*2+4).setCellValue("-");
            //评定结果=IF(I41>=I39+I35,"合格","不合格")
        }

        //合格率：
        sheet.getRow(rowstart.getRowNum()+8).getCell(cdsl*2+4).setCellFormula("IFERROR("+
                sheet.getRow(rowstart.getRowNum()+3).getCell(cdsl*2+4).getReference()+"/"
                +sheet.getRow(rowstart.getRowNum()+2).getCell(cdsl*2+4).getReference()+"*100,\"-\")");//I43=I38/I37*100

        try{
            double value = e.evaluate(sheet.getRow(rowstart.getRowNum()+3).getCell(cdsl*2+4)).getNumberValue();
            sheet.getRow(rowstart.getRowNum()+3).getCell(cdsl*2+4).setCellFormula(null);
            sheet.getRow(rowstart.getRowNum()+3).getCell(cdsl*2+4).setCellValue(value);
        }catch(Exception e1){
            sheet.getRow(rowstart.getRowNum()+3).getCell(cdsl*2+4).setCellFormula(null);
            sheet.getRow(rowstart.getRowNum()+3).getCell(cdsl*2+4).setCellValue("");
        }

        sheet.getRow(rowstart.getRowNum()).createCell(cdsl*3+7).setCellFormula(sheet.getRow(rowstart.getRowNum()+2).getCell(cdsl*2+4).getReference());
        double value = e.evaluate(sheet.getRow(rowstart.getRowNum()).getCell(cdsl*3+7)).getNumberValue();
        sheet.getRow(rowstart.getRowNum()).createCell(cdsl*3+7).setCellFormula(null);
        sheet.getRow(rowstart.getRowNum()).createCell(cdsl*3+7).setCellValue(value);

        sheet.getRow(rowstart.getRowNum()).createCell(cdsl*3+8).setCellFormula(sheet.getRow(rowstart.getRowNum()+3).getCell(cdsl*2+4).getReference());
        value = e.evaluate(sheet.getRow(rowstart.getRowNum()).getCell(cdsl*3+8)).getNumberValue();
        sheet.getRow(rowstart.getRowNum()).createCell(cdsl*3+8).setCellFormula(null);
        sheet.getRow(rowstart.getRowNum()).createCell(cdsl*3+8).setCellValue(value);

        sheet.getRow(rowstart.getRowNum()).createCell(cdsl*3+9).setCellFormula(
                sheet.getRow(rowstart.getRowNum()).getCell(cdsl*3+8).getReference()+"*100/"+
                        sheet.getRow(rowstart.getRowNum()).getCell(cdsl*3+7).getReference());
    }


    /**
     *左右幅路面
     * @param sheet
     */
    private void calculatePavementSheet(XSSFWorkbook wb,XSSFSheet sheet,int cdsl){
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        boolean flag = false;
        FormulaEvaluator e = new XSSFFormulaEvaluator(wb);
        int a = 0;
        int b = 0;
        if (cdsl == 2 || cdsl == 3){
            a= 39;
            b=39;

        }else if (cdsl == 4){
            a=23;
            b=23;

        }else if (cdsl == 5){
            a= 27;
            b=27;
        }
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            /**
             * 2c 38
             * 3c 38
             * 4c 22
             * 5c 26
             */
            // 下一张表
            if (!"".equals(row.getCell(0).toString()) && row.getCell(0).toString().contains("质量鉴定表") && flag) {
                i+=4;
            }
            if(flag){
                rowstart = sheet.getRow(i);
                rowend = sheet.getRow(i+a);
                fillPavementTotal(wb,sheet, rowstart, rowend,cdsl);
                i += a;
            }
            if ("桩号".equals(row.getCell(0).toString())) {
                i+=1;
                rowstart = sheet.getRow(i+1);
                rowend = sheet.getRow(i+b);
                fillPavementTotal(wb,sheet, rowstart, rowend,cdsl);
                i+=b;
                flag = true;
            }
        }
        sheet.getRow(4).createCell(cdsl*3+4).setCellFormula("MAX("
                +sheet.getRow(6).createCell(cdsl*3+4).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(cdsl*3+4).getReference()+")");

        sheet.getRow(4).createCell(cdsl*3+5).setCellFormula("MIN("
                +sheet.getRow(6).createCell(cdsl*3+5).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(cdsl*3+5).getReference()+")");


        sheet.getRow(4).createCell(cdsl*3+6).setCellFormula("SUM("
                +sheet.getRow(6).createCell(cdsl*3+6).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(cdsl*3+6).getReference()+")");

        sheet.getRow(4).createCell(cdsl*3+7).setCellFormula("SUM("
                +sheet.getRow(6).createCell(cdsl*3+7).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(cdsl*3+7).getReference()+")");
        /** cdsl*3+8
         * 2c 13
         * 3c 16
         * 4c 19
         * 5c 22
         */
        /*sheet.getRow(4).createCell(cdsl*3+7).setCellFormula("SUM("
                +sheet.getRow(6).createCell(cdsl*3+7).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).getCell(cdsl*3+7).getReference()+")");
        double value = e.evaluate(sheet.getRow(4).getCell(cdsl*3+7)).getNumberValue();
        sheet.getRow(4).getCell(cdsl*3+7).setCellFormula(null);
        sheet.getRow(4).getCell(cdsl*3+7).setCellValue(value);

        sheet.getRow(4).createCell(cdsl*3+8).setCellFormula("SUM("
                +sheet.getRow(6).createCell(cdsl*3+8).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).getCell(cdsl*3+8).getReference()+")");
        value = e.evaluate(sheet.getRow(4).getCell(cdsl*3+8)).getNumberValue();
        sheet.getRow(4).getCell(cdsl*3+8).setCellFormula(null);
        sheet.getRow(4).getCell(cdsl*3+8).setCellValue(value);*/
        /**cdsl*3+5
         * 2c 10
         * 3c 13
         * 4c 16
         * 5c 19
         */
        //最大值，最小值
        /*sheet.getRow(4).createCell(cdsl*3+4).setCellFormula("MAX("
                +sheet.getRow(6).createCell(cdsl*3+4).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(cdsl*3+4).getReference()+")");
        value = e.evaluate(sheet.getRow(4).getCell(cdsl*3+4)).getNumberValue();
        sheet.getRow(4).getCell(cdsl*3+4).setCellFormula(null);
        sheet.getRow(4).getCell(cdsl*3+4).setCellValue(value);

        sheet.getRow(4).createCell(cdsl*3+5).setCellFormula("MIN("
                +sheet.getRow(6).createCell(cdsl*3+5).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(cdsl*3+5).getReference()+")");
        value = e.evaluate(sheet.getRow(4).getCell(cdsl*3+5)).getNumberValue();
        sheet.getRow(4).getCell(cdsl*3+5).setCellFormula(null);
        sheet.getRow(4).getCell(cdsl*3+5).setCellValue(value);*/
    }

    /**
     *
     * @param sheet
     * @param rowstart
     * @param rowend
     */
    public void fillPavementTotal(XSSFWorkbook wb,XSSFSheet sheet, XSSFRow rowstart, XSSFRow rowend,int cdsl){
        /**cdsl*2+5
         * 2c 8
         * 3c 10
         * 4c 12
         * 5c 14
         */
        //代表值允许偏差（cm）
        System.out.println(sheet.getRow(rowend.getRowNum()-9).getCell(cdsl*2+4));
        sheet.getRow(rowend.getRowNum()-9).getCell(cdsl*2+4).setCellFormula("-0.05*"+ sheet.getRow(rowend.getRowNum()-5).getCell(cdsl*2+4).getReference());//代表值允许偏差（cm）=-0.05*I39
        //合格值允许偏差（cm）
        sheet.getRow(rowend.getRowNum()-8).getCell(cdsl*2+4).setCellFormula("-0.1*"+
                sheet.getRow(rowend.getRowNum()-5).getCell(cdsl*2+4).getReference());//合格值允许偏差（cm）=-0.1*I39

        //总点数
        sheet.getRow(rowend.getRowNum()-7).getCell(cdsl*2+4).setCellFormula("COUNT("
                +rowstart.getCell(1).getReference()+":"
                +rowend.getCell(cdsl).getReference()+","

                /**
                 * 2c 5
                 * 3c 7
                 * 4c 9
                 */
                +rowstart.getCell(cdsl+2).getReference()+":"
                +rowend.getCell(cdsl*2+1).getReference()+","

                /**
                 * 2c 7
                 * 3c 9
                 * 4c 11
                 */
                +rowstart.getCell(cdsl*2+3).getReference()+":"
                +sheet.getRow(rowend.getRowNum()-10).getCell(cdsl*2+5).getReference()+")");//I37=COUNT(B6:C44,E6:F44,H6:I27)

        XSSFFormulaEvaluator e=new XSSFFormulaEvaluator(wb);

        Cell cell = sheet.getRow(rowend.getRowNum()-5).getCell(cdsl*2+4);
        double data = 0;
        if (cell.getCellTypeEnum() == CellType.NUMERIC) {
            data = cell.getNumericCellValue() + e.evaluate(sheet.getRow(rowend.getRowNum()-8).getCell(cdsl*2+4)).getNumberValue();
        }
        //data = sheet.getRow(rowend.getRowNum()-5).getCell(cdsl*2+4).getNumericCellValue()+ e.evaluate(sheet.getRow(rowend.getRowNum()-8).getCell(cdsl*2+4)).getNumberValue();

        //合格数
        sheet.getRow(rowend.getRowNum()-6).getCell(cdsl*2+4).setCellFormula(
                sheet.getRow(rowend.getRowNum()-7).getCell(cdsl*2+4).getReference()+"-SUM(COUNTIF("

                        +rowstart.getCell(cdsl*2+3).getReference()+":"
                        +sheet.getRow(rowend.getRowNum()-10).getCell(cdsl*2+5).getReference()+",\"<"+data+"\"),COUNTIF("

                        +rowstart.getCell(cdsl+2).getReference()+":"
                        +rowend.getCell(cdsl*2+1).getReference()+",\"<"+data+"\"),COUNTIF("

                        +rowstart.getCell(1).getReference()+":"
                        +rowend.getCell(cdsl).getReference()+",\"<"+data+"\"))");//I38=I37-SUM(COUNTIF(H6:I27,"<16.2"),COUNTIF(E6:F44,"<16.2"),COUNTIF(B6:C44,"<16.2"))
        try{
            double value = e.evaluate(sheet.getRow(rowend.getRowNum()-6).getCell(cdsl*2+4)).getNumberValue();
            sheet.getRow(rowend.getRowNum()-6).getCell(cdsl*2+4).setCellFormula(null);
            sheet.getRow(rowend.getRowNum()-6).getCell(cdsl*2+4).setCellValue(value);
        }catch(Exception e1){
            sheet.getRow(rowend.getRowNum()-6).getCell(cdsl*2+4).setCellFormula(null);
            sheet.getRow(rowend.getRowNum()-6).getCell(cdsl*2+4).setCellValue("");
        }
        //平均值(cm)
        sheet.getRow(rowend.getRowNum()-4).getCell(cdsl*2+4).setCellFormula("IFERROR(AVERAGE("
                +rowstart.getCell(1).getReference()+":"
                +rowend.getCell(cdsl).getReference()+","
                +rowstart.getCell(cdsl+2).getReference()+":"
                +rowend.getCell(cdsl*2+1).getReference()+","
                +rowstart.getCell(cdsl*2+3).getReference()+":"
                +sheet.getRow(rowend.getRowNum()-10).getCell(cdsl*2+5).getReference()+"),\"-\")");//I40=AVERAGE(B6:C44,E6:F44,H6:I27)

        //均方差
        sheet.getRow(rowend.getRowNum()-2).getCell(cdsl*2+4).setCellFormula("IFERROR(STDEV("
                +rowstart.getCell(1).getReference()+":"
                +rowend.getCell(cdsl).getReference()+","
                +rowstart.getCell(cdsl+2).getReference()+":"
                +rowend.getCell(cdsl*2+1).getReference()+","
                +rowstart.getCell(cdsl*2+3).getReference()+":"
                +sheet.getRow(rowend.getRowNum()-10).getCell(cdsl*2+5).getReference()+"),\"-\")");//I42=STDEV(B6:C44,E6:F44,H6:I27)

        //代表值(cm)
        sheet.getRow(rowend.getRowNum()-3).getCell(cdsl*2+4).setCellFormula("IFERROR(IF("
                +sheet.getRow(rowend.getRowNum()-7).getCell(cdsl*2+4).getReference()+">100,("
                +sheet.getRow(rowend.getRowNum()-4).getCell(cdsl*2+4).getReference()+"-"
                +sheet.getRow(rowend.getRowNum()-2).getCell(cdsl*2+4).getReference()+"*1.6449/SQRT("
                +sheet.getRow(rowend.getRowNum()-7).getCell(cdsl*2+4).getReference()+")),("
                +sheet.getRow(rowend.getRowNum()-4).getCell(cdsl*2+4).getReference()+"-VLOOKUP("
                +sheet.getRow(rowend.getRowNum()-7).getCell(cdsl*2+4).getReference()+",Sheet1!A:D,"+(3)+",FALSE)*"
                +sheet.getRow(rowend.getRowNum()-2).getCell(cdsl*2+4).getReference()+")),\"-\")");//I41=IF(I37>100,(I40-I42*1.6449/SQRT(I37)),(I40-VLOOKUP(I37,Sheet1!A:D,3,FALSE)*I42))

        //合格率
        sheet.getRow(rowend.getRowNum()-1).getCell(cdsl*2+4).setCellFormula("IFERROR("+
                sheet.getRow(rowend.getRowNum()-6).getCell(cdsl*2+4).getReference()+"/"
                +sheet.getRow(rowend.getRowNum()-7).getCell(cdsl*2+4).getReference()+"*100,\"-\")");//I43=I38/I37*100

        //评定结果
        double totalNum = e.evaluate(sheet.getRow(rowend.getRowNum()-7).getCell(cdsl*2+4)).getNumberValue();//总点数不为0的时候才计算最大值最小值
        if(totalNum != 0){
            sheet.getRow(rowend.getRowNum()).getCell(cdsl*2+4).setCellFormula("IFERROR(IF("+
                    sheet.getRow(rowend.getRowNum()-3).getCell(cdsl*2+4).getReference()+">="+
                    sheet.getRow(rowend.getRowNum()-5).getCell(cdsl*2+4).getReference()+"+"+
                    sheet.getRow(rowend.getRowNum()-9).getCell(cdsl*2+4).getReference()+",\"合格\",\"不合格\"),\"-\")");
            //评定结果=IF(I41>=I39+I35,"合格","不合格")
        }else{
            sheet.getRow(rowend.getRowNum()).getCell(cdsl*2+4).setCellValue("-");
            //评定结果=IF(I41>=I39+I35,"合格","不合格")
        }

        //MAX
        //=IF(K37<>0,MAX(B6:D44,F6:H44,J6:L34),"")

        /**
         * cdsl*3+4
         * 2c 10
         * 3c 13
         * 4c 16
         * 5c 19
         */
        sheet.getRow(rowend.getRowNum()-7).createCell(cdsl*3+4).setCellFormula("IF("
                +sheet.getRow(rowend.getRowNum()-7).getCell(cdsl*2+4).getReference()+"<>0,MAX("

                +rowstart.getCell(1).getReference()+":"
                +rowend.getCell(cdsl).getReference()+","

                +rowstart.getCell(cdsl+2).getReference()+":"
                +rowend.getCell(cdsl*2+1).getReference()+","

                +rowstart.getCell(cdsl*2+3).getReference()+":"
                +sheet.getRow(rowend.getRowNum()-10).getCell(cdsl*2+5).getReference()+"),\"\")");//I37=COUNT(B6:C44,E6:F44,H6:I27)


        //=IF(K38<>0,MIN(B6:D44,F6:H44,J6:L34),"")
        sheet.getRow(rowend.getRowNum()-7).createCell(cdsl*3+5).setCellFormula("IF("
                +sheet.getRow(rowend.getRowNum()-6).getCell(cdsl*2+4).getReference()+"<>0,MIN("

                +rowstart.getCell(1).getReference()+":"
                +rowend.getCell(cdsl).getReference()+","

                +rowstart.getCell(cdsl+2).getReference()+":"
                +rowend.getCell(cdsl*2+1).getReference()+","

                +rowstart.getCell(cdsl*2+3).getReference()+":"
                +sheet.getRow(rowend.getRowNum()-10).getCell(cdsl*2+5).getReference()+"),\"\")");//I37=COUNT(B6:C44,E6:F44,H6:I27)

        sheet.getRow(rowend.getRowNum()-7).createCell(cdsl*3+6).setCellFormula(sheet.getRow(rowend.getRowNum()-7).getCell(cdsl*2+4).getReference());
        sheet.getRow(rowend.getRowNum()-7).createCell(cdsl*3+7).setCellFormula(sheet.getRow(rowend.getRowNum()-6).getCell(cdsl*2+4).getReference());

        /*rowend.createCell(cdsl*3+8).setCellFormula(sheet.getRow(rowend.getRowNum()-7).getCell(cdsl*2+4).getReference());

        double value = e.evaluate(rowend.getCell(cdsl*3+8)).getNumberValue();
        rowend.getCell(cdsl*3+8).setCellFormula(null);
        rowend.getCell(cdsl*3+8).setCellValue(value);

        rowend.createCell(cdsl*3+9).setCellFormula(sheet.getRow(rowend.getRowNum()-6).getCell(cdsl*2+4).getReference());
        value = e.evaluate(rowend.getCell(cdsl*3+9)).getNumberValue();
        rowend.getCell(cdsl*3+9).setCellFormula(null);
        rowend.getCell(cdsl*3+9).setCellValue(value);
        *//*
         * 最大值，最小值
         * 原先统计的是实测值的最大最小值，现在是代表值的最大最小值
         *//*
        //代表值的最大最小值
        sheet.getRow(rowend.getRowNum()-7).createCell(cdsl*3+4).setCellFormula(sheet.getRow(rowend.getRowNum()-3).getCell(cdsl*2+4).getReference());
        sheet.getRow(rowend.getRowNum()-7).createCell(cdsl*3+5).setCellFormula(sheet.getRow(rowend.getRowNum()-3).getCell(cdsl*2+4).getReference());*/
    }


    /**
     *
     * @param sheet
     * @return
     */
    private boolean shouldBeCalculate(XSSFSheet sheet) {
        sheet.getRow(5).getCell(0).setCellType(CellType.STRING);
        if(sheet.getRow(5).getCell(0)==null ||"".equals(sheet.getRow(5).getCell(0).getStringCellValue())){
            return false;
        }

        return true;
    }



    /**
     *
     * @param proname
     * @param htd
     * @param data
     * @param sdqllmData
     * @param wb
     * @param sheetname
     * @param cdsl
     * @param sjz
     * @param zx
     */
    private void DBtoExcelZd(String proname, String htd, List<Map<String, Object>> data, List<Map<String, Object>> sdqllmData, XSSFWorkbook wb, String sheetname, int cdsl, String sjz, String zx) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        SimpleDateFormat outputDateFormat  = new SimpleDateFormat("yyyy.MM.dd");
        if (data!=null && !data.isEmpty()) {
            int b = 0;
            if (cdsl == 2){
                b=42;
            }else if (cdsl == 3){
                b=42;
            }else if (cdsl == 4){
                b=23;
            }else if (cdsl == 5){
                b=27;
            }
            List<Map<String, Object>> lmdata = handlezdData(proname,data,sdqllmData,zx);
            createTable(getNumZd(lmdata,cdsl)+1, wb, sheetname, cdsl);
            XSSFSheet sheet = wb.getSheet(sheetname);

            String name = data.get(0).get("name").toString();
            int index = 5;
            int tableNum = 0;
            String time1 = String.valueOf(data.get(0).get("createTime")) ;
            Date parse = simpleDateFormat.parse(time1);
            String time = outputDateFormat.format(parse);

            fillTitleCellData(sheet, tableNum, proname, htd, name,time,sheetname,cdsl,sjz);

            String zdbs = lmdata.get(0).get("zdbs").toString();
            List<Map<String, Object>> rowAndcol = new ArrayList<>();
            int startRow = -1, endRow = -1, startCol = -1, endCol = -1;
            for (Map<String, Object> lm : lmdata) {
                System.out.println(zdbs);
                if (lm.get("zdbs").toString().equals(zdbs)) {
                    int vv = tableNum * b + index % b;
                    System.out.println("IF  "+zdbs+" "+tableNum+"  "+index+"  "+lm.get("zh")+"  "+vv);
                    //if(index % (b-1) == 0){  D 22-23
                    if(index > (b-1)){
                        tableNum ++;
                        index = 0;
                    }
                    if (!lm.get("ld").toString().equals("") && !lm.get("ld").toString().isEmpty()) {
                        String[] sfc = lm.get("ld").toString().split(",");
                        for (int i = 0; i < sfc.length; i++) {
                            sheet.getRow(tableNum * b + index % b).getCell(0).setCellValue(lm.get("zdbs").toString()+"匝道");
                            sheet.getRow(tableNum * b + index % b).getCell(1).setCellValue((Double.parseDouble(lm.get("zh").toString())));
                            if (sfc[i].equals("-")){
                                sheet.getRow(tableNum * b + index % b).getCell(2 + i).setCellValue(sfc[i]);
                            }else {
                                sheet.getRow(tableNum * b + index % b).getCell(2+ i).setCellValue(Double.parseDouble(sfc[i]));
                            }

                        }
                    }else {
                        for (int i = 0; i < cdsl; i++) {
                            sheet.getRow(tableNum * b + index % b).getCell(0).setCellValue(lm.get("zdbs").toString()+"匝道");
                            sheet.getRow(tableNum * b + index % b).getCell(1).setCellValue((Double.parseDouble(lm.get("zh").toString())));
                            sheet.getRow(tableNum * b + index % b).getCell(2 + i).setCellValue(lm.get("name").toString());

                            startRow = tableNum * b + index % b ;
                            endRow = tableNum * b  + index % b ;

                            startCol = 2;
                            endCol = 1+cdsl;

                        }
                        //可以在这块记录一个行和列
                        Map<String, Object> map = new HashMap<>();
                        map.put("startRow",startRow);
                        map.put("endRow",endRow);
                        map.put("startCol",startCol);
                        map.put("endCol",endCol);
                        map.put("name",lm.get("name"));
                        map.put("tableNum",tableNum);
                        rowAndcol.add(map);

                    }
                    index ++;

                }else {
                    zdbs = lm.get("zdbs").toString();
                    System.out.println("else  "+zdbs+" "+tableNum);
                    tableNum ++;
                    index = 5;
                    if(index > (b-1)){
                        tableNum ++;
                        index = 0;
                    }
                    if (!lm.get("ld").toString().equals("") && !lm.get("ld").toString().isEmpty()) {
                        String[] sfc = lm.get("ld").toString().split(",");
                        for (int i = 0; i < sfc.length; i++) {
                            sheet.getRow(tableNum * b + index % b).getCell(0).setCellValue(lm.get("zdbs").toString()+"匝道");
                            sheet.getRow(tableNum * b + index % b).getCell(1).setCellValue((Double.parseDouble(lm.get("zh").toString())));
                            if (sfc[i].equals("-")){
                                sheet.getRow(tableNum * b + index % b).getCell(2 + i).setCellValue(sfc[i]);
                            }else {
                                sheet.getRow(tableNum * b + index % b).getCell(2+ i).setCellValue(Double.parseDouble(sfc[i]));
                            }

                        }
                    }else {
                        for (int i = 0; i < cdsl; i++) {
                            sheet.getRow(tableNum * b + index % b).getCell(0).setCellValue(lm.get("zdbs").toString()+"匝道");
                            sheet.getRow(tableNum * b + index % b).getCell(1).setCellValue((Double.parseDouble(lm.get("zh").toString())));
                            sheet.getRow(tableNum * b + index % b).getCell(2 + i).setCellValue(lm.get("name").toString());

                            startRow = tableNum * b + index % b ;
                            endRow = tableNum * b  + index % b ;

                            startCol = 2;
                            endCol = 1+cdsl;

                        }
                        //可以在这块记录一个行和列
                        Map<String, Object> map = new HashMap<>();
                        map.put("startRow",startRow);
                        map.put("endRow",endRow);
                        map.put("startCol",startCol);
                        map.put("endCol",endCol);
                        map.put("name",lm.get("name"));
                        map.put("tableNum",tableNum);
                        rowAndcol.add(map);

                    }
                    index ++;

                }
            }

            List<Map<String, Object>> maps = mergeCells(rowAndcol);
            for (Map<String, Object> map : maps) {
                sheet.addMergedRegion(new CellRangeAddress(Integer.parseInt(map.get("startRow").toString()), Integer.parseInt(map.get("endRow").toString()), Integer.parseInt(map.get("startCol").toString()), Integer.parseInt(map.get("endCol").toString())));
            }
        }
    }



    /**
     *
     * @param proname
     * @param htd
     * @param data
     * @param sdqllmData
     * @param wb
     * @param sheetname
     * @param cdsl
     * @param sjz
     * @param zx
     */
    private void DBtoExcelLM(String proname, String htd, List<Map<String, Object>> data, List<Map<String, Object>> sdqllmData, XSSFWorkbook wb, String sheetname, int cdsl, String sjz, String zx) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        SimpleDateFormat outputDateFormat  = new SimpleDateFormat("yyyy.MM.dd");
        if (data!=null && !data.isEmpty()) {
            int a = 0;
            int b = 0;
            int c = 0;
            int d =0;
            if (cdsl == 2 || cdsl == 3){
                a = 106;
                b = 78;
                c = 44;
                d = 39;
            }else if (cdsl == 4){
                a = 58;
                b = 46;
                c = 28;
                d = 23;
            }else if (cdsl == 5){
                a = 70;
                b = 54;
                c = 32;
                d = 27;
            }
            createTableLM(getNumLM(data,cdsl), wb, sheetname, cdsl);
            XSSFSheet sheet = wb.getSheet(sheetname);

            String name = data.get(0).get("name").toString();
            int index = 0;
            int tableNum = 0;
            String time1 = String.valueOf(data.get(0).get("createTime")) ;
            Date parse = simpleDateFormat.parse(time1);
            String time = outputDateFormat.format(parse);

            fillTitleCellData(sheet, tableNum, proname, htd, name,time,sheetname,cdsl,sjz);

            List<Map<String, Object>> lmdata = handleLmData(data,sdqllmData);

            List<Map<String, Object>> rowAndcol = new ArrayList<>();
            int startRow = -1, endRow = -1, startCol = -1, endCol = -1;
            for (Map<String, Object> lm : lmdata) {
                if (index > a) {
                    tableNum++;
                    fillTitleCellData(sheet, tableNum, proname, htd, name, time, sheetname,cdsl,sjz);
                    index = 0;
                }
                if (!lm.get("ld").toString().equals("") && !lm.get("ld").toString().isEmpty()) {
                    String[] sfc = lm.get("ld").toString().split(",");
                    for (int i = 0; i < sfc.length; i++) {
                        if (index < b) {
                            sheet.getRow(tableNum * c + 5 + index % d).getCell((cdsl+1) * (index / d)).setCellValue((Double.parseDouble(lm.get("zh").toString())));
                            if (sfc[i].equals("-")){
                                sheet.getRow(tableNum * c + 5 + index % d).getCell((cdsl+1) * (index / d) + 1 + i).setCellValue(sfc[i]);
                            }else {
                                sheet.getRow(tableNum * c + 5 + index % d).getCell((cdsl+1) * (index / d) + 1 + i).setCellValue(Double.parseDouble(sfc[i]));
                            }

                        } else {
                            sheet.getRow(tableNum * c + 5 + index % b).getCell((cdsl*2+2)  * (index / b)).setCellValue((Double.parseDouble(lm.get("zh").toString())));
                            if (sfc[i].equals("-")){
                                sheet.getRow(tableNum * c + 5 + index % b).getCell((cdsl*2+2)  * (index / b) + 1 + i).setCellValue(sfc[i]);
                            }else {
                                sheet.getRow(tableNum * c + 5 + index % b).getCell((cdsl*2+2)  * (index / b) + 1 + i).setCellValue(Double.parseDouble(sfc[i]));
                            }

                        }

                    }

                } else {
                    for (int i = 0; i < cdsl; i++) {
                        if (index < b) {
                            sheet.getRow(tableNum * c + 5 + index % d).getCell((cdsl+1) * (index / d)).setCellValue((Double.parseDouble(lm.get("zh").toString())));
                            sheet.getRow(tableNum * c + 5 + index % d).getCell((cdsl+1) * (index / d) + 1 + i).setCellValue(lm.get("name").toString());

                            startRow = tableNum * c + 5 + index % d ;
                            endRow = tableNum * c + 5 + index % d ;

                            startCol = (cdsl+1)  * (index / d) + 1;
                            endCol = (cdsl+1)  * (index / d) + cdsl;

                        } else {
                            sheet.getRow(tableNum * c + 5 + index % b).getCell((cdsl*2+2)  * (index / b)).setCellValue((Double.parseDouble(lm.get("zh").toString())));
                            sheet.getRow(tableNum * c + 5 + index % b).getCell((cdsl*2+2)  * (index / b) + 1 + i).setCellValue(lm.get("name").toString());

                            startRow = tableNum * c + 5 + index % b ;
                            endRow = tableNum * c + 5 + index % b ;

                            startCol = 2*cdsl+3;
                            endCol = 3*cdsl+2;

                        }
                    }
                    //可以在这块记录一个行和列
                    Map<String, Object> map = new HashMap<>();
                    map.put("startRow",startRow);
                    map.put("endRow",endRow);
                    map.put("startCol",startCol);
                    map.put("endCol",endCol);
                    map.put("name",lm.get("name"));
                    map.put("tableNum",tableNum);
                    rowAndcol.add(map);

                }
                index++;

            }

            List<Map<String, Object>> maps = mergeCells(rowAndcol);
            for (Map<String, Object> map : maps) {
                sheet.addMergedRegion(new CellRangeAddress(Integer.parseInt(map.get("startRow").toString()), Integer.parseInt(map.get("endRow").toString()), Integer.parseInt(map.get("startCol").toString()), Integer.parseInt(map.get("endCol").toString())));
            }
        }
    }

    /**
     *
     * @param data
     * @param sdqllmData
     * @return
     */
    private List<Map<String, Object>> handleLmData(List<Map<String, Object>> data, List<Map<String, Object>> sdqllmData) {
        for (Map<String, Object> datum : data) {
            for (Map<String, Object> zfsdqlDatum : sdqllmData) {
                if (datum.get("zh").toString().equals(zfsdqlDatum.get("zh")) && datum.get("cd").toString().equals(zfsdqlDatum.get("cd"))){
                    datum.put("ld","");
                    datum.put("name",zfsdqlDatum.get("name"));
                }
            }
        }
        Collections.sort(data, new Comparator<Map<String, Object>>() {
            @Override
            public int compare(Map<String, Object> o1, Map<String, Object> o2) {
                // 名字相同时按照 qdzh 排序
                Double qdzh1 = Double.parseDouble(o1.get("zh").toString());
                Double qdzh2 = Double.parseDouble(o2.get("zh").toString());
                return qdzh1.compareTo(qdzh2);
            }
        });
        return data;
    }

    /**
     *
     * @param data
     * @param zfsdqlData
     * @return
     */
    private List<Map<String, Object>> handlezdData(String proname,List<Map<String, Object>> data, List<Map<String, Object>> zfsdqlData,String zx) {
        /**
         * data是全部的数据，zfsdqlData是包含桥梁和隧道的数据，还有可能为空
         */
        String name = data.get(0).get("name").toString();
        String lf = "";
        if (name.contains("左幅")) {
            lf = "左幅";
        } else if (name.contains("右幅")) {
            lf = "右幅";
        }
        QueryWrapper<JjgSfz> wrapper = new QueryWrapper<>();
        wrapper.like("proname", proname);
        wrapper.like("sshtmc", zx);
        wrapper.like("lf", lf);
        List<JjgSfz> jjgSfzs = jjgLqsSfzMapper.selectList(wrapper);//(id=1, zdsfzname=淮宁湾收费站, htd=土建2标, lf=左幅, zhq=930.0, zhz=1250.0, pzlx=水泥混凝土, sszd=E, sshtmc=淮宁湾立交, proname=陕西高速, createTime=Tue Jun 13 21:03:29 CST 2023)
        List<Map<String, Object>> sfzlist = new ArrayList<>();
        for (int i = 0; i < jjgSfzs.size(); i++) {
            double zhq = Double.parseDouble(jjgSfzs.get(i).getZhq());
            double zhz = Double.parseDouble(jjgSfzs.get(i).getZhz());
            String zdsfzname = jjgSfzs.get(i).getZdsfzname();
            String sszd = jjgSfzs.get(i).getSszd();
            sfzlist.addAll(incrementByTen(zhq, zhz, zdsfzname, sszd));
        }

        if (zfsdqlData.size() > 0) {
            for (Map<String, Object> datum : data) {
                for (Map<String, Object> zfsdqlDatum : zfsdqlData) {
                    if (datum.get("zdbs").toString().equals(zfsdqlDatum.get("zdbs")) && datum.get("zh").toString().equals(zfsdqlDatum.get("zh"))) {
                        datum.put("ld", "");
                        datum.put("name", zfsdqlDatum.get("name"));
                    }
                }
            }
        }
        if (sfzlist.size() > 0) {
            data.addAll(sfzlist);
        }
        Collections.sort(data, new Comparator<Map<String, Object>>() {
            @Override
            public int compare(Map<String, Object> o1, Map<String, Object> o2) {
                String name1 = o1.get("zdbs").toString();
                String name2 = o2.get("zdbs").toString();
                // 按照名字进行排序
                int cmp = name1.compareTo(name2);
                if (cmp != 0) {
                    return cmp;
                }
                // 名字相同时按照 qdzh 排序
                Double qdzh1 = Double.parseDouble(o1.get("zh").toString());
                Double qdzh2 = Double.parseDouble(o2.get("zh").toString());
                return qdzh1.compareTo(qdzh2);
            }
        });
        return data;

    }

    /**
     * 收费站桩号加10
     * @param start
     * @param end
     * @param zdsfzname
     * @param sszd
     * @return
     */
    private List<Map<String,Object>> incrementByTen(double start, double end,String zdsfzname,String sszd) {
        List<Map<String,Object>> result = new ArrayList<>();
        if (start <= end) {
            int a = (int) start/10;
            double s = (a*10)+10;
            for (double i = s; i <= end; i += 10) {
                Map map = new HashMap();
                map.put("zh",i);
                map.put("name",zdsfzname);
                map.put("zdbs",sszd);
                map.put("ld","");
                result.add(map);
            }
            return result;
        } else {
            return new ArrayList();

        }
    }

    /**
     * 合并单元格
     * @param rowAndcol
     * @return
     */
    private List<Map<String, Object>> mergeCells(List<Map<String, Object>> rowAndcol) {
        List<Map<String, Object>> result = new ArrayList<>();
        int currentEndRow = -1;
        int currentStartRow = -1;
        int currentStartCol = -1;
        int currentEndCol = -1;
        String currentName = null;
        int currentTableNum = -1;
        for (Map<String, Object> row : rowAndcol) {
            int tableNum = (int) row.get("tableNum");
            int startRow = (int) row.get("startRow");
            int endRow = (int) row.get("endRow");
            int startCol = (int) row.get("startCol");
            int endCol = (int) row.get("endCol");
            String name = (String) row.get("name");
            if (currentName == null || !currentName.equals(name) || currentStartCol != startCol || currentEndCol != endCol || currentTableNum != tableNum) {
                if (currentStartRow != -1) {
                    for (int i = currentStartRow; i <= currentEndRow && i < result.size(); i++) {
                        Map<String, Object> newRow = new HashMap<>();
                        newRow.put("name", currentName);
                        newRow.put("startRow", currentStartRow);
                        newRow.put("endRow", currentEndRow);
                        newRow.put("startCol", currentStartCol);
                        newRow.put("endCol", currentEndCol);
                        newRow.put("tableNum", currentTableNum);
                        newRow.putAll(result.get(i));
                        result.set(i, newRow);
                    }
                }
                currentName = name;
                currentStartCol = startCol;
                currentEndCol = endCol;
                currentTableNum = tableNum;
                currentStartRow = startRow;
                currentEndRow = endRow;
                result.add(row);
            } else {
                Map<String, Object> lastRow = result.get(result.size() - 1);
                lastRow.put("endRow", endRow);
                currentEndRow = endRow;
            }
        }
        if (currentStartRow != -1) {
            for (int i = currentStartRow; i <= currentEndRow && i < result.size(); i++) {
                Map<String, Object> newRow = new HashMap<>();
                newRow.put("name", currentName);
                newRow.put("startRow", currentStartRow);
                newRow.put("endRow", currentEndRow);
                newRow.put("startCol", currentStartCol);
                newRow.put("endCol", currentEndCol);
                newRow.put("tableNum", currentTableNum);
                newRow.putAll(result.get(i));
                result.set(i, newRow);
            }
        }
        return result;
    }

    /**
     *
     * @param data
     * @param cdsl
     * @return
     */
    private int getNumLM(List<Map<String, Object>> data, int cdsl) {
        int a = 0;
        if (cdsl == 2 || cdsl == 3){
            a = 107;
        }else if (cdsl == 4){
            a = 59;
        }else if (cdsl == 5){
            a = 71;
        }
        Map<String, Integer> resultMap = new HashMap<>();
        for (Map<String, Object> map : data) {
            String name = map.get("name").toString();
            if (resultMap.containsKey(name)) {
                resultMap.put(name, resultMap.get(name) + 1);
            } else {
                resultMap.put(name, 1);
            }
        }
        int num = 0;
        for (Map.Entry<String, Integer> entry : resultMap.entrySet()) {
            int value = entry.getValue();
            if (value%a==0){
                num += value/a;
            }else {
                num += value/a+1;
            }
        }
        return num;
    }

    /**
     *
     * @param data
     * @param cdsl
     * @return
     */
    private int getNumZd(List<Map<String, Object>> data, int cdsl) {
        int m = 0;
        if (cdsl == 2){
            m = 42;
        }else if (cdsl == 3){
            m = 42;
        }else if (cdsl == 4){
            m = 23;
        }else if (cdsl == 5){
            m = 27;

        }

        Map<String, Integer> resultMap = new HashMap<>();
        for (Map<String, Object> map : data) {
            String name = map.get("zdbs").toString();
            if (resultMap.containsKey(name)) {
                resultMap.put(name, resultMap.get(name) + 1);
            } else {
                resultMap.put(name, 1);
            }
        }
        int num = 0;
        for (Map.Entry<String, Integer> entry : resultMap.entrySet()) {
            int value = entry.getValue();
            if (value%m==0){
                num += value/m;
            }else {
                num += value/m+1;
            }
        }
        return num;
    }


    /**
     *
     * @param proname
     * @param htd
     * @param data
     * @param wb
     * @param sheetname
     * @param cdsl
     * @param sjz
     * @param zx
     */
    private void DBtoExcel(String proname, String htd, List<Map<String, Object>> data, XSSFWorkbook wb, String sheetname, int cdsl, String sjz, String zx) throws ParseException {
        int b = 0;
        if (cdsl == 2 || cdsl == 3){
            b=42;
        }else if (cdsl == 4){
            b=23;
        }else if (cdsl == 5){
            b=27;
        }
        if (data!=null && !data.isEmpty()){
            createTable(getNum(data,cdsl),wb,sheetname,cdsl);
            writesdqlzyf(wb,data,proname, htd,sheetname,cdsl,sjz,b,zx);

        }
    }

    /**
     *
     * @param data
     * @param proname
     * @param htd
     * @param sheetname
     * @param cdsl
     * @param sjz
     * @param b
     * @param zx
     */
    private void writesdqlzyf(XSSFWorkbook wb, List<Map<String, Object>> data, String proname, String htd, String sheetname, int cdsl, String sjz,int b,String zx) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        SimpleDateFormat outputDateFormat  = new SimpleDateFormat("yyyy.MM.dd");
        XSSFSheet sheet = wb.getSheet(sheetname);
        String time = String.valueOf(data.get(0).get("createTime")) ;
        Date parse = simpleDateFormat.parse(time);
        String sj = outputDateFormat.format(parse);

        sheet.getRow(1).getCell(1).setCellValue(proname);
        sheet.getRow(2).getCell(cdsl*2+3).setCellValue(sj);
        sheet.getRow(1).getCell(cdsl*2+3).setCellValue(htd);
        sheet.getRow(2).getCell(1).setCellValue("路面面层");

        String name = data.get(0).get("name").toString();
        int index = 5;
        int tableNum = 0;

        for(int i =0; i < data.size(); i++){
            if (name.equals(data.get(i).get("name"))){
                if(index > (b-1)){
                    tableNum ++;
                    index = 0;
                }
                fillCommonCellData(sheet, tableNum, index, data.get(i),cdsl,zx);
                index ++;
            }else {
                name = data.get(i).get("name").toString();
                tableNum ++;//去掉是因为有空白页
                index = 5;
                fillCommonCellData(sheet, tableNum, index, data.get(i),cdsl,zx);
                index += 1;

            }
        }

    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param index
     * @param row
     * @param cdsl
     * @param zx
     */
    private void fillCommonCellData(XSSFSheet sheet, int tableNum, int index, Map<String, Object> row, int cdsl, String zx) {
        int b = 0;
        if (cdsl == 2){
            b=42;
        }else if (cdsl == 3){
            b=42;
        }else if (cdsl == 4){
            b=23;
        }else if (cdsl ==5){
            b=27;
        }
        sheet.getRow(tableNum * b + index % b).getCell(0).setCellValue(row.get("name").toString());
        sheet.getRow(tableNum * b + index % b).getCell(1).setCellValue(Double.valueOf(row.get("zh").toString()));

        String[] sfc = row.get("ld").toString().split(",");
        if (!sfc[0].isEmpty()) {
            for (int i = 0 ; i < sfc.length ; i++) {
                if (!sfc[i].equals("-")){
                    sheet.getRow(tableNum * b + index % b).getCell(2+i).setCellValue(Double.parseDouble(sfc[i]));
                }

            }
        }

    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param proname
     * @param htd
     * @param name
     * @param sj
     * @param sheetname
     * @param cdsl
     * @param sjz
     */
    private void fillTitleCellData(XSSFSheet sheet, int tableNum, String proname, String htd, String name, String sj, String sheetname, int cdsl, String sjz) {
        int a = 0;
        int dd = 0;
        if (sheetname.equals("左幅") || sheetname.equals("右幅")){
            if (cdsl == 2){
                dd = 38;
                a = 44;
            }else if (cdsl == 3){
                dd = 38;
                a = 44;
            }else if (cdsl == 4 ){
                dd = 22;
                a = 28;
            }else if(cdsl ==5){
                dd = 26;
                a = 32;
            }
        }else {
            if (cdsl == 2){
                a = 47;
            }else if (cdsl == 3){
                a = 47;
            }else if (cdsl == 4 ){
                a = 28;
            }else if(cdsl ==5){
                a = 32;
            }
        }
        sheet.getRow(tableNum * a + 1).getCell(1).setCellValue(proname);
        sheet.getRow(tableNum * a + 1).getCell(cdsl*2+3).setCellValue(htd);
        sheet.getRow(tableNum * a + 2).getCell(1).setCellValue("路面面层");
        sheet.getRow(tableNum * a + 2).getCell(cdsl*2+3).setCellValue(sj);
        if (sheetname.equals("左幅") || sheetname.equals("右幅")){
            sheet.getRow(tableNum * a + dd).getCell(cdsl*2+4).setCellValue(Double.parseDouble(sjz));
        }else {
            sheet.getRow(tableNum * a + 9).getCell(cdsl*2+4).setCellValue(Double.parseDouble(sjz));
        }
    }

    /**
     * 创建路面的页数
     * @param tableNum
     * @param wb
     * @param sheetname
     * @param cdsl
     */
    private void createTableLM(int tableNum, XSSFWorkbook wb, String sheetname, int cdsl) {
        int endrow = 0;
        int printrow = 0;
        int cdnum = 0;
        if (cdsl == 2 ) {
            endrow = 43;
            printrow = 44;
            cdnum = cdsl*2+5;
        }else if (cdsl == 3){
            endrow = 43;
            printrow = 44;
            cdnum = cdsl*2+6;
        } else if (cdsl == 4){
            endrow = 27;
            printrow=28;
            cdnum = 15;
        }else if (cdsl ==5){
            endrow = 31;
            printrow=32;
            cdnum = 18;
        }
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            RowCopy.copyRows(wb, sheetname, sheetname, 0, endrow, i* (endrow+1));
        }
        if(record >= 1){
            wb.setPrintArea(wb.getSheetIndex(sheetname), 0, cdnum, 0,(record) * printrow-1);
        }
    }


    /**
     * 创建隧道匝道的页数
     * @param tableNum
     * @param wb
     * @param sheetname
     * @param cdsl
     */
    private void createTable(int tableNum, XSSFWorkbook wb, String sheetname, int cdsl) {
        int endrow = 0;
        int printrow = 0;
        int cc= 0;
        if (cdsl == 2 || cdsl == 3){
            endrow = 46;
            printrow=47;
            cc=42;
        }else if (cdsl == 4){
            endrow = 27;
            printrow=28;
            cc=23;
        }else if (cdsl ==5){
            endrow = 31;
            printrow=32;
            cc=27;
        }
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            RowCopy.copyRows(wb, sheetname, sheetname, 5, endrow, (i-1)* cc+printrow);

        }
        if(record >= 1){
            wb.setPrintArea(wb.getSheetIndex(sheetname), 0, cdsl*2+4, 0,(record) * cc+4);

        }
    }


    /**
     *
     * @param data
     * @return
     */
    private int getNum(List<Map<String, Object>> data,int cdsl) {
        int a = 0;
        if (cdsl == 2){
            a = 42;
        }else if (cdsl == 3){
            a = 42;
        }else if (cdsl == 4){
            a = 23;
        }else if (cdsl == 5){
            a = 27;
        }
        Map<String, Integer> resultMap = new HashMap<>();
        for (Map<String, Object> map : data) {
            String name = map.get("name").toString();
            if (resultMap.containsKey(name)) {
                resultMap.put(name, resultMap.get(name) + 1);
            } else {
                resultMap.put(name, 1);
            }
        }
        int num = 0;
        for (Map.Entry<String, Integer> entry : resultMap.entrySet()) {
            int value = entry.getValue();
            if (value%a==0){
                num += value/a;
            }else {
                num += value/a+1;
            }
        }
        return num;
    }

    /**
     *
     * @param dataLists
     * @param cds
     * @return
     */
    private List<Map<String, Object>> mergedList(List<Map<String, Object>> dataLists,int cds) {
        //处理拼接的iri
        int cdsl = cds;
        String iris = "-";
        StringBuilder iriBuilder = new StringBuilder();
        for (int i = 0; i < cdsl; i++) {
            if (i == cdsl - 1) {
                iriBuilder.append("-");
            } else {
                iriBuilder.append("-,");
            }
        }
        iris = iriBuilder.toString();

        Map<String, List<Map<String, Object>>> groupedData =
                dataLists.stream()
                        .collect(Collectors.groupingBy(
                                item -> item.get("name") + "-" + item.get("zh")
                        ));
        List<Map<String, Object>> toBeRemoved = new ArrayList<>();
        for (Map.Entry<String, List<Map<String, Object>>> entry : groupedData.entrySet()) {
            String groupName = entry.getKey(); // 获取分组名字
            List<Map<String, Object>> groupData = entry.getValue(); // 获取分组数据

            if (groupData.size() > 1) {
                Map<String, Object> firstItem = groupData.get(0);
                Map<String, Object> secondItem = groupData.get(1);
                if (firstItem.get("cd").toString().contains("左幅")) {
                    // 将第二条数据的iri拼接在第一条数据的iri后面
                    String iri = firstItem.get("ld").toString() + "," + secondItem.get("ld").toString();
                    firstItem.put("ld", iri);
                } else {
                    // 将第二条数据的iri拼接在第一条数据的iri前面
                    String iri = secondItem.get("ld").toString() + "," + firstItem.get("ld").toString();
                    firstItem.put("ld", iri);
                }
                // 将secondItem条数据删除
                toBeRemoved.add(secondItem);

            } else if (groupData.size() == 1) {
                Map<String, Object> item = groupData.get(0);
                if (item.get("cd").toString().contains("左幅")) {
                    // 在iri后面拼接逗号
                    String iri = item.get("ld").toString() +","+ iris;
                    item.put("ld", iri);
                } else {
                    // 在iri前面拼接逗号
                    String iri = iris +","+ item.get("ld").toString();
                    item.put("ld", iri);
                }
            }
        }

        dataLists.removeAll(toBeRemoved);

        Collections.sort(dataLists, new Comparator<Map<String, Object>>() {
            @Override
            public int compare(Map<String, Object> o1, Map<String, Object> o2) {
                String name1 = o1.get("name").toString();
                String name2 = o2.get("name").toString();
                // 按照名字进行排序
                int cmp = name1.compareTo(name2);
                if (cmp != 0) {
                    return cmp;
                }
                // 名字相同时按照 qdzh 排序
                Double qdzh1 = Double.parseDouble(o1.get("zh").toString());
                Double qdzh2 = Double.parseDouble(o2.get("zh").toString());
                return qdzh1.compareTo(qdzh2);
            }
        });


        return dataLists;
    }



    /**
     *
     * @param list
     * @return
     */
    private static List<Map<String, Object>> montageld(List<Map<String, Object>> list) {
        if (list == null || list.isEmpty()){
            return new ArrayList<>();
        }else {
            Map<String, List<String>> resultMap = new TreeMap<>();
            for (Map<String, Object> map : list) {
                String zh = map.get("zh").toString();
                String sfc = map.get("ld").toString();
                if (resultMap.containsKey(zh)) {
                    resultMap.get(zh).add(sfc);
                } else {
                    List<String> sfcList = new ArrayList<>();
                    sfcList.add(sfc);
                    resultMap.put(zh, sfcList);
                }
            }
            List<Map<String, Object>> resultList = new LinkedList<>();
            for (Map.Entry<String, List<String>> entry : resultMap.entrySet()) {
                Map<String, Object> map = new TreeMap<>();
                map.put("zh", entry.getKey());
                map.put("ld", String.join(",", entry.getValue()));
                for (Map<String, Object> item : list) {
                    if (item.get("zh").toString().equals(entry.getKey())) {
                        map.put("name", item.get("name"));
                        map.put("cd", item.get("cd").toString().substring(0,2));
                        map.put("createTime", item.get("createTime"));
                        if (item.get("zdbs") != null){
                            map.put("zdbs", item.get("zdbs"));
                        }
                        break;
                    }
                }
                resultList.add(map);
            }

            Collections.sort(resultList, new Comparator<Map<String, Object>>() {
                @Override
                public int compare(Map<String, Object> o1, Map<String, Object> o2) {
                    String name1 = o1.get("name").toString();
                    String name2 = o2.get("name").toString();
                    // 按照名字进行排序
                    int cmp = name1.compareTo(name2);
                    if (cmp != 0) {
                        return cmp;
                    }
                    // 名字相同时按照 qdzh 排序
                    Double qdzh1 = Double.parseDouble(o1.get("zh").toString());
                    Double qdzh2 = Double.parseDouble(o2.get("zh").toString());
                    return qdzh1.compareTo(qdzh2);
                }
            });


            return resultList;
        }
    }

    /**
     *
     * @param list
     * @return
     */
    private static List<Map<String, Object>> montageldzd(List<Map<String, Object>> list) {
        if (list == null || list.isEmpty()){
            return new ArrayList<>();
        }else {
            Map<String, Map<String, List<String>>> map = new HashMap<>();
            for (Map<String, Object> item : list) {
                String zdbs = String.valueOf(item.get("zdbs"));
                String qdzh = String.valueOf(item.get("zh"));
                String sfc = String.valueOf(item.get("ld"));
                if (map.containsKey(zdbs)) {
                    Map<String, List<String>> mapItem = map.get(zdbs);
                    if (mapItem.containsKey(qdzh)) {
                        mapItem.get(qdzh).add(sfc);
                    } else {
                        List<String> sfcList = new ArrayList<>();
                        sfcList.add(sfc);
                        mapItem.put(qdzh, sfcList);
                    }
                } else {
                    Map<String, List<String>> mapItem = new HashMap<>();
                    List<String> sfcList = new ArrayList<>();
                    sfcList.add(sfc);
                    mapItem.put(qdzh, sfcList);
                    map.put(zdbs, mapItem);
                }
            }
            List<Map<String, Object>> result = new ArrayList<>();
            for (Map.Entry<String, Map<String, List<String>>> entry : map.entrySet()) {
                Map<String, Object> values = new HashMap<>();
                values.put("zdbs", entry.getKey());
                Map<String, List<String>> mapValue = entry.getValue();
                for (Map.Entry<String, List<String>> innerEntry : mapValue.entrySet()) {
                    String qdzh = innerEntry.getKey();
                    List<String> sfcList = innerEntry.getValue();
                    String sfc = String.join(",", sfcList);
                    values.put("zh", qdzh);
                    values.put("ld", sfc);
                    // 遍历整个list，查找相同的name和createTime
                    boolean flag = true;
                    for (Map<String, Object> item : list) {
                        if (!String.valueOf(item.get("zh")).equals(qdzh) || !String.valueOf(item.get("zdbs")).equals(entry.getKey())) {
                            continue;
                        }
                        if (flag) { // 第一次找到匹配的元素，将name和createTime保存到values中
                            values.put("name", item.get("name"));
                            values.put("createTime", item.get("createTime"));
                            values.put("cd", item.get("cd"));
                            flag = false;
                        }
                        // 如果name和createTime不同，则跳出循环
                        if (!item.get("name").equals(values.get("name")) || !item.get("createTime").equals(values.get("createTime"))) {
                            flag = false;
                            break;
                        }
                    }

                    result.add(new HashMap<>(values));
                }
            }

            Collections.sort(result, new Comparator<Map<String, Object>>() {
                @Override
                public int compare(Map<String, Object> o1, Map<String, Object> o2) {
                    String name1 = o1.get("zdbs").toString();
                    String name2 = o2.get("zdbs").toString();
                    // 按照名字进行排序
                    int cmp = name1.compareTo(name2);
                    if (cmp != 0) {
                        return cmp;
                    }
                    // 名字相同时按照 qdzh 排序
                    Double qdzh1 = Double.parseDouble(o1.get("zh").toString());
                    Double qdzh2 = Double.parseDouble(o2.get("zh").toString());
                    return qdzh1.compareTo(qdzh2);
                }
            });


            return result;
        }
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        List<Map<String, Object>> mapList = new ArrayList<>();
        List<Map<String,Object>> lxlist = jjgZdhLdhdMapper.selectlx(proname,htd);

        if (lxlist.size()>0){
            for (Map<String, Object> map : lxlist) {
                String zx = map.get("lxbs").toString();
                int num = jjgZdhLdhdMapper.selectcdnum(proname,htd,zx);
                int n = 0;
                if (num == 1){
                    n=2;
                }else {
                    n=num/2;
                }
                List<Map<String, Object>> looksdjdb = lookjdb(proname, htd, zx,n);
                mapList.addAll(looksdjdb);
            }
            return mapList;
        }else {
            return new ArrayList<>();
        }
    }

    /**
     *
     * @param proname
     * @param htd
     * @param zx
     * @param cds
     * @return
     */
    private List<Map<String, Object>> lookjdb(String proname, String htd, String zx, int cds) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        File f;
        if (zx.equals("主线")){
            f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "16路面雷达厚度.xlsx");
        }else {
            f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "16路面雷达厚度-"+zx+".xlsx");
        }
        if (!f.exists()) {
            return new ArrayList<>();
        } else {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
            List<Map<String, Object>> jgmap = new ArrayList<>();
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                if (!wb.isSheetHidden(wb.getSheetIndex(wb.getSheetAt(j)))) {
                    int zc = 0;
                    int dd = 0;
                    int m = 0;
                    if (wb.getSheetAt(j).getSheetName().equals("左幅") || wb.getSheetAt(j).getSheetName().equals("右幅") ) {
                        zc = cds*3+6;
                        m = cds*3+4;
                        if (cds == 2){
                            dd = 38;
                        }else if (cds == 3){
                            dd = 38;
                        }else if (cds == 4 ){
                            dd = 22;
                        }else if(cds ==5){
                            dd = 26;
                        }
                    }else {
                        zc = cds*3+7;
                        dd = 9;
                    }
                    XSSFSheet slSheet = wb.getSheetAt(j);

                    XSSFCell xmname = slSheet.getRow(1).getCell(1);//项目名
                    XSSFCell htdname = slSheet.getRow(1).getCell(cds*2+3);//合同段名
                    Map map = new HashMap();

                    if (proname.equals(xmname.toString()) && htd.equals(htdname.toString())) {
                        slSheet.getRow(4).getCell(zc).setCellType(CellType.STRING);//总点数
                        slSheet.getRow(4).getCell(zc+1).setCellType(CellType.STRING);//合格点数

                        slSheet.getRow(4).getCell(m).setCellType(CellType.STRING);
                        slSheet.getRow(4).getCell(m+1).setCellType(CellType.STRING);

                        double zds = Double.valueOf(slSheet.getRow(4).getCell(zc).getStringCellValue());
                        double hgds = Double.valueOf(slSheet.getRow(4).getCell(zc+1).getStringCellValue());
                        String zdsz1 = decf.format(zds);
                        String hgdsz1 = decf.format(hgds);
                        map.put("检测项目", zx);
                        map.put("路面类型", wb.getSheetName(j));
                        map.put("总点数", zdsz1);
                        map.put("设计值", slSheet.getRow(dd).getCell(cds*2+4).getStringCellValue());
                        map.put("合格点数", hgdsz1);
                        map.put("合格率", zds!=0 ? df.format(hgds/zds*100) : 0);
                        map.put("Max", slSheet.getRow(4).getCell(m).getStringCellValue());
                        map.put("Min", slSheet.getRow(4).getCell(m+1).getStringCellValue());
                    }
                    jgmap.add(map);

                }
            }
            return jgmap;
        }
    }

    @Override
    public void exportldhd(HttpServletResponse response, String cdsl) throws IOException {
        int cd = Integer.parseInt(cdsl);
        String fileName = "雷达厚度实测数据";
        String[][] sheetNames = {
                {"左幅一车道","左幅二车道","右幅一车道","右幅二车道"},
                {"左幅一车道","左幅二车道","左幅三车道","右幅一车道","右幅二车道","右幅三车道"},
                {"左幅一车道","左幅二车道","左幅三车道","左幅四车道","右幅一车道","右幅二车道","右幅三车道","右幅四车道"},
                {"左幅一车道","左幅二车道","左幅三车道","左幅四车道","左幅五车道","右幅一车道","右幅二车道","右幅三车道","右幅四车道","右幅五车道"}
        };
        String[] sheetName = sheetNames[cd-2];
        ExcelUtil.writeExcelMultipleSheets(response, null, fileName, sheetName, new JjgZdhLdhdVo());

    }

    @Override
    public void importldhd(MultipartFile file, CommonInfoVo commonInfoVo) throws IOException {
        // 获取文件输入流
        InputStream inputStream = file.getInputStream();
        // 创建工作簿
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        int number = workbook.getNumberOfSheets();
        for (int i = 0; i < number; i++) {
            String sheetName = workbook.getSheetName(i);
            int sheetIndex = workbook.getSheetIndex(workbook.getSheetAt(i));
            try {
                EasyExcel.read(file.getInputStream())
                        .sheet(sheetIndex)
                        .head(JjgZdhLdhdVo.class)
                        .headRowNumber(1)
                        .registerReadListener(
                                new ExcelHandler<JjgZdhLdhdVo>(JjgZdhLdhdVo.class) {
                                    @Override
                                    public void handle(List<JjgZdhLdhdVo> dataList) {
                                        for(JjgZdhLdhdVo ldhdVo: dataList)
                                        {
                                            JjgZdhLdhd ldhd = new JjgZdhLdhd();
                                            BeanUtils.copyProperties(ldhdVo,ldhd);
                                            ldhd.setCreatetime(new Date());
                                            ldhd.setProname(commonInfoVo.getProname());
                                            ldhd.setHtd(commonInfoVo.getHtd());
                                            ldhd.setZh(Double.parseDouble(ldhdVo.getZh()));
                                            ldhd.setCd(sheetName);
                                            jjgZdhLdhdMapper.insert(ldhd);
                                        }
                                    }
                                }
                        ).doRead();
            } catch (IOException e) {
                throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
            }
        }

        // 关闭输入流
        inputStream.close();

    }
}
