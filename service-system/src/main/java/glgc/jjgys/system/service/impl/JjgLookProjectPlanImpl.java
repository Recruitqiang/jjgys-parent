package glgc.jjgys.system.service.impl;


import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgHtd;
import glgc.jjgys.model.project.JjgPlaninfo;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.lqs.JjgPlaninfoVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgLookProjectPlanMapper;
import glgc.jjgys.system.service.*;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;


@Service
public class JjgLookProjectPlanImpl extends ServiceImpl<JjgLookProjectPlanMapper, JjgPlaninfo> implements JjgLookProjectPlanService {


    @Autowired
    private JjgHtdService jjgHtdService;

    @Autowired
    private JjgFbgcLjgcHdgqdService jjgFbgcLjgcHdgqdService;

    @Autowired
    private JjgFbgcLjgcHdjgccService jjgFbgcLjgcHdjgccService;

    @Autowired
    private JjgFbgcLjgcLjtsfysdSlService jjgFbgcLjgcLjtsfysdSlService;

    @Autowired
    private JjgFbgcLjgcLjtsfysdHtService jjgFbgcLjgcLjtsfysdHtService;

    @Autowired
    private JjgFbgcLjgcLjwcService jjgFbgcLjgcLjwcService;

    @Autowired
    private JjgFbgcLjgcLjwcLcfService jjgFbgcLjgcLjwcLcfService;

    @Autowired
    private JjgFbgcLjgcLjcjService jjgFbgcLjgcLjcjService;

    @Autowired
    private JjgFbgcLjgcLjbpService jjgFbgcLjgcLjbpService;

    @Autowired
    private JjgFbgcLjgcPsdmccService jjgFbgcLjgcPsdmccService;

    @Autowired
    private JjgFbgcLjgcPspqhdService jjgFbgcLjgcPspqhdService;

    @Autowired
    private JjgFbgcLjgcXqgqdService jjgFbgcLjgcXqgqdService;

    @Autowired
    private JjgFbgcLjgcXqjgccService jjgFbgcLjgcXqjgccService;

    @Autowired
    private JjgFbgcLjgcZddmccService jjgFbgcLjgcZddmccService;

    @Autowired
    private JjgFbgcLjgcZdgqdService jjgFbgcLjgcZdgqdService;

    @Autowired
    private JjgFbgcJtaqssJabxService jjgFbgcJtaqssJabxService;

    @Autowired
    private JjgFbgcJtaqssJabxfhlService jjgFbgcJtaqssJabxfhlService;

    @Autowired
    private JjgFbgcJtaqssJabzService jjgFbgcJtaqssJabzService;

    @Autowired
    private JjgFbgcJtaqssJathldmccService jjgFbgcJtaqssJathldmccService;

    @Autowired
    private JjgFbgcJtaqssJathlqdService jjgFbgcJtaqssJathlqdService;

    @Autowired
    private JjgFbgcLmgcGslqlmhdzxfService jjgFbgcLmgcGslqlmhdzxfService;

    @Autowired
    private JjgFbgcLmgcHntlmhdzxfService jjgFbgcLmgcHntlmhdzxfService;

    @Autowired
    private JjgFbgcLmgcHntlmqdService jjgFbgcLmgcHntlmqdService;

    @Autowired
    private JjgFbgcLmgcLmgzsdsgpsfService jjgFbgcLmgcLmgzsdsgpsfService;

    @Autowired
    private JjgFbgcLmgcLmhpService jjgFbgcLmgcLmhpService;

    @Autowired
    private JjgFbgcLmgcLmssxsService jjgFbgcLmgcLmssxsService;

    @Autowired
    private JjgFbgcLmgcLmwcService jjgFbgcLmgcLmwcService;

    @Autowired
    private JjgFbgcLmgcLmwcLcfService jjgFbgcLmgcLmwcLcfService;

    @Autowired
    private JjgFbgcLmgcLqlmysdService jjgFbgcLmgcLqlmysdService;

    @Autowired
    private JjgFbgcLmgcTlmxlbgcService jjgFbgcLmgcTlmxlbgcService;

    @Autowired
    private JjgLookProjectPlanMapper jjgLookProjectPlanMapper;




    @Override
    public List<Map<String,Object>> lookplan(CommonInfoVo commonInfoVo) {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        DecimalFormat df = new DecimalFormat("0.00");
        //当前合同下的项目计划情况
        List<Map<String,Object>> planList = jjgLookProjectPlanMapper.selectplan(proname,htd);
        /**
         * [{fbgc=涵洞砼强度, num=40, htd=LJ-1, proname=陕西高速},
         * {fbgc=涵洞结构尺寸, num=40, htd=LJ-1, proname=陕西高速},
         * {fbgc=排水断面尺寸, num=50, htd=LJ-1, proname=陕西高速},
         * {fbgc=排水铺砌厚度, num=20, htd=LJ-1, proname=陕西高速}]
         */
        List<Map<String,Object>> resultmap = new ArrayList<>();
        for (Map<String, Object> map : planList) {
            String zb = map.get("zb").toString();
            if (zb.contains("涵洞砼强度")){
                int selectnum = jjgFbgcLjgcHdgqdService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("涵洞结构尺寸")){
                int selectnum = jjgFbgcLjgcHdjgccService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("路基土石方压实度沙砾")){
                int selectnum = jjgFbgcLjgcLjtsfysdSlService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("路基土石方压实度灰土")){
                int selectnum = jjgFbgcLjgcLjtsfysdHtService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("压实度沉降")){
                int selectnum = jjgFbgcLjgcLjcjService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("路基弯沉贝克曼梁法")){
                int selectnum = jjgFbgcLjgcLjwcService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("路基弯沉落锤法")){
                int selectnum = jjgFbgcLjgcLjwcLcfService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            } else if (zb.contains("路基边坡")){
                int selectnum = jjgFbgcLjgcLjbpService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("排水断面尺寸")){
                int selectnum = jjgFbgcLjgcPsdmccService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("排水铺砌厚度")){
                int selectnum = jjgFbgcLjgcPspqhdService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("小桥砼强度")){
                int selectnum = jjgFbgcLjgcXqgqdService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("小桥结构尺寸")){
                int selectnum = jjgFbgcLjgcXqjgccService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("支挡砼强度")){
                int selectnum = jjgFbgcLjgcZdgqdService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("支挡断面尺寸")){
                int selectnum = jjgFbgcLjgcZddmccService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            } else if (zb.contains("交安标线")){
                int selectnum = jjgFbgcJtaqssJabxService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("交安标志")){
                int selectnum = jjgFbgcJtaqssJabzService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("交安波形防护栏")){
                int selectnum = jjgFbgcJtaqssJabxfhlService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("交安砼护栏强度")){
                int selectnum = jjgFbgcJtaqssJathlqdService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("交安砼护栏断面尺寸")){
                int selectnum = jjgFbgcJtaqssJathldmccService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("高速沥青路面厚度钻芯法")){
                int selectnum = jjgFbgcLmgcGslqlmhdzxfService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("混凝土路面厚度钻芯法")){
                int selectnum = jjgFbgcLmgcHntlmhdzxfService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("混凝土路面强度")){
                int selectnum = jjgFbgcLmgcHntlmqdService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("路面构造深度手工铺砂法")){
                int selectnum = jjgFbgcLmgcLmgzsdsgpsfService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("路面横坡")){
                int selectnum = jjgFbgcLmgcLmhpService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("沥青路面渗水系数")){
                int selectnum = jjgFbgcLmgcLmssxsService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("路面弯沉贝克曼梁法")){
                int selectnum = jjgFbgcLmgcLmwcService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("路面弯沉落锤法")){
                int selectnum = jjgFbgcLmgcLmwcLcfService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("沥青路面压实度")){
                int selectnum = jjgFbgcLmgcLqlmysdService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }else if (zb.contains("砼路面相邻板高差")){
                int selectnum = jjgFbgcLmgcTlmxlbgcService.selectnum(proname, htd);
                Double num = Double.valueOf(map.get("num").toString());
                map.put("jcs",selectnum);
                map.put("jdl",num!=0?df.format(selectnum/num):0);
                resultmap.add(map);
            }
        }
        return resultmap;
    }

    @Override
    public void exportxmjd(HttpServletResponse response, String proname) {
        String fileName = proname+"指标计划情况";
        String sheetName = "指标计划清单";
        List<JjgPlaninfoVo> list = new ArrayList<>();

        /**
         * 先获取到项目下的所有合同段，然后判断这个合同段的类型
         * 根据类型不同，产生不同的指标
         */
        String[] ljlist = {"涵洞砼强度","涵洞结构尺寸","路基土石方压实度沙砾","路基土石方压实度灰土","压实度沉降","路基弯沉贝克曼梁法","路基弯沉落锤法","路基边坡","排水断面尺寸","排水铺砌厚度","小桥砼强度","小桥结构尺寸","支挡砼强度","支挡断面尺寸"};
        String[] jalist = {"交安标线","交安标志","交安波形防护栏","交安砼护栏强度","交安砼护栏断面尺寸"};
        String[] lmlist = {"高速沥青路面厚度钻芯法","混凝土路面厚度钻芯法","混凝土路面强度","路面构造深度手工铺砂法","路面横坡","沥青路面渗水系数","路面弯沉贝克曼梁法","路面弯沉落锤法","沥青路面压实度","砼路面相邻板高差"};

        List<JjgHtd> gethtd = jjgHtdService.gethtd(proname);
        for (JjgHtd jjgHtd : gethtd) {
            String lx = jjgHtd.getLx();
            String htd = jjgHtd.getName();
            String[] arr = lx.split(",");
            for (String s : arr) {
                if (s.equals("路基工程")){
                    for (String s1 : ljlist) {
                        JjgPlaninfoVo jjgPlaninfoVo = new JjgPlaninfoVo();
                        jjgPlaninfoVo.setProname(proname);
                        jjgPlaninfoVo.setHtd(htd);
                        jjgPlaninfoVo.setFbgc(s);
                        jjgPlaninfoVo.setZb(s1);
                        list.add(jjgPlaninfoVo);
                    }
                }else if (s.equals("路面工程")){
                    for (String s1 : lmlist) {
                        JjgPlaninfoVo jjgPlaninfoVo = new JjgPlaninfoVo();
                        jjgPlaninfoVo.setProname(proname);
                        jjgPlaninfoVo.setHtd(htd);
                        jjgPlaninfoVo.setFbgc(s);
                        jjgPlaninfoVo.setZb(s1);
                        list.add(jjgPlaninfoVo);
                    }
                }else if (s.equals("交安工程")){
                    for (String s1 : jalist) {
                        JjgPlaninfoVo jjgPlaninfoVo = new JjgPlaninfoVo();
                        jjgPlaninfoVo.setProname(proname);
                        jjgPlaninfoVo.setHtd(htd);
                        jjgPlaninfoVo.setFbgc(s);
                        jjgPlaninfoVo.setZb(s1);
                        list.add(jjgPlaninfoVo);
                    }
                }
            }
        }
        ExcelUtil.writeExcelWithSheets(response, list, fileName, sheetName, new JjgPlaninfoVo()).finish();

    }

    @Override
    public void importxmjd(MultipartFile file, String proname) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgPlaninfoVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgPlaninfoVo>(JjgPlaninfoVo.class) {
                                @Override
                                public void handle(List<JjgPlaninfoVo> dataList) {
                                    for(JjgPlaninfoVo ql: dataList)
                                    {
                                        JjgPlaninfo jjgPlaninfo = new JjgPlaninfo();
                                        BeanUtils.copyProperties(ql,jjgPlaninfo);
                                        jjgPlaninfo.setCreateTime(new Date());
                                        System.out.println(jjgPlaninfo);
                                        jjgLookProjectPlanMapper.insert(jjgPlaninfo);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }
}
