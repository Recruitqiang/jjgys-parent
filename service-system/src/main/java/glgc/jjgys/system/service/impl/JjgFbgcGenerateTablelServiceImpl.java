package glgc.jjgys.system.service.impl;

import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.model.project.JjgHtd;
import glgc.jjgys.model.project.JjgLqsQl;
import glgc.jjgys.model.project.JjgLqsSd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.mapper.JjgFbgcGenerateTableMapper;
import glgc.jjgys.system.service.*;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-03-01
 */
@Service
public class JjgFbgcGenerateTablelServiceImpl extends ServiceImpl<JjgFbgcGenerateTableMapper,Object> implements JjgFbgcGenerateTablelService {

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @Autowired
    private JjgFbgcLjgcHdgqdService jjgFbgcLjgcHdgqdService;

    @Autowired
    private JjgFbgcLjgcHdjgccService jjgFbgcLjgcHdjgccService;

    @Autowired
    private JjgFbgcLjgcLjbpService jjgFbgcLjgcLjbpService;

    @Autowired
    private JjgFbgcLjgcLjcjService jjgFbgcLjgcLjcjService;

    @Autowired
    private JjgFbgcLjgcLjtsfysdHtService jjgFbgcLjgcLjtsfysdHtService;

    @Autowired
    private JjgFbgcLjgcLjwcService jjgFbgcLjgcLjwcService;

    @Autowired
    private JjgFbgcLjgcLjwcLcfService jjgFbgcLjgcLjwcLcfService;

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
    private JjgFbgcQlgcXbTqdService jjgFbgcQlgcXbTqdService;

    @Autowired
    private JjgFbgcQlgcXbJgccService jjgFbgcQlgcXbJgccService;

    @Autowired
    private JjgFbgcQlgcXbBhchdService jjgFbgcQlgcXbBhchdService;

    @Autowired
    private JjgFbgcQlgcXbSzdService jjgFbgcQlgcXbSzdService;

    @Autowired
    private JjgFbgcQlgcSbTqdService jjgFbgcQlgcSbTqdService;

    @Autowired
    private JjgFbgcQlgcSbJgccService jjgFbgcQlgcSbJgccService;

    @Autowired
    private JjgFbgcQlgcSbBhchdService jjgFbgcQlgcSbBhchdService;

    @Autowired
    private JjgFbgcSdgcCqtqdService jjgFbgcSdgcCqtqdService;

    @Autowired
    private JjgFbgcSdgcDmpzdService jjgFbgcSdgcDmpzdService;

    @Autowired
    private JjgFbgcSdgcZtkdService jjgFbgcSdgcZtkdService;

    @Autowired
    private JjgFbgcQlgcQmpzdService jjgFbgcQlgcQmpzdService;

    @Autowired
    private JjgFbgcQlgcQmhpService jjgFbgcQlgcQmhpService;

    @Autowired
    private JjgFbgcQlgcQmgzsdService jjgFbgcQlgcQmgzsdService;

    @Autowired
    private JjgFbgcSdgcCqhdService jjgFbgcSdgcCqhdService;

    @Autowired
    private JjgFbgcSdgcSdlqlmysdService jjgFbgcSdgcSdlqlmysdService;

    @Autowired
    private JjgFbgcSdgcLmssxsService jjgFbgcSdgcLmssxsService;

    @Autowired
    private JjgFbgcSdgcHntlmqdService jjgFbgcSdgcHntlmqdService;

    @Autowired
    private JjgFbgcSdgcTlmxlbgcService jjgFbgcSdgcTlmxlbgcService;

    @Autowired
    private JjgFbgcSdgcSdhpService jjgFbgcSdgcSdhpService;

    @Autowired
    private JjgFbgcSdgcSdhntlmhdzxfService jjgFbgcSdgcSdhntlmhdzxfService;

    @Autowired
    private JjgFbgcSdgcGssdlqlmhdzxfService jjgFbgcSdgcGssdlqlmhdzxfService;

    @Autowired
    private JjgFbgcSdgcLmgzsdsgpsfService jjgFbgcSdgcLmgzsdsgpsfService;

    @Autowired
    private JjgHtdService jjgHtdService;


    @Autowired
    private JjgZdhGzsdService jjgZdhGzsdService;

    @Autowired
    private JjgZdhMcxsService jjgZdhMcxsService;

    @Autowired
    private JjgZdhPzdService jjgZdhPzdService;

    @Autowired
    private JjgZdhLdhdService jjgZdhLdhdService;

    @Autowired
    private JjgZdhCzService jjgZdhCzService;

    @Autowired
    private JjgLqsSdService jjgLqsSdService;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;


    @Override
    public void generatePdb(CommonInfoVo commonInfoVo) throws IOException {

        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");

        /**
         * 路基工程，路面工程，交安工程
         * 桥梁工程和隧道工程是按  桥梁和隧道
         *
         * 或许也可以不用查询这个合同段的类型，根据当前项目和合同段路基下的所有文件
         * 只需要把有关路基，路面和交安相关的文件筛选出来。
         */

        String path = filespath+ File.separator+proname+File.separator+htd+File.separator;

        List<Map<String,Object>> resultlist = new ArrayList<>();
        List<String> filteredFiles = filterFiles(path);

        for (String value : filteredFiles) {
            switch (value) {
                case "08路基涵洞砼强度.xlsx":
                    // 路基涵洞
                    List<Map<String, Object>> maps1 = jjgFbgcLjgcHdgqdService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps1 = new ArrayList<>();

                    List<String> sjqdlist = jjgFbgcLjgcHdgqdService.selectsjqd(proname,htd);
                    for (String s : sjqdlist) {
                        for (Map<String, Object> map : maps1) {
                            Map<String, Object> newMap = new HashMap<>(map);
                            newMap.put("filename", "详见《涵洞砼强度质量鉴定表》检测"+map.get("总点数")+"点,合格"+map.get("合格点数")+"点");
                            newMap.put("ccname", "*砼强度");
                            newMap.put("yxps", "C"+s);
                            newMap.put("sheetname", "分部-路基");
                            newMap.put("fbgc", "涵洞");
                            newMaps1.add(newMap);
                        }
                    }
                    maps1 = newMaps1;
                    resultlist.addAll(maps1);
                    break;
                case "09路基涵洞结构尺寸.xlsx":
                    // 路基涵洞结构尺寸
                    List<Map<String, Object>> maps2 = jjgFbgcLjgcHdjgccService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps2 = new ArrayList<>();

                    List<String> yxpslist = jjgFbgcLjgcHdjgccService.selectyxps(proname,htd);
                    for (String s : yxpslist) {
                        for (Map<String, Object> map : maps2) {
                            Map<String, Object> newMap = new HashMap<>(map);
                            newMap.put("filename", "详见《涵洞结构尺寸质量鉴定表》检测"+map.get("总点数")+"点,合格"+map.get("合格点数")+"点");
                            newMap.put("ccname", "结构尺寸");
                            newMap.put("yxps", s);
                            newMap.put("sheetname", "分部-路基");
                            newMap.put("fbgc", "涵洞");
                            newMaps2.add(newMap);
                        }
                    }
                    maps2 = newMaps2;
                    resultlist.addAll(maps2);
                    break;
                case "03路基边坡.xlsx":
                    List<Map<String, Object>> maps3 = jjgFbgcLjgcLjbpService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps3 = new ArrayList<>();

                    //List<String> yxpslistbp = jjgFbgcLjgcLjbpService.selectyxps(proname,htd);暂时先这样
                    for (Map<String, Object> map : maps3) {
                        Map<String, Object> newMap = new HashMap<>(map);
                        newMap.put("filename", "详见《路基边坡质量鉴定表》检测"+map.get("总点数")+"点,合格"+map.get("合格点数")+"点");
                        newMap.put("ccname", "边坡");
                        newMap.put("yxps", "不陡于设计");//先写死
                        newMap.put("sheetname", "分部-路基");
                        newMap.put("fbgc", "路基土石方");
                        newMaps3.add(newMap);
                    }
                    maps3 = newMaps3;
                    resultlist.addAll(maps3);
                    break;
                case "01路基压实度沉降.xlsx":
                    List<Map<String, Object>> maps4 = jjgFbgcLjgcLjcjService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps4 = new ArrayList<>();

                    List<String> yxpslistcj = jjgFbgcLjgcLjcjService.selectyxps(proname,htd);
                    for (String s : yxpslistcj) {
                        for (Map<String, Object> map : maps4) {
                            Map<String, Object> newMap = new HashMap<>(map);
                            newMap.put("filename", "详见《路基压实度沉降质量鉴定表》检测"+map.get("总点数")+"点,合格"+map.get("合格点数")+"点");
                            newMap.put("ccname", "△沉降");
                            newMap.put("yxps", "≤"+s);
                            newMap.put("sheetname", "分部-路基");
                            newMap.put("fbgc", "路基土石方");
                            newMaps4.add(newMap);
                        }
                    }
                    maps4 = newMaps4;
                    resultlist.addAll(maps4);
                    break;
                case "01路基土石方压实度.xlsx":
                    List<Map<String, Object>> maps5 = jjgFbgcLjgcLjtsfysdHtService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps5 = new ArrayList<>();
                    for (Map<String, Object> map : maps5) {
                        Map<String, Object> newMap = new HashMap<>(map);
                        newMap.put("filename", "详见《路基压实度质量鉴定表》检测"+map.get("检测点数")+"点,合格"+map.get("合格点数")+"点");
                        newMap.put("ccname", "△压实度"+map.get("压实度项目"));
                        newMap.put("yxps", map.get("规定值"));
                        newMap.put("sheetname", "分部-路基");
                        newMap.put("fbgc", "路基土石方");
                        newMaps5.add(newMap);

                    }
                    maps5 = newMaps5;
                    resultlist.addAll(maps5);
                    break;

                case "02路基弯沉(贝克曼梁法).xlsx":
                    List<Map<String, Object>> maps6 = jjgFbgcLjgcLjwcService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps6 = new ArrayList<>();
                    for (Map<String, Object> map : maps6) {
                        Map<String, Object> newMap = new HashMap<>(map);
                        newMap.put("filename", "详见《路基弯沉(贝克曼梁法)质量鉴定表》检测"+map.get("检测单元数")+"个评定单元,合格"+map.get("合格单元数")+"个评定单元");
                        newMap.put("ccname", "△弯沉(贝克曼梁法)");
                        newMap.put("sheetname", "分部-路基");
                        newMap.put("yxps", map.get("规定值"));
                        newMap.put("fbgc", "路基土石方");
                        newMaps6.add(newMap);
                    }
                    maps6 = newMaps6;
                    resultlist.addAll(maps6);
                    break;
                case "02路基弯沉(落锤法).xlsx":
                    List<Map<String, Object>> maps66 = jjgFbgcLjgcLjwcLcfService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps66 = new ArrayList<>();
                    for (Map<String, Object> map : maps66) {
                        Map<String, Object> newMap = new HashMap<>(map);
                        newMap.put("filename", "详见《路基弯沉(落锤法)质量鉴定表》检测"+map.get("检测单元数")+"个评定单元,合格"+map.get("合格单元数")+"个评定单元");
                        newMap.put("ccname", "△弯沉(落锤法)");
                        newMap.put("sheetname", "分部-路基");
                        newMap.put("yxps", map.get("规定值"));
                        newMap.put("fbgc", "路基土石方");
                        newMaps66.add(newMap);
                    }
                    maps66 = newMaps66;
                    resultlist.addAll(maps66);
                    break;
                case "04路基排水断面尺寸.xlsx":
                    List<Map<String, Object>> maps7 = jjgFbgcLjgcPsdmccService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps7 = new ArrayList<>();
                    List<Map<String,Object>> yxpslistcc = jjgFbgcLjgcPsdmccService.selectyxps(proname,htd);
                    for (Map<String, Object> s : yxpslistcc) {
                        for (Map<String, Object> map : maps7) {
                            Map<String, Object> newMap = new HashMap<>(map);
                            String va="";
                            if (s.get("yxwcz").equals(s.get("yxwcf"))){
                                va = s.get("sjz")+"±"+s.get("yxwcz");
                            }else {
                                va = s.get("sjz")+"+"+s.get("yxwcz")+";"+s.get("sjz")+"-"+s.get("yxwcf");
                            }
                            newMap.put("filename", "详见《结构（断面）尺寸质量鉴定表》检测"+map.get("检测总点数")+"点,合格"+map.get("合格点数")+"点");
                            newMap.put("ccname", "断面尺寸");
                            newMap.put("yxps", va);
                            newMap.put("sheetname", "分部-路基");
                            newMap.put("fbgc", "排水工程");
                            newMaps7.add(newMap);
                        }

                    }
                    maps7 = newMaps7;
                    resultlist.addAll(maps7);
                    break;
                case "05路基排水铺砌厚度.xlsx":
                    List<Map<String, Object>> maps8 = jjgFbgcLjgcPspqhdService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps8 = new ArrayList<>();
                    List<String> sjzlistpqhd = jjgFbgcLjgcPspqhdService.selectyxps(proname,htd);
                    for (String s : sjzlistpqhd) {
                        for (Map<String, Object> map : maps8) {
                            Map<String, Object> newMap = new HashMap<>(map);
                            newMap.put("filename", "详见《排水铺砌厚度质量鉴定表》检测"+map.get("检测总点数")+"点,合格"+map.get("合格点数")+"点");
                            newMap.put("ccname", "铺砌厚度");
                            newMap.put("yxps", s);
                            newMap.put("sheetname", "分部-路基");
                            newMap.put("fbgc", "排水工程");
                            newMaps8.add(newMap);
                        }
                    }
                    maps8 = newMaps8;
                    resultlist.addAll(maps8);
                    break;
                case "06路基小桥砼强度.xlsx":
                    List<Map<String, Object>> maps9 = jjgFbgcLjgcXqgqdService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps9 = new ArrayList<>();
                    List<String> sjzlistsjqd = jjgFbgcLjgcXqgqdService.selectsjqd(proname,htd);
                    for (String s : sjzlistsjqd) {
                        for (Map<String, Object> map : maps9) {
                            Map<String, Object> newMap = new HashMap<>(map);
                            newMap.put("filename", "详见《小桥砼强度质量鉴定表》检测"+map.get("检测总点数")+"点,合格"+map.get("合格点数")+"点");
                            newMap.put("ccname", "*砼强度");
                            newMap.put("yxps", s);
                            newMap.put("sheetname", "分部-路基");
                            newMap.put("fbgc", "小桥");
                            maps9.add(newMap);
                        }
                    }
                    maps9 = newMaps9;
                    resultlist.addAll(maps9);
                    break;
                case "07路基小桥结构尺寸.xlsx":
                    List<Map<String, Object>> maps10 = jjgFbgcLjgcXqjgccService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps10 = new ArrayList<>();
                    List<Map<String,Object>> yxpslistxq = jjgFbgcLjgcXqjgccService.selectyxps(proname,htd);
                    for (Map<String, Object> yxpcmap : yxpslistxq) {
                        for (Map<String, Object> jdjgmap : maps10) {
                            Map<String, Object> newMap = new HashMap<>(jdjgmap);
                            String va="";
                            if (yxpcmap.get("yxwcz").equals(yxpcmap.get("yxwcf"))){
                                va = "±"+yxpcmap.get("yxwcz");
                            }else {
                                va = "+"+yxpcmap.get("yxwcz")+";"+"-"+yxpcmap.get("yxwcf");
                            }
                            newMap.put("filename", "详见《小桥结构尺寸质量鉴定表》检测"+jdjgmap.get("检测总点数")+"点,合格"+jdjgmap.get("合格点数")+"点");
                            newMap.put("ccname", "主要结构尺寸");
                            newMap.put("yxps", va);
                            newMap.put("sheetname", "分部-路基");
                            newMap.put("fbgc", "小桥");
                            newMaps10.add(newMap);
                        }

                    }
                    maps10 = newMaps10;
                    resultlist.addAll(maps10);
                    break;
                case "11路基支挡断面尺寸.xlsx":
                    List<Map<String, Object>> maps11 = jjgFbgcLjgcZddmccService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps11 = new ArrayList<>();
                    List<Map<String,Object>> yxpslistzd = jjgFbgcLjgcZddmccService.selectyxps(proname,htd);
                    for (Map<String, Object> yxps : yxpslistzd) {
                        for (Map<String, Object> jdjg : maps11) {
                            Map<String, Object> newMap = new HashMap<>(jdjg);
                            newMap.put("filename", "详见《支挡工程结构尺寸质量鉴定表》检测"+jdjg.get("检测总点数")+"点,合格"+jdjg.get("合格点数")+"点");
                            newMap.put("ccname", "△断面尺寸");
                            newMap.put("yxps", yxps.get("result"));
                            newMap.put("sheetname", "分部-路基");
                            newMap.put("fbgc", "支挡工程");
                            newMaps11.add(newMap);
                        }
                    }
                    maps11 = newMaps11;
                    resultlist.addAll(maps11);
                    break;
                case "10路基支挡砼强度.xlsx":
                    List<Map<String, Object>> maps12 = jjgFbgcLjgcZdgqdService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps12 = new ArrayList<>();

                    List<Map<String,Object>> yxpslistzdtqd = jjgFbgcLjgcZdgqdService.selectsjqd(proname,htd);
                    for (Map<String, Object> sjqd : yxpslistzdtqd) {
                        for (Map<String, Object> jdjg : maps12) {

                            Map<String, Object> newMap = new HashMap<>(jdjg);
                            newMap.put("filename", "详见《支挡工程砼强度质量鉴定表》检测"+jdjg.get("总点数")+"点,合格"+jdjg.get("合格点数")+"点");
                            newMap.put("ccname", "*砼强度");
                            newMap.put("yxps", sjqd.get("sjqd"));
                            newMap.put("sheetname", "分部-路基");
                            newMap.put("fbgc", "支挡工程");
                            newMaps12.add(newMap);
                        }
                    }
                    maps12 = newMaps12;
                    resultlist.addAll(maps12);
                    break;

                case "56交安标志.xlsx":
                    List<Map<String, Object>> maps13 = jjgFbgcJtaqssJabzService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps13 = new ArrayList<>();
                    double zdsd = 0;
                    double hgdsd = 0;
                    String hgl = "";
                    for (Map<String, Object> jdjg : maps13) {
                        if (jdjg.get("项目").toString().contains("反光膜逆反射系数")){
                            double zds = Double.valueOf(jdjg.get("总点数").toString());
                            double hgds = Double.valueOf(jdjg.get("合格点数").toString());
                            zdsd+=zds;
                            hgdsd+=hgds;
                        }
                    }
                    if (zdsd != 0|| hgdsd !=0 ){
                        hgl = df.format(hgdsd/zdsd*100);
                    }else {
                        hgl = "0";
                    }
                    for (Map<String, Object> result : maps13) {
                        Map<String, Object> newMap = new HashMap<>(result);
                        if (result.get("项目").toString().contains("反光膜逆反射系数")){
                            newMap.put("filename", "详见《交通标志板安装质量鉴定表》检测"+zdsd+"点,合格"+hgdsd+"点");
                            newMap.put("合格率", hgl);
                            newMap.put("总点数", zdsd);
                            newMap.put("合格点数", hgdsd);

                        }else {
                            newMap.put("filename", "详见《交通标志板安装质量鉴定表》检测"+result.get("总点数")+"点,合格"+result.get("合格点数")+"点");
                        }
                        newMap.put("ccname", result.get("项目"));
                        newMap.put("yxps", result.get("规定值或允许偏差"));
                        newMap.put("sheetname", "分部-交安");
                        newMap.put("fbgc", "标志");
                        newMaps13.add(newMap);
                    }
                    resultlist.addAll(newMaps13);
                    break;
                case "57交安标线厚度.xlsx":
                    List<Map<String, Object>> maps14 = jjgFbgcJtaqssJabxService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps14 = new ArrayList<>();
                    for (Map<String, Object> jdjg : maps14) {
                        Map<String, Object> newMap = new HashMap<>(jdjg);
                        if (jdjg.get("检测项目").toString().equals("交安标线厚度")){
                            newMap.put("filename", "详见《道路交通标线施工质量鉴定表》检测"+jdjg.get("总点数")+"点,合格"+jdjg.get("合格点数")+"点");
                            newMap.put("ccname", "△标线厚度");
                            newMap.put("yxps", jdjg.get("规定值或允许偏差"));
                            newMap.put("sheetname", "分部-交安");
                            newMap.put("fbgc", "标线");
                            newMaps14.add(newMap);
                        }
                        if (jdjg.get("检测项目").toString().contains("逆反射系数")){
                            newMap.put("filename", "详见《道路交通标线施工质量鉴定表》检测"+jdjg.get("总点数")+"点,合格"+jdjg.get("合格点数")+"点");
                            newMap.put("ccname", "△反光标线逆反射亮度系数");
                            newMap.put("yxps", jdjg.get("规定值或允许偏差"));
                            newMap.put("sheetname", "分部-交安");
                            newMap.put("fbgc", "标线");
                            newMaps14.add(newMap);
                        }

                    }
                    maps14 = newMaps14;
                    resultlist.addAll(maps14);
                    break;
                case "58交安钢防护栏.xlsx":
                    List<Map<String, Object>> maps15 = jjgFbgcJtaqssJabxfhlService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps15 = new ArrayList<>();
                    for (Map<String, Object> jdjg : maps15) {
                        Map<String, Object> newMap = new HashMap<>(jdjg);
                        newMap.put("filename", "详见《道路防护栏施工质量鉴定表（波形梁钢护栏）》检测"+jdjg.get("总点数")+"点,合格"+jdjg.get("合格点数")+"点");
                        newMap.put("ccname", "*"+jdjg.get("检测项目"));
                        newMap.put("yxps", jdjg.get("规定值或允许偏差"));
                        newMap.put("sheetname", "分部-交安");
                        newMap.put("fbgc", "防护栏");
                        newMaps15.add(newMap);
                    }
                    maps15 = newMaps15;
                    resultlist.addAll(maps15);
                    break;
                case "59交安砼护栏强度.xlsx":
                    List<Map<String, Object>> maps16 = jjgFbgcJtaqssJathlqdService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps16 = new ArrayList<>();
                    List<String> yxps = jjgFbgcJtaqssJathlqdService.selectsjqd(proname,htd);
                    for (String s : yxps) {
                        for (Map<String, Object> jdjg : maps16) {
                            Map<String, Object> newMap = new HashMap<>(jdjg);
                            newMap.put("filename", "详见《交安工程砼护栏强度质量鉴定表》检测"+jdjg.get("总点数")+"点,合格"+jdjg.get("合格点数")+"点");
                            newMap.put("ccname", "*砼护栏强度");
                            newMap.put("yxps", s);
                            newMap.put("sheetname", "分部-交安");
                            newMap.put("fbgc", "防护栏");
                            newMaps16.add(newMap);
                        }
                    }
                    maps16 = newMaps16;
                    resultlist.addAll(maps16);
                    break;
                case "60交安砼护栏断面尺寸.xlsx":
                    List<Map<String, Object>> maps17 = jjgFbgcJtaqssJathldmccService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> newMaps17 = new ArrayList<>();
                    List<Map<String,Object>> yxpsjalist = jjgFbgcJtaqssJathldmccService.selectyxpc(proname,htd);
                    for (Map<String, Object> yxpcmap : yxpsjalist) {
                        for (Map<String, Object> jdjgmap : maps17) {

                            Map<String, Object> newMap = new HashMap<>(jdjgmap);
                            String va="";
                            if (yxpcmap.get("yxwcz").equals(yxpcmap.get("yxwcf"))){
                                va = "±"+yxpcmap.get("yxwcz");
                            }else {
                                va = "+"+yxpcmap.get("yxwcz")+";"+"-"+yxpcmap.get("yxwcf");
                            }
                            newMap.put("filename", "详见《交安砼护栏断面尺寸质量鉴定表》检测"+jdjgmap.get("总点数")+"点,合格"+jdjgmap.get("合格点数")+"点");
                            newMap.put("ccname", "△砼护栏断面尺寸");
                            newMap.put("yxps", va);
                            newMap.put("sheetname", "分部-交安");
                            newMap.put("fbgc", "防护栏");
                            newMaps17.add(newMap);
                        }
                    }
                    maps17 = newMaps17;
                    resultlist.addAll(maps17);
                    break;

                case "12沥青路面压实度.xlsx":
                    //分工作簿
                    List<Map<String, Object>> list = jjgFbgcLmgcLqlmysdService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> resultysd = new ArrayList<>();
                    double sdjcds = 0;
                    double sdhgds = 0;
                    double lmjcds = 0;
                    double lmhgds = 0;
                    String sdgdz="";
                    String lmgdz="";
                    for (Map<String, Object> map : list) {
                        if (map.get("路面类型").toString().contains("隧道右幅") || map.get("路面类型").toString().contains("隧道左幅")){
                            sdjcds += Double.valueOf(map.get("检测点数").toString());
                            sdhgds += Double.valueOf(map.get("合格点数").toString());
                            sdgdz = map.get("规定值").toString();

                        }else {
                        //if (map.get("路面类型").toString().contains("沥青路面压实度右幅") || map.get("路面类型").toString().contains("沥青路面压实度左幅")){
                            lmjcds += Double.valueOf(map.get("检测点数").toString());
                            lmhgds += Double.valueOf(map.get("合格点数").toString());
                            lmgdz = map.get("规定值").toString();

                        }
                    }
                    double gdz1 = Double.valueOf(sdgdz)+1;
                    String gdz2= String.valueOf(gdz1);
                    Map<String, Object> newMap1 = new HashMap<>();
                    newMap1.put("filename","详见《沥青路面压实度质量鉴定表》检测"+sdjcds+"点,合格"+sdhgds+"点");
                    newMap1.put("ccname", "△沥青路面压实度(隧道路面)");
                    newMap1.put("ccname2", "隧道路面");
                    newMap1.put("yxps", gdz2);
                    newMap1.put("sheetname", "分部-路面");
                    newMap1.put("fbgc", "路面面层");
                    newMap1.put("检测点数", sdjcds);
                    newMap1.put("合格点数", sdhgds);
                    newMap1.put("合格率", (sdjcds != 0) ? df.format(sdhgds/sdjcds*100) : "0");

                    double gdz3 = Double.valueOf(lmgdz)+1;
                    String gdz4 = String.valueOf(gdz3);
                    Map<String, Object> newMap2 = new HashMap<>();
                    newMap2.put("filename","详见《沥青路面压实度质量鉴定表》检测"+lmjcds+"点,合格"+lmhgds+"点");
                    newMap2.put("ccname", "△沥青路面压实度(路面面层)");
                    newMap2.put("ccname2", "路面面层");
                    newMap2.put("yxps", gdz4);
                    newMap2.put("sheetname", "分部-路面");
                    newMap2.put("fbgc", "路面面层");
                    newMap2.put("检测点数", lmjcds);
                    newMap2.put("合格点数", lmhgds);
                    newMap2.put("合格率", (lmjcds != 0) ? df.format(lmhgds/lmjcds*100) : "0");
                    resultysd.add(newMap1);
                    resultysd.add(newMap2);
                    resultlist.addAll(resultysd);
                    break;
                case "13路面弯沉(贝克曼梁法).xlsx":
                    List<Map<String, Object>> lmwclist = jjgFbgcLmgcLmwcService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> resultwc = new ArrayList<>();
                    Map<String, Object> wcMap = new HashMap<>();
                    wcMap.put("filename","详见《路面弯沉质量鉴定结果汇总表》检测"+lmwclist.get(0).get("检测单元数")+"点,合格"+lmwclist.get(0).get("合格单元数")+"点");
                    wcMap.put("ccname","△沥青路面弯沉(贝克曼梁法)");
                    wcMap.put("ccname2","贝克曼梁法");
                    wcMap.put("sheetname","分部-路面");
                    wcMap.put("fbgc","路面面层");
                    wcMap.put("合格率",lmwclist.get(0).get("合格率"));
                    wcMap.put("yxps",lmwclist.get(0).get("规定值"));
                    resultwc.add(wcMap);
                    resultlist.addAll(resultwc);
                    break;
                case "13路面弯沉(落锤法).xlsx":
                    List<Map<String, Object>> wclcflist = jjgFbgcLmgcLmwcLcfService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> resultwclcf = new ArrayList<>();
                    Map<String, Object> wclcfMap = new HashMap<>();
                    wclcfMap.put("filename","详见《路面弯沉质量鉴定结果汇总表》检测"+wclcflist.get(0).get("检测单元数")+"点,合格"+wclcflist.get(0).get("合格单元数")+"点");
                    wclcfMap.put("ccname","△沥青路面弯沉(落锤法)");
                    wclcfMap.put("ccname2","落锤法");
                    wclcfMap.put("sheetname","分部-路面");
                    wclcfMap.put("fbgc","路面面层");
                    wclcfMap.put("合格率",wclcflist.get(0).get("合格率"));
                    wclcfMap.put("yxps",wclcflist.get(0).get("规定值"));
                    resultwclcf.add(wclcfMap);
                    resultlist.addAll(resultwclcf);
                    break;
                case "15沥青路面渗水系数.xlsx":
                    //分工作簿
                    List<Map<String, Object>> ssxslist = jjgFbgcLmgcLmssxsService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> resultss = new ArrayList<>();
                    double sdssjcds = 0;
                    double sdsshgds = 0;
                    double lmssjcds = 0;
                    double lmsshgds = 0;
                    String sdssgdz="";
                    String lmssgdz="";
                    for (Map<String, Object> map : ssxslist) {
                        if (map.get("检测项目").toString().contains("沥青路面")){
                            sdssjcds += Double.valueOf(map.get("检测点数").toString());
                            sdsshgds += Double.valueOf(map.get("合格点数").toString());
                            sdssgdz = map.get("规定值").toString();

                        }
                        if (map.get("检测项目").toString().contains("隧道路面")){
                            lmssjcds += Double.valueOf(map.get("检测点数").toString());
                            lmsshgds += Double.valueOf(map.get("合格点数").toString());
                            lmssgdz = map.get("规定值").toString();

                        }
                    }
                    double ssgdz1 = Double.valueOf(sdssgdz);
                    String ssgdz2= String.valueOf(ssgdz1);
                    Map<String, Object> newMapss = new HashMap<>();
                    newMapss.put("filename","详见《沥青路面渗水系数质量鉴定表》检测"+decf.format(sdssjcds)+"点,合格"+decf.format(sdsshgds)+"点");
                    newMapss.put("ccname", "沥青路面渗水系数(隧道路面)");
                    newMapss.put("ccname2", "隧道路面");
                    newMapss.put("yxps", ssgdz2);
                    newMapss.put("sheetname", "分部-路面");
                    newMapss.put("fbgc", "路面面层");
                    newMapss.put("检测点数", decf.format(sdssjcds));
                    newMapss.put("合格点数", decf.format(sdsshgds));
                    newMapss.put("合格率", (sdssjcds != 0) ? df.format(sdsshgds/sdssjcds*100) : "0");

                    double ssgdz3 = Double.valueOf(lmssgdz);
                    String ssgdz4 = String.valueOf(ssgdz3);
                    Map<String, Object> newMapss1 = new HashMap<>();
                    newMapss1.put("filename","详见《沥青路面渗水系数质量鉴定表》检测"+decf.format(lmssjcds)+"点,合格"+decf.format(lmsshgds)+"点");
                    newMapss1.put("ccname", "沥青路面渗水系数(路面面层)");
                    newMapss1.put("ccname2", "路面面层");
                    newMapss1.put("yxps", ssgdz4);
                    newMapss1.put("sheetname", "分部-路面");
                    newMapss1.put("fbgc", "路面面层");
                    newMapss1.put("检测点数", decf.format(lmssjcds));
                    newMapss1.put("合格点数", decf.format(lmsshgds));
                    newMapss1.put("合格率", (lmssjcds != 0) ? df.format(lmsshgds/lmssjcds*100) : "0");
                    resultss.add(newMapss);
                    resultss.add(newMapss1);
                    resultlist.addAll(resultss);
                    break;
                case "16混凝土路面强度.xlsx":
                    List<Map<String, Object>> lmqdlist = jjgFbgcLmgcHntlmqdService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> resultqd = new ArrayList<>();
                    Map<String, Object> qdMap = new HashMap<>();
                    qdMap.put("filename","详见《混凝土路面强度鉴定结果汇总表》检测"+lmqdlist.get(0).get("总点数")+"点,合格"+lmqdlist.get(0).get("合格点数")+"点");
                    qdMap.put("ccname","*砼路面强度");
                    qdMap.put("ccname2","");
                    qdMap.put("sheetname","分部-路面");
                    qdMap.put("fbgc","路面面层");
                    qdMap.put("合格率",lmqdlist.get(0).get("合格率"));
                    qdMap.put("检测点数",lmqdlist.get(0).get("检测点数"));
                    qdMap.put("合格点数",lmqdlist.get(0).get("合格点数"));
                    qdMap.put("yxps",lmqdlist.get(0).get("规定值"));
                    resultqd.add(qdMap);
                    resultlist.addAll(resultqd);
                    break;
                case "17混凝土路面相邻板高差.xlsx":
                    List<Map<String, Object>> list1 = jjgFbgcLmgcTlmxlbgcService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> resultgs = new ArrayList<>();
                    Map<String, Object> gsMap = new HashMap<>();
                    gsMap.put("filename","详见《混凝土路面相邻板高差质量鉴定表》检测"+list1.get(0).get("总点数")+"点,合格"+list1.get(0).get("合格点数")+"点");
                    gsMap.put("ccname","砼路面相邻板高差");
                    gsMap.put("ccname2","");
                    gsMap.put("sheetname","分部-路面");
                    gsMap.put("fbgc","路面面层");
                    gsMap.put("合格率",list1.get(0).get("合格率"));
                    gsMap.put("检测点数",list1.get(0).get("总点数"));
                    gsMap.put("合格点数",list1.get(0).get("合格点数"));
                    gsMap.put("yxps",list1.get(0).get("规定值"));
                    resultgs.add(gsMap);
                    resultlist.addAll(resultgs);
                    break;

                case "20构造深度手工铺沙法.xlsx":
                    List<Map<String, Object>> list2 = jjgFbgcLmgcLmgzsdsgpsfService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> resultgz = new ArrayList<>();
                    Map<String, Object> gzMap = new HashMap<>();
                    gzMap.put("filename","详见《混凝土路面构造深度质量鉴定表》检测"+list2.get(0).get("检测点数")+"点,合格"+list2.get(0).get("合格点数")+"点");
                    gzMap.put("ccname","构造深度");
                    gzMap.put("ccname2","");
                    gzMap.put("sheetname","分部-路面");
                    gzMap.put("fbgc","路面面层");
                    gzMap.put("合格率",list2.get(0).get("合格率"));
                    gzMap.put("检测点数",list2.get(0).get("检测点数"));
                    gzMap.put("合格点数",list2.get(0).get("合格点数"));
                    gzMap.put("yxps",list2.get(0).get("规定值"));
                    resultgz.add(gzMap);
                    resultlist.addAll(resultgz);
                    break;
                case "22沥青路面厚度-钻芯法.xlsx":
                    //匝道
                    List<Map<String, Object>> list3 = jjgFbgcLmgcGslqlmhdzxfService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> listhdzxf = new ArrayList<>();
                    double hdjcds = 0;
                    double hdhgds = 0;
                    String sjz1 = "";
                    String sjz2 = "";

                    double hdjcds2 = 0;
                    double hdhgds2 = 0;
                    String sjz12 = "";
                    String sjz22 = "";
                    for (Map<String, Object> map : list3) {
                        if (map.get("路面类型").toString().contains("隧道")){
                            hdjcds += Double.valueOf(map.get("总厚度检测点数").toString())+Double.valueOf(map.get("上面层厚度检测点数").toString());
                            hdhgds += Double.valueOf(map.get("总厚度合格点数").toString())+Double.valueOf(map.get("上面层厚度合格点数").toString());

                            sjz1 = map.get("总厚度设计值").toString();
                            sjz2 = map.get("上面层设计值").toString();


                        }else if (map.get("路面类型").toString().contains("路面") || map.get("路面类型").toString().contains("路面") ){
                            hdjcds2 += Double.valueOf(map.get("总厚度检测点数").toString())+Double.valueOf(map.get("上面层厚度检测点数").toString());
                            hdhgds2 += Double.valueOf(map.get("总厚度合格点数").toString())+Double.valueOf(map.get("上面层厚度合格点数").toString());

                            sjz12 = map.get("总厚度设计值").toString();
                            sjz22 = map.get("上面层设计值").toString();

                        }
                    }
                    Map map1 = new HashMap();
                    Map map2 = new HashMap();
                    map1.put("检测点数",decf.format(hdjcds));
                    map1.put("合格点数",decf.format(hdhgds));
                    map1.put("yxps",sjz1);
                    map1.put("filename","详见《沥青隧道路面厚度质量鉴定表（钻芯法）》检测"+decf.format(hdjcds)+"点,合格"+decf.format(hdhgds)+"点");
                    map1.put("ccname","△厚度");
                    map1.put("ccname2","隧道路面");
                    map1.put("ccname3","沥青路面");
                    map1.put("ccname4","钻芯法");
                    map1.put("fbgc","路面面层");
                    map1.put("合格率",(hdjcds != 0) ? df.format(hdhgds/hdjcds*100) : "0");

                    map2.put("检测点数",decf.format(hdjcds));
                    map2.put("合格点数",decf.format(hdhgds));
                    map2.put("yxps",sjz2);
                    map2.put("filename","详见《沥青隧道路面厚度质量鉴定表（钻芯法）》检测"+decf.format(hdjcds)+"点,合格"+decf.format(hdhgds)+"点");
                    map2.put("ccname","△厚度");
                    map2.put("ccname2","隧道路面");
                    map2.put("ccname3","沥青路面");
                    map2.put("ccname4","钻芯法");
                    map2.put("fbgc","路面面层");
                    map2.put("合格率",(hdjcds != 0) ? df.format(hdhgds/hdjcds*100) : "0");

                    Map map3 = new HashMap();
                    Map map4 = new HashMap();
                    map3.put("检测点数",decf.format(hdjcds2));
                    map3.put("合格点数",decf.format(hdhgds2));
                    map3.put("yxps",sjz12);
                    map3.put("filename","详见《沥青路面厚度质量鉴定表（钻芯法）》检测"+decf.format(hdjcds2)+"点,合格"+decf.format(hdhgds2)+"点");
                    map3.put("ccname","△厚度");
                    map3.put("ccname2","路面面层");
                    map3.put("ccname3","沥青路面");
                    map3.put("ccname4","钻芯法");
                    map3.put("fbgc","路面面层");
                    map3.put("合格率",(hdjcds2 != 0) ? df.format(hdhgds2/hdjcds2*100) : "0");

                    map4.put("检测点数",decf.format(hdjcds2));
                    map4.put("合格点数",decf.format(hdhgds2));
                    map4.put("yxps",sjz22);
                    map4.put("filename","详见《沥青路面厚度质量鉴定表（钻芯法）》检测"+decf.format(hdjcds2)+"点,合格"+decf.format(hdhgds2)+"点");
                    map4.put("ccname","△厚度");
                    map4.put("ccname2","路面面层");
                    map4.put("ccname3","沥青路面");
                    map4.put("ccname4","钻芯法");
                    map4.put("fbgc","路面面层");
                    map4.put("合格率",(hdjcds2 != 0) ? df.format(hdhgds2/hdjcds2*100) : "0");
                    listhdzxf.add(map1);
                    listhdzxf.add(map2);
                    listhdzxf.add(map3);
                    listhdzxf.add(map4);
                    resultlist.addAll(listhdzxf);
                    break;
                case "23混凝土路面厚度-钻芯法.xlsx":
                    List<Map<String, Object>> list4 = jjgFbgcLmgcHntlmhdzxfService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> resulthdzxf = new ArrayList<>();
                    Map<String, Object> maphdzxf = new HashMap<>();
                    maphdzxf.put("filename","详见《混凝土路面厚度质量鉴定表（钻芯法）》检测"+list4.get(0).get("检测点数")+"点,合格"+list4.get(0).get("合格点数")+"点");
                    maphdzxf.put("ccname", "△厚度");
                    maphdzxf.put("ccname2", "路面面层");
                    maphdzxf.put("ccname3", "混凝土路面");
                    maphdzxf.put("ccname4", "钻芯法");
                    maphdzxf.put("yxps", list4.get(0).get("允许偏差"));
                    maphdzxf.put("sheetname", "分部-路面");
                    maphdzxf.put("fbgc", "路面面层");
                    maphdzxf.put("检测点数", list4.get(0).get("检测点数"));
                    maphdzxf.put("合格点数", list4.get(0).get("合格点数"));
                    maphdzxf.put("合格率", list4.get(0).get("合格率"));
                    resulthdzxf.add(maphdzxf);
                    resultlist.addAll(resulthdzxf);
                    break;
                case "25桥梁下部墩台砼强度.xlsx":
                    List<Map<String, Object>> list11 = jjgFbgcQlgcXbTqdService.lookjg(commonInfoVo);
                    List<Map<String, Object>> qltqdlist = new ArrayList<>();
                    for (Map<String, Object> map : list11) {
                        double z = Double.valueOf(map.get("zds").toString());
                        double h = Double.valueOf(map.get("hgds").toString());
                        Map map5 = new HashMap();
                        map5.put("ccname","墩台混凝土强度(MPa)");
                        map5.put("yxps",map.get("sjqd"));
                        map5.put("filename","详见《桥梁下部墩台砼强度质量鉴定表》检测"+map.get("zds")+"点,合格"+map.get("hgds")+"点");
                        map5.put("sheetname", "分部-"+map.get("qlmc"));
                        map5.put("fbgc", "桥梁下部");
                        map5.put("合格率", (z != 0) ? df.format(h/z*100) : "0");
                        qltqdlist.add(map5);
                    }
                    resultlist.addAll(qltqdlist);
                    break;

                case "26桥梁下部主要结构尺寸.xlsx":
                    List<Map<String, Object>> list12 = jjgFbgcQlgcXbJgccService.lookjg(commonInfoVo);
                    List<Map<String, Object>> qljgcclist = new ArrayList<>();
                    for (Map<String, Object> map : list12) {
                        double z = Double.valueOf(map.get("zds").toString());
                        double h = Double.valueOf(map.get("hgds").toString());
                        Map map5 = new HashMap();
                        map5.put("ccname","主要构件尺寸(mm)");
                        map5.put("yxps",map.get("sjqd"));
                        map5.put("filename","详见《桥梁下部主要结构尺寸质量鉴定表》检测"+map.get("zds")+"点,合格"+map.get("hgds")+"点");
                        map5.put("sheetname", "分部-"+map.get("qlmc"));
                        map5.put("fbgc", "桥梁下部");
                        map5.put("合格率", (z != 0) ? df.format(h/z*100) : "0");
                        qljgcclist.add(map5);
                    }
                    resultlist.addAll(qljgcclist);
                    break;

                case "27桥梁下部保护层厚度.xlsx":
                    List<Map<String, Object>> list13 = jjgFbgcQlgcXbBhchdService.lookjg(commonInfoVo);
                    List<Map<String, Object>> qlbhclist = new ArrayList<>();
                    for (Map<String, Object> map : list13) {
                        double z = Double.valueOf(map.get("zds").toString());
                        double h = Double.valueOf(map.get("hgds").toString());
                        Map map5 = new HashMap();
                        map5.put("ccname","钢筋保护层厚度(mm)");
                        map5.put("yxps",map.get("sjqd"));
                        map5.put("filename","详见《桥梁下部钢筋保护层厚度质量鉴定表》检测"+map.get("zds")+"点,合格"+map.get("hgds")+"点");
                        map5.put("sheetname", "分部-"+map.get("qlmc"));
                        map5.put("fbgc", "桥梁下部");
                        map5.put("合格率", (z != 0) ? df.format(h/z*100) : "0");
                        qlbhclist.add(map5);
                    }
                    resultlist.addAll(qlbhclist);
                    break;

                case "28桥梁下部墩台垂直度.xlsx":
                    List<Map<String, Object>> list14 = jjgFbgcQlgcXbSzdService.lookjg(commonInfoVo);
                    List<Map<String, Object>> qlczdlist = new ArrayList<>();
                    for (Map<String, Object> map : list14) {
                        double z = Double.valueOf(map.get("zds").toString());
                        double h = Double.valueOf(map.get("hgds").toString());
                        Map map5 = new HashMap();
                        map5.put("ccname","墩台垂直度(mm)");
                        map5.put("yxps",map.get("sjqd"));
                        map5.put("filename","详见《桥梁下部墩台垂直度质量鉴定表》检测"+map.get("zds")+"点,合格"+map.get("hgds")+"点");
                        map5.put("sheetname", "分部-"+map.get("qlmc"));
                        map5.put("fbgc", "桥梁下部");
                        map5.put("合格率", (z != 0) ? df.format(h/z*100) : "0");
                        qlczdlist.add(map5);
                    }
                    resultlist.addAll(qlczdlist);
                    break;

                case "29桥梁上部砼强度.xlsx":
                    List<Map<String, Object>> list15 = jjgFbgcQlgcSbTqdService.lookjg(commonInfoVo);
                    List<Map<String, Object>> qlsbqdlist = new ArrayList<>();
                    for (Map<String, Object> map : list15) {
                        double z = Double.valueOf(map.get("zds").toString());
                        double h = Double.valueOf(map.get("hgds").toString());
                        Map map5 = new HashMap();
                        map5.put("ccname","墩台混凝土强度(MPa)");
                        map5.put("yxps",map.get("sjqd"));
                        map5.put("filename","详见《桥梁上部墩台砼强度质量鉴定表》检测"+map.get("zds")+"点,合格"+map.get("hgds")+"点");
                        map5.put("sheetname", "分部-"+map.get("qlmc"));
                        map5.put("fbgc", "桥梁上部");
                        map5.put("合格率", (z != 0) ? df.format(h/z*100) : "0");
                        qlsbqdlist.add(map5);
                    }
                    resultlist.addAll(qlsbqdlist);
                    break;

                case "30桥梁上部主要结构尺寸.xlsx":
                    List<Map<String, Object>> list16 = jjgFbgcQlgcSbJgccService.lookjg(commonInfoVo);
                    List<Map<String, Object>> qlsbjgcclist = new ArrayList<>();
                    for (Map<String, Object> map : list16) {
                        double z = Double.valueOf(map.get("zds").toString());
                        double h = Double.valueOf(map.get("hgds").toString());
                        Map map5 = new HashMap();
                        map5.put("ccname","主要构件尺寸(mm)");
                        map5.put("yxps",map.get("sjqd"));
                        map5.put("filename","详见《桥梁上部主要结构尺寸质量鉴定表》检测"+map.get("zds")+"点,合格"+map.get("hgds")+"点");
                        map5.put("sheetname", "分部-"+map.get("qlmc"));
                        map5.put("fbgc", "桥梁上部");
                        map5.put("合格率", (z != 0) ? df.format(h/z*100) : "0");
                        qlsbjgcclist.add(map5);
                    }
                    resultlist.addAll(qlsbjgcclist);
                    break;
                case "31桥梁上部保护层厚度.xlsx":
                    List<Map<String, Object>> list17 = jjgFbgcQlgcSbBhchdService.lookjg(commonInfoVo);
                    List<Map<String, Object>> qlsbbhclist = new ArrayList<>();
                    for (Map<String, Object> map : list17) {
                        double z = Double.valueOf(map.get("zds").toString());
                        double h = Double.valueOf(map.get("hgds").toString());
                        Map map5 = new HashMap();
                        map5.put("ccname","钢筋保护层厚度(mm)");
                        map5.put("yxps",map.get("sjqd"));
                        map5.put("filename","详见《桥梁上部钢筋保护层厚度质量鉴定表》检测"+map.get("zds")+"点,合格"+map.get("hgds")+"点");
                        map5.put("sheetname", "分部-"+map.get("qlmc"));
                        map5.put("fbgc", "桥梁上部");
                        map5.put("合格率", (z != 0) ? df.format(h/z*100) : "0");
                        qlsbbhclist.add(map5);
                    }
                    resultlist.addAll(qlsbbhclist);
                    break;
                case "24路面横坡.xlsx":
                    //工作簿分路面，桥，隧道
                    List<Map<String, String>> list5 = jjgFbgcLmgcLmhpService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> resulthp = new ArrayList<>();
                    double lmhphgds = 0;
                    double qmhphgds = 0;
                    double sdhphgds = 0;
                    double hnthphgds = 0;

                    double lmhpjcds = 0;
                    double qmhpjcds = 0;
                    double sdhpjcds = 0;
                    double hnthpjcds = 0;
                    String lmyxpc = "";
                    String sdyxpc = "";
                    String qmyxpc = "";
                    String hntyxpc = "";
                    boolean a = false;
                    boolean b = false;
                    boolean c = false;
                    boolean d = false;
                    for (Map<String, String> map : list5) {
                        if (map.get("路面类型").contains("沥青路面") ){
                            lmhphgds += Double.valueOf(map.get("合格点数"));
                            lmhpjcds += Double.valueOf(map.get("检测点数"));
                            lmyxpc = map.get("允许偏差");
                            a = true;

                        } else if (map.get("路面类型").contains("沥青桥面") ){
                            qmhphgds += Double.valueOf(map.get("合格点数"));
                            qmhpjcds += Double.valueOf(map.get("检测点数"));
                            qmyxpc = map.get("允许偏差");
                            b = true;

                        }else if (map.get("路面类型").contains("沥青隧道") ){
                            sdhphgds += Double.valueOf(map.get("合格点数"));
                            sdhpjcds += Double.valueOf(map.get("检测点数"));
                            sdyxpc = map.get("允许偏差");
                            c = true;
                        }else if (map.get("路面类型").contains("混凝土路面") ){
                            hnthphgds += Double.valueOf(map.get("合格点数"));
                            hnthpjcds += Double.valueOf(map.get("检测点数"));
                            hntyxpc = map.get("允许偏差");
                            d = true;
                        }

                    }
                    if (a){
                        Map<String, Object> maphplm = new HashMap<>();
                        maphplm.put("filename","详见《沥青路面横坡质量鉴定表》检测"+decf.format(lmhpjcds)+"点,合格"+decf.format(lmhphgds)+"点");
                        maphplm.put("ccname", "横坡(沥青路面)");
                        maphplm.put("ccname2", "路面面层");
                        maphplm.put("ccname3", "沥青路面");
                        maphplm.put("yxps", lmyxpc);
                        maphplm.put("sheetname", "分部-路面");
                        maphplm.put("fbgc", "路面面层");
                        maphplm.put("检测点数", decf.format(lmhpjcds));
                        maphplm.put("合格点数", decf.format(lmhphgds));
                        maphplm.put("合格率", (lmhpjcds != 0) ? df.format(lmhphgds/lmhpjcds*100) : "0");
                        resulthp.add(maphplm);
                    }
                    if (b){
                        Map<String, Object> maphpqm = new HashMap<>();
                        maphpqm.put("filename","详见《沥青桥面横坡质量鉴定表》检测"+decf.format(qmhpjcds)+"点,合格"+decf.format(qmhphgds)+"点");
                        maphpqm.put("ccname", "横坡(桥面系)");
                        maphpqm.put("ccname2", "桥面系");
                        maphpqm.put("ccname3", "沥青路面");
                        maphpqm.put("yxps", qmyxpc);
                        maphpqm.put("sheetname", "分部-路面");
                        maphpqm.put("fbgc", "路面面层");
                        maphpqm.put("检测点数", decf.format(qmhpjcds));
                        maphpqm.put("合格点数", decf.format(qmhphgds));
                        maphpqm.put("合格率", (qmhpjcds != 0) ? df.format(qmhphgds/qmhpjcds*100) : "0");
                        resulthp.add(maphpqm);

                    }
                    if (c){
                        Map<String, Object> maphpsd = new HashMap<>();
                        maphpsd.put("filename","详见《沥青桥面横坡质量鉴定表》检测"+decf.format(sdhpjcds)+"点,合格"+decf.format(sdhphgds)+"点");
                        maphpsd.put("ccname", "横坡(隧道路面)");
                        maphpsd.put("ccname2", "隧道路面");
                        maphpsd.put("ccname3", "沥青路面");
                        maphpsd.put("yxps", sdyxpc);
                        maphpsd.put("sheetname", "分部-路面");
                        maphpsd.put("fbgc", "路面面层");
                        maphpsd.put("检测点数", decf.format(sdhpjcds));
                        maphpsd.put("合格点数", decf.format(sdhphgds));
                        maphpsd.put("合格率", (sdhpjcds != 0) ? df.format(sdhphgds/sdhpjcds*100) : "0");
                        resulthp.add(maphpsd);


                    }
                    if (d){
                        Map<String, Object> maphphnt = new HashMap<>();
                        maphphnt.put("filename","详见《混凝土路面横坡质量鉴定表》检测"+decf.format(hnthpjcds)+"点,合格"+decf.format(hnthphgds)+"点");
                        maphphnt.put("ccname", "横坡(水泥路面)");
                        maphphnt.put("ccname2", "路面面层");
                        maphphnt.put("ccname3", "水泥混凝土");
                        maphphnt.put("yxps", hntyxpc);
                        maphphnt.put("sheetname", "分部-路面");
                        maphphnt.put("fbgc", "路面面层");
                        maphphnt.put("检测点数", decf.format(hnthpjcds));
                        maphphnt.put("合格点数", decf.format(hnthphgds));
                        maphphnt.put("合格率", (hnthpjcds != 0) ? df.format(hnthphgds/hnthpjcds*100) : "0");
                        resulthp.add(maphphnt);

                    }
                    resultlist.addAll(resulthp);
                    break;
                case "20路面构造深度.xlsx":
                    List<Map<String, Object>> list6 = jjgZdhGzsdService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> resultgzsd = new ArrayList<>();
                    double lmzds = 0;
                    double lmhgs = 0;
                    String lmsjz = "";
                    double sdzds = 0;
                    double sdhgd = 0;
                    String sdsjz = "";
                    double qzds = 0;
                    double qhgds = 0;
                    String qsjz = "";
                    boolean aa = false;
                    boolean bb = false;
                    boolean cc = false;
                    for (Map<String, Object> map : list6) {
                        if (map.get("路面类型").toString().contains("路面")){
                            lmzds += Double.valueOf(map.get("总点数").toString());
                            lmhgs += Double.valueOf(map.get("合格点数").toString());
                            lmsjz = map.get("设计值").toString();
                            aa = true;
                        }
                        if (map.get("路面类型").toString().contains("隧道")){
                            sdzds += Double.valueOf(map.get("总点数").toString());
                            sdhgd += Double.valueOf(map.get("合格点数").toString());
                            sdsjz = map.get("设计值").toString();
                            bb =true;
                        }
                        if (map.get("路面类型").toString().contains("桥")){
                            qzds += Double.valueOf(map.get("总点数").toString());
                            qhgds += Double.valueOf(map.get("合格点数").toString());
                            qsjz = map.get("设计值").toString();
                            cc = true;
                        }
                    }
                    if (aa){
                        Map<String, Object> map = new HashMap<>();
                        map.put("filename","详见《沥青路面构造深度质量鉴定表》检测"+decf.format(lmzds)+"点,合格"+decf.format(lmhgs)+"点");
                        map.put("ccname", "构造深度");
                        map.put("ccname2", "路面面层");
                        map.put("ccname3", "沥青路面");
                        map.put("yxps", lmsjz);
                        map.put("sheetname", "分部-路面");
                        map.put("fbgc", "路面面层");
                        map.put("检测点数", decf.format(lmzds));
                        map.put("合格点数", decf.format(lmhgs));
                        map.put("合格率", (lmhgs != 0) ? df.format(lmhgs/lmzds*100) : "0");
                        resultgzsd.add(map);
                    }
                    if (bb){
                        Map<String, Object> map = new HashMap<>();
                        map.put("filename","详见《隧道路面构造深度质量鉴定表》检测"+decf.format(sdzds)+"点,合格"+decf.format(sdhgd)+"点");
                        map.put("ccname", "构造深度");
                        map.put("ccname2", "隧道路面");
                        map.put("ccname3", "沥青路面");
                        map.put("yxps", sdsjz);
                        map.put("sheetname", "分部-路面");
                        map.put("fbgc", "路面面层");
                        map.put("检测点数", decf.format(sdzds));
                        map.put("合格点数", decf.format(sdhgd));
                        map.put("合格率", (sdhgd != 0) ? df.format(sdhgd/sdzds*100) : "0");
                        resultgzsd.add(map);
                    }
                    if (cc){
                        Map<String, Object> map = new HashMap<>();
                        map.put("filename","详见《桥面系构造深度质量鉴定表》检测"+decf.format(qzds)+"点,合格"+decf.format(qhgds)+"点");
                        map.put("ccname", "构造深度");
                        map.put("ccname2", "桥面系");
                        map.put("ccname3", "沥青路面");
                        map.put("yxps", qsjz);
                        map.put("sheetname", "分部-路面");
                        map.put("fbgc", "路面面层");
                        map.put("检测点数", decf.format(qzds));
                        map.put("合格点数", decf.format(qhgds));
                        map.put("合格率", (qhgds != 0) ? df.format(qhgds/qzds*100) : "0");
                        resultgzsd.add(map);
                    }
                    resultlist.addAll(resultgzsd);
                    break;
                case "19路面摩擦系数.xlsx":
                    //分工作簿
                    List<Map<String, Object>> list7 = jjgZdhMcxsService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> resultmcxs = new ArrayList<>();
                    double lmzds2 = 0;
                    double lmhgs2 = 0;
                    String lmsjz2 = "";
                    double sdzds2 = 0;
                    double sdhgd2 = 0;
                    String sdsjz2 = "";
                    double qzds2 = 0;
                    double qhgds2 = 0;
                    String qsjz2 = "";
                    boolean aaa = false;
                    boolean bbb = false;
                    boolean ccc = false;
                    for (Map<String, Object> map : list7) {
                        if (map.get("路面类型").toString().contains("路面")){
                            lmzds2 += Double.valueOf(map.get("总点数").toString());
                            lmhgs2 += Double.valueOf(map.get("合格点数").toString());
                            lmsjz2 = map.get("设计值").toString();
                            aaa = true;
                        }
                        if (map.get("路面类型").toString().contains("隧道")){
                            sdzds2 += Double.valueOf(map.get("总点数").toString());
                            sdhgd2 += Double.valueOf(map.get("合格点数").toString());
                            sdsjz2 = map.get("设计值").toString();
                            bbb =true;
                        }
                        if (map.get("路面类型").toString().contains("桥")){
                            qzds2 += Double.valueOf(map.get("总点数").toString());
                            qhgds2 += Double.valueOf(map.get("合格点数").toString());
                            qsjz2 = map.get("设计值").toString();
                            ccc = true;
                        }
                    }
                    if (aaa){
                        Map<String, Object> map = new HashMap<>();
                        map.put("filename","详见《沥青路面摩擦系数质量鉴定表》检测"+decf.format(lmzds2)+"点,合格"+decf.format(lmhgs2)+"点");
                        map.put("ccname", "摩擦系数");
                        map.put("ccname2", "路面面层");
                        map.put("ccname3", "沥青路面");
                        map.put("yxps", lmsjz2);
                        map.put("sheetname", "分部-路面");
                        map.put("fbgc", "路面面层");
                        map.put("检测点数", decf.format(lmzds2));
                        map.put("合格点数", decf.format(lmhgs2));
                        map.put("合格率", (lmhgs2 != 0) ? df.format(lmhgs2/lmzds2*100) : "0");
                        resultmcxs.add(map);
                    }
                    if (bbb){
                        Map<String, Object> map = new HashMap<>();
                        map.put("filename","详见《隧道路面摩擦系数质量鉴定表》检测"+decf.format(sdzds2)+"点,合格"+decf.format(sdhgd2)+"点");
                        map.put("ccname", "摩擦系数");
                        map.put("ccname2", "隧道路面");
                        map.put("ccname3", "沥青路面");
                        map.put("yxps", sdsjz2);
                        map.put("sheetname", "分部-路面");
                        map.put("fbgc", "路面面层");
                        map.put("检测点数", decf.format(sdzds2));
                        map.put("合格点数", decf.format(sdhgd2));
                        map.put("合格率", (sdhgd2 != 0) ? df.format(sdhgd2/sdzds2*100) : "0");
                        resultmcxs.add(map);
                    }
                    if (ccc){
                        Map<String, Object> map = new HashMap<>();
                        map.put("filename","详见《桥面系摩擦系数质量鉴定表》检测"+decf.format(qzds2)+"点,合格"+decf.format(qhgds2)+"点");
                        map.put("ccname", "摩擦系数");
                        map.put("ccname2", "桥面系");
                        map.put("ccname3", "沥青路面");
                        map.put("yxps", qsjz2);
                        map.put("sheetname", "分部-路面");
                        map.put("fbgc", "路面面层");
                        map.put("检测点数", decf.format(qzds2));
                        map.put("合格点数", decf.format(qhgds2));
                        map.put("合格率", (qhgds2 != 0) ? df.format(qhgds2/qzds2*100) : "0");
                        resultmcxs.add(map);
                    }
                    resultlist.addAll(resultmcxs);
                    break;
                case "18路面平整度.xlsx":
                    List<Map<String, Object>> list8 = jjgZdhPzdService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> resultpzd = new ArrayList<>();
                    double lmzds3 = 0;
                    double lmhgs3 = 0;
                    String lmsjz3 = "";
                    double sdzds3 = 0;
                    double sdhgd3 = 0;
                    String sdsjz3 = "";
                    double qzds3 = 0;
                    double qhgds3 = 0;
                    String qsjz3 = "";
                    boolean aaaa = false;
                    boolean bbbb = false;
                    boolean cccc = false;
                    for (Map<String, Object> map : list8) {
                        if (map.get("路面类型").toString().contains("路面")){
                            lmzds3 += Double.valueOf(map.get("总点数").toString());
                            lmhgs3 += Double.valueOf(map.get("合格点数").toString());
                            lmsjz3 = map.get("设计值").toString();
                            aaaa = true;
                        }
                        if (map.get("路面类型").toString().contains("隧道")){
                            sdzds3 += Double.valueOf(map.get("总点数").toString());
                            sdhgd3 += Double.valueOf(map.get("合格点数").toString());
                            sdsjz3 = map.get("设计值").toString();
                            bbbb =true;
                        }
                        if (map.get("路面类型").toString().contains("桥")){
                            qzds3 += Double.valueOf(map.get("总点数").toString());
                            qhgds3 += Double.valueOf(map.get("合格点数").toString());
                            qsjz3 = map.get("设计值").toString();
                            cccc = true;
                        }
                    }
                    if (aaaa){
                        Map<String, Object> map = new HashMap<>();
                        map.put("filename","详见《沥青路面平整度质量鉴定表》检测"+decf.format(lmzds3)+"点,合格"+decf.format(lmhgs3)+"点");
                        map.put("ccname", "平整度");
                        map.put("ccname2", "路面面层");
                        map.put("ccname3", "沥青路面");
                        map.put("yxps", lmsjz3);
                        map.put("sheetname", "分部-路面");
                        map.put("fbgc", "路面面层");
                        map.put("检测点数", decf.format(lmzds3));
                        map.put("合格点数", decf.format(lmhgs3));
                        map.put("合格率", (lmhgs3 != 0) ? df.format(lmhgs3/lmzds3*100) : "0");
                        resultpzd.add(map);
                    }
                    if (bbbb){
                        Map<String, Object> map = new HashMap<>();
                        map.put("filename","详见《隧道路面平整度质量鉴定表》检测"+decf.format(sdzds3)+"点,合格"+decf.format(sdhgd3)+"点");
                        map.put("ccname", "平整度");
                        map.put("ccname2", "隧道路面");
                        map.put("ccname3", "沥青路面");
                        map.put("yxps", sdsjz3);
                        map.put("sheetname", "分部-路面");
                        map.put("fbgc", "路面面层");
                        map.put("检测点数", decf.format(sdzds3));
                        map.put("合格点数", decf.format(sdhgd3));
                        map.put("合格率", (sdhgd3 != 0) ? df.format(sdhgd3/sdzds3*100) : "0");
                        resultpzd.add(map);
                    }
                    if (cccc){
                        Map<String, Object> map = new HashMap<>();
                        map.put("filename","详见《桥面系平整度质量鉴定表》检测"+decf.format(qzds3)+"点,合格"+decf.format(qhgds3)+"点");
                        map.put("ccname", "平整度");
                        map.put("ccname2", "桥面系");
                        map.put("ccname3", "沥青路面");
                        map.put("yxps", qsjz3);
                        map.put("sheetname", "分部-路面");
                        map.put("fbgc", "路面面层");
                        map.put("检测点数", decf.format(qzds3));
                        map.put("合格点数", decf.format(qhgds3));
                        map.put("合格率", (qhgds3 != 0) ? df.format(qhgds3/qzds3*100) : "0");
                        resultpzd.add(map);
                    }
                    resultlist.addAll(resultpzd);

                    break;
                case "16路面雷达厚度.xlsx":
                    List<Map<String, Object>> list9 = jjgZdhLdhdService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> resultldhd = new ArrayList<>();
                    double lmzds4 = 0;
                    double lmhgs4 = 0;
                    String lmsjz4 = "";
                    double sdzds4 = 0;
                    double sdhgd4 = 0;
                    String sdsjz4 = "";
                    double qzds4 = 0;
                    double qhgds4 = 0;
                    String qsjz4 = "";
                    boolean aaaaaa = false;
                    boolean bbbbbb = false;
                    boolean cccccc = false;
                    for (Map<String, Object> map : list9) {
                        if (map.get("路面类型").toString().contains("右幅") || map.get("路面类型").toString().contains("左幅")){
                            lmzds4 += Double.valueOf(map.get("总点数").toString());
                            lmhgs4 += Double.valueOf(map.get("合格点数").toString());
                            lmsjz4 = map.get("设计值").toString();
                            aaaaaa = true;
                        }
                        if (map.get("路面类型").toString().contains("隧道")){
                            sdzds4 += Double.valueOf(map.get("总点数").toString());
                            sdhgd4 += Double.valueOf(map.get("合格点数").toString());
                            sdsjz4 = map.get("设计值").toString();
                            bbbbbb =true;
                        }
                        if (map.get("路面类型").toString().contains("桥")){
                            qzds4 += Double.valueOf(map.get("总点数").toString());
                            qhgds4 += Double.valueOf(map.get("合格点数").toString());
                            qsjz4 = map.get("设计值").toString();
                            cccccc = true;
                        }
                    }
                    if (aaaaaa){
                        Map<String, Object> map = new HashMap<>();
                        map.put("filename","详见《路面厚度质量鉴定表（雷达法）》检测"+decf.format(lmzds4)+"点,合格"+decf.format(lmhgs4)+"点");
                        map.put("ccname", "△雷达厚度");
                        map.put("ccname2", "路面面层");
                        map.put("ccname3", "沥青路面");
                        map.put("yxps", lmsjz4);
                        map.put("sheetname", "分部-路面");
                        map.put("fbgc", "路面面层");
                        map.put("检测点数", decf.format(lmzds4));
                        map.put("合格点数", decf.format(lmhgs4));
                        map.put("合格率", (lmhgs4 != 0) ? df.format(lmhgs4/lmzds4*100) : "0");
                        resultldhd.add(map);
                    }
                    if (bbbbbb){
                        Map<String, Object> map = new HashMap<>();
                        map.put("filename","详见《隧道路面厚度质量鉴定表（雷达法）》检测"+decf.format(sdzds4)+"点,合格"+decf.format(sdhgd4)+"点");
                        map.put("ccname", "△雷达厚度");
                        map.put("ccname2", "隧道路面");
                        map.put("ccname3", "沥青路面");
                        map.put("yxps", sdsjz4);
                        map.put("sheetname", "分部-路面");
                        map.put("fbgc", "路面面层");
                        map.put("检测点数", decf.format(sdzds4));
                        map.put("合格点数", decf.format(sdhgd4));
                        map.put("合格率", (sdhgd4 != 0) ? df.format(sdhgd4/sdzds4*100) : "0");
                        resultldhd.add(map);
                    }
                    if (cccccc){
                        Map<String, Object> map = new HashMap<>();
                        map.put("filename","详见《路面厚度质量鉴定表（雷达法）》检测"+decf.format(qzds4)+"点,合格"+decf.format(qhgds4)+"点");
                        map.put("ccname", "△雷达厚度");
                        map.put("ccname2", "桥面系");
                        map.put("ccname3", "沥青路面");
                        map.put("yxps", qsjz4);
                        map.put("sheetname", "分部-路面");
                        map.put("fbgc", "路面面层");
                        map.put("检测点数", decf.format(qzds4));
                        map.put("合格点数", decf.format(qhgds4));
                        map.put("合格率", (qhgds4 != 0) ? df.format(qhgds4/qzds4*100) : "0");
                        resultldhd.add(map);
                    }
                    resultlist.addAll(resultldhd);
                    break;
                case "14路面车辙.xlsx":
                    List<Map<String, Object>> list10 = jjgZdhCzService.lookJdbjg(commonInfoVo);
                    List<Map<String, Object>> resultcz = new ArrayList<>();
                    double lmzds5 = 0;
                    double lmhgs5 = 0;
                    String lmsjz5 = "";
                    double sdzds5 = 0;
                    double sdhgd5 = 0;
                    String sdsjz5 = "";
                    double qzds5 = 0;
                    double qhgds5 = 0;
                    String qsjz5 = "";
                    boolean aaaaa = false;
                    boolean bbbbb = false;
                    boolean ccccc = false;
                    for (Map<String, Object> map : list10) {
                        if (map.get("路面类型").toString().contains("路面")){
                            lmzds5 += Double.valueOf(map.get("总点数").toString());
                            lmhgs5 += Double.valueOf(map.get("合格点数").toString());
                            lmsjz5 = map.get("设计值").toString();
                            aaaaa = true;
                        }
                        if (map.get("路面类型").toString().contains("隧道")){
                            sdzds5 += Double.valueOf(map.get("总点数").toString());
                            sdhgd5 += Double.valueOf(map.get("合格点数").toString());
                            sdsjz5 = map.get("设计值").toString();
                            bbbbb =true;
                        }
                        if (map.get("路面类型").toString().contains("桥")){
                            qzds5 += Double.valueOf(map.get("总点数").toString());
                            qhgds5 += Double.valueOf(map.get("合格点数").toString());
                            qsjz5 = map.get("设计值").toString();
                            ccccc = true;
                        }
                    }
                    if (aaaaa){
                        Map<String, Object> map = new HashMap<>();
                        map.put("filename","详见《沥青路面车辙质量鉴定表》检测"+decf.format(lmzds5)+"点,合格"+decf.format(lmhgs5)+"点");
                        map.put("ccname", "车辙");
                        map.put("ccname2", "路面面层");
                        map.put("ccname3", "沥青路面");
                        map.put("yxps", lmsjz5);
                        map.put("sheetname", "分部-路面");
                        map.put("fbgc", "路面面层");
                        map.put("检测点数", decf.format(lmzds5));
                        map.put("合格点数", decf.format(lmhgs5));
                        map.put("合格率", (lmhgs5 != 0) ? df.format(lmhgs5/lmzds5*100) : "0");
                        resultcz.add(map);
                    }
                    if (bbbbb){
                        Map<String, Object> map = new HashMap<>();
                        map.put("filename","详见《隧道路面车辙质量鉴定表》检测"+decf.format(sdzds5)+"点,合格"+decf.format(sdhgd5)+"点");
                        map.put("ccname", "车辙");
                        map.put("ccname2", "隧道路面");
                        map.put("ccname3", "沥青路面");
                        map.put("yxps", sdsjz5);
                        map.put("sheetname", "分部-路面");
                        map.put("fbgc", "路面面层");
                        map.put("检测点数", decf.format(sdzds5));
                        map.put("合格点数", decf.format(sdhgd5));
                        map.put("合格率", (sdhgd5 != 0) ? df.format(sdhgd5/sdzds5*100) : "0");
                        resultcz.add(map);
                    }
                    if (ccccc){
                        Map<String, Object> map = new HashMap<>();
                        map.put("filename","详见《桥面系车辙质量鉴定表》检测"+decf.format(qzds5)+"点,合格"+decf.format(qhgds5)+"点");
                        map.put("ccname", "车辙");
                        map.put("ccname2", "桥面系");
                        map.put("ccname3", "沥青路面");
                        map.put("yxps", qsjz5);
                        map.put("sheetname", "分部-路面");
                        map.put("fbgc", "路面面层");
                        map.put("检测点数", decf.format(qzds5));
                        map.put("合格点数", decf.format(qhgds5));
                        map.put("合格率", (qhgds5 != 0) ? df.format(qhgds5/qzds5*100) : "0");
                        resultcz.add(map);
                    }
                    resultlist.addAll(resultcz);
                    break;
                case "38隧道衬砌砼强度.xlsx":
                    List<Map<String, Object>> list18 = jjgFbgcSdgcCqtqdService.lookjg(commonInfoVo);
                    List<Map<String, Object>> sdqdlist = new ArrayList<>();
                    for (Map<String, Object> map : list18) {
                        double z = Double.valueOf(map.get("zds").toString());
                        double h = Double.valueOf(map.get("hgds").toString());
                        Map map5 = new HashMap();
                        map5.put("ccname","*衬砌强度");
                        map5.put("yxps",map.get("sjqd"));
                        map5.put("filename","详见《隧道衬砌强度质量鉴定表》检测"+map.get("zds")+"点,合格"+map.get("hgds")+"点");
                        map5.put("sheetname", "分部-"+map.get("qlmc"));
                        map5.put("fbgc", "衬砌");
                        map5.put("合格率", (z != 0) ? df.format(h/z*100) : "0");
                        sdqdlist.add(map5);
                    }
                    resultlist.addAll(sdqdlist);
                    break;
                case "40隧道大面平整度.xlsx":
                    List<Map<String, Object>> list19 = jjgFbgcSdgcDmpzdService.lookjg(commonInfoVo);
                    List<Map<String, Object>> dmlist = new ArrayList<>();
                    for (Map<String, Object> map : list19) {
                        double z = Double.valueOf(map.get("zds").toString());
                        double h = Double.valueOf(map.get("hgds").toString());
                        Map map5 = new HashMap();
                        map5.put("ccname","大面平整度");
                        map5.put("yxps",map.get("sjqd"));
                        map5.put("filename","详见《隧道大面平整度质量鉴定表》检测"+map.get("zds")+"点,合格"+map.get("hgds")+"点");
                        map5.put("sheetname", "分部-"+map.get("qlmc"));
                        map5.put("fbgc", "衬砌");
                        map5.put("合格率", (z != 0) ? df.format(h/z*100) : "0");
                        dmlist.add(map5);
                    }
                    resultlist.addAll(dmlist);
                    break;
                case "41隧道总体宽度.xlsx":
                    List<Map<String, Object>> list20 = jjgFbgcSdgcZtkdService.lookjg(commonInfoVo);
                    List<Map<String, Object>> kdlist = new ArrayList<>();
                    for (Map<String, Object> map : list20) {
                        double z = Double.valueOf(map.get("zds").toString());
                        double h = Double.valueOf(map.get("hgds").toString());
                        Map map5 = new HashMap();
                        map5.put("ccname","宽度");
                        map5.put("yxps",map.get("sjqd"));
                        map5.put("filename","详见《隧道总体宽度质量鉴定表》检测"+map.get("zds")+"点,合格"+map.get("hgds")+"点");
                        map5.put("sheetname", "分部-"+map.get("qlmc"));
                        map5.put("fbgc", "总体");
                        map5.put("合格率", (z != 0) ? df.format(h/z*100) : "0");
                        kdlist.add(map5);
                    }
                    resultlist.addAll(kdlist);
                    break;
                default:
                    // 默认操作
                    break;
            }

            if (value.contains("34桥面平整度3米直尺法-")){
                List<Map<String, Object>> list = jjgFbgcQlgcQmpzdService.lookjg(commonInfoVo,value);
                List<Map<String, Object>> pzdlist = new ArrayList<>();
                for (Map<String, Object> map : list) {
                    Map map5 = new HashMap();
                    map5.put("ccname","平整度");
                    map5.put("ccname1","沥青路面");
                    map5.put("yxps",map.get("yxpc"));
                    map5.put("filename","详见《桥面系平整度质量鉴定表》检测"+map.get("检测点数")+"点,合格"+map.get("合格点数")+"点");
                    map5.put("sheetname", "分部-"+map.get("检测项目"));
                    map5.put("fbgc", "桥面系");
                    map5.put("合格率", map.get("合格率"));
                    pzdlist.add(map5);
                }
                resultlist.addAll(pzdlist);

            }else if (value.contains("35桥面横坡-")){
                List<Map<String, Object>> list = jjgFbgcQlgcQmhpService.lookJdb(commonInfoVo,value);
                double zds = 0;
                double hgds = 0;
                for (Map<String, Object> map : list) {
                    zds += Double.valueOf(map.get("总点数").toString());
                    hgds += Double.valueOf(map.get("合格点数").toString());
                }
                List<Map<String, Object>> qhplist = new ArrayList<>();
                Map map5 = new HashMap();
                map5.put("ccname","横坡");
                map5.put("ccname1","砼路面");
                map5.put("yxps",list.get(0).get("yxpc"));
                map5.put("filename","详见《桥面系横坡质量鉴定表》检测"+decf.format(zds)+"点,合格"+decf.format(hgds)+"点");
                map5.put("sheetname", "分部-"+list.get(0).get("检测项目"));
                map5.put("fbgc", "桥面系");
                map5.put("合格率", (zds != 0) ? df.format(hgds/zds*100) : "0");
                qhplist.add(map5);
                resultlist.addAll(qhplist);

            }else if (value.contains("37构造深度手工铺沙法-")){
                List<Map<String, Object>> list = jjgFbgcQlgcQmgzsdService.lookjg(commonInfoVo,value);
                List<Map<String, Object>> qgzsdlist = new ArrayList<>();
                for (Map<String, Object> map : list) {
                    Map map5 = new HashMap();
                    map5.put("ccname","抗滑");
                    map5.put("ccname1","砼路面");
                    map5.put("ccname2","构造深度");
                    map5.put("yxps",map.get("yxpc"));
                    map5.put("filename","详见《桥面系构造深度质量鉴定表》检测"+map.get("检测点数")+"点,合格"+map.get("合格点数")+"点");
                    map5.put("sheetname", "分部-"+map.get("检测项目"));
                    map5.put("fbgc", "桥面系");
                    map5.put("合格率", map.get("合格率"));
                    qgzsdlist.add(map5);
                }
                resultlist.addAll(qgzsdlist);

            }else if (value.contains("39隧道衬砌厚度-")){
                List<Map<String, Object>> list = jjgFbgcSdgcCqhdService.lookjg(commonInfoVo,value);
                List<Map<String, Object>> cqhdlist = new ArrayList<>();
                //查询设计厚度
                String s = StringUtils.substringBetween(value, "-", ".");
                List<Map<String, Object>> sjhd = jjgFbgcSdgcCqhdService.selectsjhd(commonInfoVo,s);
                for (Map<String, Object> map : sjhd) {
                    for (Map<String, Object> stringObjectMap : list) {
                        //if (map.get("sdmc").equals(stringObjectMap.get("检测项目"))){
                            Map map5 = new HashMap();
                            map5.put("ccname","衬砌厚度");
                            map5.put("ccname1","衬砌厚度（mm）");
                            map5.put("ccname2","衬砌厚度（mm）");
                            map5.put("yxps",map.get("sjhd"));
                            map5.put("filename","详见《隧道衬砌厚度质量鉴定表》检测"+stringObjectMap.get("检测总点数")+"点,合格"+stringObjectMap.get("合格点数")+"点");
                            map5.put("sheetname", "分部-"+stringObjectMap.get("检测项目"));
                            map5.put("fbgc", "衬砌");
                            map5.put("合格率", stringObjectMap.get("合格率"));
                            cqhdlist.add(map5);
                        //}
                    }
                }
                resultlist.addAll(cqhdlist);
            }else if (value.contains("43隧道沥青路面压实度-")){
                List<Map<String, Object>> list = jjgFbgcSdgcSdlqlmysdService.lookjg(commonInfoVo,value);
                List<Map<String, Object>> sdysdlist = new ArrayList<>();
                double sdhgds = 0;
                double sdjcds = 0;
                double lmhgds = 0;
                double lmjcds = 0;
                boolean a = false;
                boolean b = false;
                String jcxm1 = "";
                String jcxm2 = "";
                String sdgdz = "";
                String lmgdz = "";
                for (Map<String, Object> map : list) {
                    if (map.get("路面类型").toString().contains("隧道")){
                        sdhgds += Double.valueOf(map.get("合格点数").toString());
                        sdjcds += Double.valueOf(map.get("检测点数").toString());
                        sdgdz = map.get("规定值").toString();
                        jcxm1 = map.get("检测项目").toString();
                        a = true;
                    }else {
                        //沥青路面
                        lmhgds += Double.valueOf(map.get("合格点数").toString());
                        lmjcds += Double.valueOf(map.get("检测点数").toString());
                        lmgdz = map.get("规定值").toString();
                        jcxm2 = map.get("检测项目").toString();
                        b = true;
                    }
                }
                if (a){
                    double gdz1 = Double.valueOf(sdgdz)+1;
                    String gdz2= String.valueOf(gdz1);
                    Map map = new HashMap();
                    map.put("ccname","沥青路面压实度(%");
                    map.put("ccname1","隧道路面");
                    map.put("ccname2","");
                    map.put("yxps",gdz2);
                    map.put("filename","详见《沥青路面压实度质量鉴定表》检测"+decf.format(sdjcds)+"点,合格"+decf.format(sdhgds)+"点");
                    map.put("sheetname", "分部-"+jcxm1);
                    map.put("fbgc", "路面面层");
                    map.put("合格率", (sdjcds != 0) ? df.format(sdhgds/sdjcds*100) : "0");
                    sdysdlist.add(map);
                }
                if (b){
                    double gdz3 = Double.valueOf(lmgdz)+1;
                    String gdz4 = String.valueOf(gdz3);
                    Map map = new HashMap();
                    map.put("ccname","沥青路面压实度(%");
                    map.put("ccname1","路面面层");
                    map.put("ccname2","");
                    map.put("yxps",gdz4);
                    map.put("filename","详见《沥青路面压实度质量鉴定表》检测"+decf.format(lmjcds)+"点,合格"+decf.format(lmhgds)+"点");
                    map.put("sheetname", "分部-"+jcxm2);
                    map.put("fbgc", "路面面层");
                    map.put("合格率", (lmjcds != 0) ? df.format(lmhgds/lmjcds*100) : "0");
                    sdysdlist.add(map);
                }
                resultlist.addAll(sdysdlist);

            }else if (value.contains("46隧道沥青路面渗水系数-")){
                List<Map<String, Object>> list = jjgFbgcSdgcLmssxsService.lookjg(commonInfoVo,value);
                List<Map<String, Object>> ssxslist = new ArrayList<>();
                String jcxm = StringUtils.substringBetween(value, "-", ".");
                double szds = 0;
                double shgds = 0;
                String sgdz="";
                String jcxm1="";
                boolean a= false;

                double mzds = 0;
                double mhgds = 0;
                String mgdz="";
                String jcxm2="";
                boolean b= false;
                for (Map<String, Object> map : list) {
                    if (map.get("路面类型").toString().contains("隧道")){
                        szds += Double.valueOf(map.get("检测点数").toString());
                        shgds += Double.valueOf(map.get("合格点数").toString());
                        sgdz = map.get("规定值").toString();
                        jcxm1 = map.get("检测项目").toString();
                        a = true;

                    }else {
                        mzds += Double.valueOf(map.get("检测点数").toString());
                        mhgds += Double.valueOf(map.get("合格点数").toString());
                        mgdz = map.get("规定值").toString();
                        jcxm2 = map.get("检测项目").toString();
                        b= true;
                    }
                }
                if (a){
                    Map map = new HashMap();
                    map.put("ccname","沥青路面渗水系数");
                    map.put("ccname1","隧道路面");
                    map.put("ccname2","");
                    map.put("yxps",sgdz);
                    map.put("filename","详见《路面渗水系数质量鉴定表》检测"+decf.format(szds)+"点,合格"+decf.format(shgds)+"点");
                    map.put("sheetname", "分部-"+jcxm1);
                    map.put("fbgc", "路面面层");
                    map.put("合格率", (szds != 0) ? df.format(shgds/szds*100) : "0");
                    ssxslist.add(map);
                }
                if (b){
                    Map map = new HashMap();
                    map.put("ccname","沥青路面渗水系数");
                    map.put("ccname1","路面面层");
                    map.put("ccname2","");
                    map.put("yxps",mgdz);
                    map.put("filename","详见《路面渗水系数质量鉴定表》检测"+decf.format(mzds)+"点,合格"+decf.format(mhgds)+"点");
                    map.put("sheetname", "分部-"+jcxm2);
                    map.put("fbgc", "路面面层");
                    map.put("合格率", (mzds != 0) ? df.format(mhgds/mzds*100) : "0");
                    ssxslist.add(map);
                }
                resultlist.addAll(ssxslist);

            }else if (value.contains("47隧道混凝土路面强度-")){
                List<Map<String, Object>> list = jjgFbgcSdgcHntlmqdService.lookjg(commonInfoVo,value);
                List<Map<String, Object>> relist = new ArrayList<>();
                String s = StringUtils.substringBetween(value, "-", ".");
                Map map = new HashMap();
                map.put("ccname","混凝土路面强度");
                map.put("ccname1","路面面层");
                map.put("ccname2","");
                map.put("yxps",list.get(0).get("sjqd"));
                map.put("filename","详见《混凝土路面弯拉强度鉴定表》检测"+list.get(0).get("检测点数")+"点,合格"+list.get(0).get("合格点数")+"点");
                map.put("sheetname", "分部-"+s);
                map.put("fbgc", "路面面层");
                map.put("合格率", list.get(0).get("合格率"));
                relist.add(map);
                resultlist.addAll(relist);

            }else if (value.contains("48隧道混凝土路面相邻板高差-")){
                List<Map<String, Object>> list = jjgFbgcSdgcTlmxlbgcService.lookjg(commonInfoVo,value);
                List<Map<String, Object>> relist = new ArrayList<>();
                String s = StringUtils.substringBetween(value, "-", ".");
                Map map = new HashMap();
                map.put("ccname","混凝土路面相邻板高差");
                map.put("ccname1","路面面层");
                map.put("ccname2","");
                map.put("yxps",list.get(0).get("规定值"));
                map.put("filename","详见《混凝土路面相邻板高差质量鉴定表》检测"+list.get(0).get("检测点数")+"点,合格"+list.get(0).get("合格点数")+"点");
                map.put("sheetname", "分部-"+s);
                map.put("fbgc", "路面面层");
                map.put("合格率", list.get(0).get("合格率"));
                relist.add(map);
                resultlist.addAll(relist);

            }else if (value.contains("51构造深度手工铺沙法-")){
                List<Map<String, Object>> list = jjgFbgcSdgcLmgzsdsgpsfService.lookjg(commonInfoVo,value);
                List<Map<String, Object>> relist = new ArrayList<>();
                String s = StringUtils.substringBetween(value, "-", ".");
                Map map = new HashMap();
                map.put("ccname","构造深度手工铺沙法");
                map.put("ccname1","路面面层");
                map.put("ccname2","");
                map.put("yxps",list.get(0).get("设计值"));
                map.put("filename","详见《混凝土路面构造深度质量鉴定表》检测"+list.get(0).get("检测点数")+"点,合格"+list.get(0).get("合格点数")+"点");
                map.put("sheetname", "分部-"+s);
                map.put("fbgc", "路面面层");
                map.put("合格率", list.get(0).get("合格率"));
                relist.add(map);
                resultlist.addAll(relist);

            }else if (value.contains("54隧道混凝土路面厚度-钻芯法-")){
                List<Map<String, Object>> list = jjgFbgcSdgcSdhntlmhdzxfService.lookjg(commonInfoVo,value);
                List<Map<String, Object>> relist = new ArrayList<>();
                String s = StringUtils.substringBetween(value, "-", ".");
                double szds = 0;
                double shgds = 0;
                String sgdz="";
                String jcxm1="";
                boolean a= false;

                double qzds = 0;
                double qhgds = 0;
                String qgdz="";
                String jcxm2="";
                boolean b= false;

                double mzds = 0;
                double mhgds = 0;
                String mgdz="";
                String jcxm3="";
                boolean c= false;
                for (Map<String, Object> map : list) {
                    if (map.get("路面类型").toString().contains("隧道")){
                        szds += Double.valueOf(map.get("检测点数").toString());
                        shgds += Double.valueOf(map.get("合格点数").toString());
                        sgdz = map.get("设计值").toString();
                        jcxm1 = map.get("检测项目").toString();
                        a = true;

                    }else if (map.get("路面类型").toString().contains("桥面")){
                        qzds += Double.valueOf(map.get("检测点数").toString());
                        qhgds += Double.valueOf(map.get("合格点数").toString());
                        qgdz = map.get("设计值").toString();
                        jcxm2 = map.get("检测项目").toString();
                        b= true;
                    } else {
                        mzds += Double.valueOf(map.get("检测点数").toString());
                        mhgds += Double.valueOf(map.get("合格点数").toString());
                        mgdz = map.get("设计值").toString();
                        jcxm3 = map.get("检测项目").toString();
                        c= true;
                    }
                }
                if (a){
                    Map map = new HashMap();
                    map.put("ccname","混凝土路面厚度-钻芯法");
                    map.put("ccname1","隧道路面");
                    map.put("ccname2","");
                    map.put("yxps",sgdz);
                    map.put("filename","详见《混凝土路面厚度质量鉴定表（钻芯法）》检测"+decf.format(szds)+"点,合格"+decf.format(shgds)+"点");
                    map.put("sheetname", "分部-"+jcxm1);
                    map.put("fbgc", "路面面层");
                    map.put("合格率",(szds != 0) ? df.format(shgds/szds*100) : "0");
                    relist.add(map);
                }
                if (b){
                    Map map = new HashMap();
                    map.put("ccname","混凝土路面厚度-钻芯法");
                    map.put("ccname1","桥面系");
                    map.put("ccname2","");
                    map.put("yxps",qgdz);
                    map.put("filename","详见《混凝土路面厚度质量鉴定表（钻芯法） 》检测"+decf.format(qzds)+"点,合格"+decf.format(qhgds)+"点");
                    map.put("sheetname", "分部-"+jcxm2);
                    map.put("fbgc", "路面面层");
                    map.put("合格率", (qhgds != 0) ? df.format(qhgds/qzds*100) : "0");
                    relist.add(map);
                }
                if (c){
                    Map map = new HashMap();
                    map.put("ccname","混凝土路面厚度-钻芯法");
                    map.put("ccname1","路面面层");
                    map.put("ccname2","");
                    map.put("yxps",mgdz);
                    map.put("filename","详见《混凝土路面厚度质量鉴定表（钻芯法） 》检测"+decf.format(mzds)+"点,合格"+decf.format(mhgds)+"点");
                    map.put("sheetname", "分部-"+jcxm3);
                    map.put("fbgc", "路面面层");
                    map.put("合格率", (mhgds != 0) ? df.format(mhgds/mzds*100) : "0");
                    relist.add(map);
                }
                resultlist.addAll(relist);

            }else if (value.contains("55隧道横坡-")){
                List<Map<String, Object>> list = jjgFbgcSdgcSdhpService.lookjg(commonInfoVo,value);
                String ccname3 ="";
                if (list.get(0).get("路面类型").toString().contains("沥青")){
                    ccname3 = "沥青路面";
                }else if(list.get(0).get("路面类型").toString().contains("混凝土")) {
                    ccname3 = "混凝土路面";
                }

                List<Map<String, Object>> relist = new ArrayList<>();
                String s = StringUtils.substringBetween(value, "-", ".");
                double szds = 0;
                double shgds = 0;
                String sgdz="";
                String jcxm1="";
                boolean a= false;

                double qzds = 0;
                double qhgds = 0;
                String qgdz="";
                String jcxm2="";
                boolean b= false;

                double mzds = 0;
                double mhgds = 0;
                String mgdz="";
                String jcxm3="";
                boolean c= false;
                for (Map<String, Object> map : list) {
                    if (map.get("路面类型").toString().contains("隧道")){
                        szds += Double.valueOf(map.get("检测点数").toString());
                        shgds += Double.valueOf(map.get("合格点数").toString());
                        sgdz = map.get("设计值").toString();
                        jcxm1 = map.get("检测项目").toString();
                        a = true;

                    }else if (map.get("路面类型").toString().contains("桥面")){
                        qzds += Double.valueOf(map.get("检测点数").toString());
                        qhgds += Double.valueOf(map.get("合格点数").toString());
                        qgdz = map.get("设计值").toString();
                        jcxm2 = map.get("检测项目").toString();
                        b= true;
                    } else {
                        mzds += Double.valueOf(map.get("检测点数").toString());
                        mhgds += Double.valueOf(map.get("合格点数").toString());
                        mgdz = map.get("设计值").toString();
                        jcxm3 = map.get("检测项目").toString();
                        c= true;
                    }
                }
                if (a){
                    Map map = new HashMap();
                    map.put("ccname","路面横坡");
                    map.put("ccname1","隧道路面");
                    map.put("ccname2",ccname3);
                    map.put("yxps",sgdz);
                    map.put("filename","详见《"+ccname3+"横坡质量鉴定表》检测"+decf.format(szds)+"点,合格"+decf.format(shgds)+"点");
                    map.put("sheetname", "分部-"+jcxm1);
                    map.put("fbgc", "路面面层");
                    map.put("合格率",(szds != 0) ? df.format(shgds/szds*100) : "0");
                    relist.add(map);
                }
                if (b){
                    Map map = new HashMap();
                    map.put("ccname","路面横坡");
                    map.put("ccname1","桥面系");
                    map.put("ccname2",ccname3);
                    map.put("yxps",qgdz);
                    map.put("filename","详见《"+ccname3+"横坡质量鉴定表》检测"+decf.format(qzds)+"点,合格"+decf.format(qhgds)+"点");
                    map.put("sheetname", "分部-"+jcxm2);
                    map.put("fbgc", "路面面层");
                    map.put("合格率", (qhgds != 0) ? df.format(qhgds/qzds*100) : "0");
                    relist.add(map);
                }
                if (c){
                    Map map = new HashMap();
                    map.put("ccname","路面横坡");
                    map.put("ccname1","路面面层");
                    map.put("ccname2",ccname3);
                    map.put("yxps",mgdz);
                    map.put("filename","详见《"+ccname3+"横坡质量鉴定表》检测"+decf.format(mzds)+"点,合格"+decf.format(mhgds)+"点");
                    map.put("sheetname", "分部-"+jcxm3);
                    map.put("fbgc", "路面面层");
                    map.put("合格率", (mhgds != 0) ? df.format(mhgds/mzds*100) : "0");
                    relist.add(map);
                }
                resultlist.addAll(relist);


            }else if (value.contains("53隧道沥青路面厚度-钻芯法-")){
                List<Map<String, Object>> list = jjgFbgcSdgcGssdlqlmhdzxfService.lookjg(commonInfoVo,value);
                List<Map<String, Object>> relist = new ArrayList<>();
                String s = StringUtils.substringBetween(value, "-", ".");
                double szds = 0;
                double shgds = 0;
                String sgdz="";
                String jcxm1="";
                boolean a= false;

                double qzds = 0;
                double qhgds = 0;
                String qgdz="";
                String jcxm2="";
                boolean b= false;

                double mzds = 0;
                double mhgds = 0;
                String mgdz="";
                String jcxm3="";
                boolean c= false;
                for (Map<String, Object> map : list) {
                    if (map.get("路面类型").toString().contains("隧道")){
                        szds += Double.valueOf(map.get("检测点数").toString());
                        shgds += Double.valueOf(map.get("合格点数").toString());
                        sgdz = map.get("设计值").toString();
                        jcxm1 = map.get("检测项目").toString();
                        a = true;

                    }else if (map.get("路面类型").toString().contains("桥面")){
                        qzds += Double.valueOf(map.get("检测点数").toString());
                        qhgds += Double.valueOf(map.get("合格点数").toString());
                        qgdz = map.get("设计值").toString();
                        jcxm2 = map.get("检测项目").toString();
                        b= true;
                    } else {
                        mzds += Double.valueOf(map.get("检测点数").toString());
                        mhgds += Double.valueOf(map.get("合格点数").toString());
                        mgdz = map.get("设计值").toString();
                        jcxm3 = map.get("检测项目").toString();
                        c= true;
                    }
                }
                if (a){
                    Map map = new HashMap();
                    map.put("ccname","混凝土路面厚度-钻芯法");
                    map.put("ccname1","隧道路面");
                    map.put("ccname2","");
                    map.put("yxps",sgdz);
                    map.put("filename","详见《混凝土路面厚度质量鉴定表（钻芯法）》检测"+decf.format(szds)+"点,合格"+decf.format(shgds)+"点");
                    map.put("sheetname", "分部-"+jcxm1);
                    map.put("fbgc", "路面面层");
                    map.put("合格率",(szds != 0) ? df.format(shgds/szds*100) : "0");
                    relist.add(map);
                }
                if (b){
                    Map map = new HashMap();
                    map.put("ccname","混凝土路面厚度-钻芯法");
                    map.put("ccname1","桥面系");
                    map.put("ccname2","");
                    map.put("yxps",qgdz);
                    map.put("filename","详见《混凝土路面厚度质量鉴定表（钻芯法） 》检测"+decf.format(qzds)+"点,合格"+decf.format(qhgds)+"点");
                    map.put("sheetname", "分部-"+jcxm2);
                    map.put("fbgc", "路面面层");
                    map.put("合格率", (qhgds != 0) ? df.format(qhgds/qzds*100) : "0");
                    relist.add(map);
                }
                if (c){
                    Map map = new HashMap();
                    map.put("ccname","混凝土路面厚度-钻芯法");
                    map.put("ccname1","路面面层");
                    map.put("ccname2","");
                    map.put("yxps",mgdz);
                    map.put("filename","详见《混凝土路面厚度质量鉴定表（钻芯法） 》检测"+decf.format(mzds)+"点,合格"+decf.format(mhgds)+"点");
                    map.put("sheetname", "分部-"+jcxm3);
                    map.put("fbgc", "路面面层");
                    map.put("合格率", (mhgds != 0) ? df.format(mhgds/mzds*100) : "0");
                    relist.add(map);
                }
                resultlist.addAll(relist);

            }

        }
        Map<String, List<Map<String, Object>>> groupedData = resultlist.stream()
                .filter(map -> map.get("sheetname") != null) // 添加非空判断
                .collect(Collectors.groupingBy(map -> (String) map.get("sheetname")));
        //按分组后的数据，每个分组写一个工作簿
        DBwriteToExcel(groupedData,proname,htd);

    }

    @Override
    public void generateJSZLPdb(CommonInfoVo commonInfoVo) throws IOException {
        /**
         *
         */
        String proname = commonInfoVo.getProname();
        String path = filespath+ File.separator+proname+File.separator;
        List<String> filteredFiles = filterFiles(path);

        List<Map<String,Object>> mapList = new ArrayList<>();

        for (String filteredFile : filteredFiles) {
            String pdnpath = path+filteredFile+ File.separator;
            XSSFWorkbook wb = null;
            File f = new File(pdnpath+"00评定表.xlsx");
            if (!f.exists()){
                break;
            }else {
                FileInputStream in = new FileInputStream(f);
                wb = new XSSFWorkbook(in);
                XSSFSheet sheet = wb.getSheet("合同段");
                Map map = new HashMap();
                for (int i = sheet.getFirstRowNum(); i < sheet.getPhysicalNumberOfRows(); i++) {
                    XSSFRow row = sheet.getRow(i);

                    if (row.getCell(0).getStringCellValue().equals("合同段质量等级")){
                        String djValue = row.getCell(1).getStringCellValue();
                        map.put("htdValue",sheet.getRow(1).getCell(1).getStringCellValue());
                        map.put("djValue",djValue);
                        mapList.add(map);
                    }

                }
            }
        }

        DBwriteJSZLToExcel(mapList,commonInfoVo);
    }



    /**
     *
     * @param mapList
     * @param commonInfoVo
     * @throws IOException
     */
    private void DBwriteJSZLToExcel(List<Map<String, Object>> mapList, CommonInfoVo commonInfoVo) throws IOException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        String proname = commonInfoVo.getProname();
        XSSFWorkbook wb = null;
        File f = new File(filepath + File.separator + proname + File.separator + "建设项目质量评定表.xlsx");
        File fdir = new File(filepath + File.separator + proname);
        if (!fdir.exists()) {
            //创建文件根目录
            fdir.mkdirs();
        }
        File directory = new File("src/main/resources/static");
        String reportPath = directory.getCanonicalPath();
        String name = "建设项目质量评定表.xlsx";
        String path = reportPath + File.separator + name;
        Files.copy(Paths.get(path), new FileOutputStream(f));
        FileInputStream out = new FileInputStream(f);
        wb = new XSSFWorkbook(out);

        XSSFSheet sheet = wb.getSheet("建设项目");
        createdJSXMTable(wb,getJSXMNum(mapList.size()));

        sheet.getRow(1).getCell(2).setCellValue(proname);
        sheet.getRow(1).getCell(6).setCellValue("待确认");
        sheet.getRow(2).getCell(2).setCellValue(getAllzh(proname));
        sheet.getRow(2).getCell(6).setCellValue(simpleDateFormat.format(new Date()));
        int index = 0;
        for (Map<String, Object> map : mapList) {
            sheet.getRow(index + 4).getCell(0).setCellValue(map.get("htdValue").toString());
            sheet.getRow(index + 4).getCell(3).setCellValue(map.get("djValue").toString());
            index++;
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
     * @return
     */
    private String getAllzh(String proname) {
        String zh = jjgHtdService.getAllzh(proname);
        return zh;
    }

    /**
     *
     * @param wb
     * @param record
     */
    private void createdJSXMTable(XSSFWorkbook wb, int record) {
        //record = record+1;
        for(int i = 1; i < record; i++){
            //if(i < record-1){
                RowCopy.copyRows(wb, "建设项目", "建设项目", 4, 29, (i - 1) * 26 + 30);
            /*}
            else{
                RowCopy.copyRows(wb, "建设项目", "建设项目", 4, 27, (i - 1) * 26 + 30);
            }*/
        }
        XSSFSheet sheet = wb.getSheet("建设项目");
        int lastRowNum = sheet.getLastRowNum();
        int numMergedRegions = sheet.getNumMergedRegions();

        for (int i = numMergedRegions - 1; i >= 0; i--) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            int firstRow = mergedRegion.getFirstRow();
            int lastRow = mergedRegion.getLastRow();

            if (lastRow >= lastRowNum - 1 && lastRow <= lastRowNum) {
                sheet.removeMergedRegion(i);
            }
        }
        RowCopy.copyRows(wb, "source", "建设项目", 0, 1,(record) * 26+2);
        wb.setPrintArea(wb.getSheetIndex("建设项目"), 0, 7, 0, (record) * 26+3);
    }

    /**
     *
     * @param size
     * @return
     */
    private int getJSXMNum(int size) {

        return size%26 ==0 ? size/26 : size/26+1;
    }

    /**
     *
     * @param groupedData
     */
    private void DBwriteToExcel(Map<String, List<Map<String, Object>>> groupedData,String proname,String htd) throws IOException {
        XSSFWorkbook wb = null;
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
        File fdir = new File(filepath + File.separator + proname + File.separator + htd);
        if (!fdir.exists()) {
            //创建文件根目录
            fdir.mkdirs();
        }
        File directory = new File("src/main/resources/static");
        String reportPath = directory.getCanonicalPath();
        String name = "合同段评定表.xlsx";
        String path = reportPath + File.separator + name;
        Files.copy(Paths.get(path), new FileOutputStream(f));
        FileInputStream out = new FileInputStream(f);
        wb = new XSSFWorkbook(out);

        //分部工程
        for (String key : groupedData.keySet()) {
            List<Map<String, Object>> list = groupedData.get(key);

            if (key.equals("分部-路基")){
                writeLJData(wb,list,key,proname,htd);
            }else if (key.equals("分部-路面")){
                writeLJData(wb,list,key,proname,htd);
            }else if (key.equals("分部-交安")){
                writeLJData(wb,list,key,proname,htd);
            }else {
                //桥梁和隧道
                writeLJData(wb,list,key,proname,htd);
            }
        }
        //单位工程
        List<Map<String, Object>> dwgclist = new ArrayList<>();
        for (Sheet sheet : wb) {
            String sheetName = sheet.getSheetName();
            // 检查工作表名是否以"分部-"开头
            if (sheetName.startsWith("分部-")) {
                // 处理工作表数据
                List<Map<String, Object>> list = processSheet(sheet);
                dwgclist.addAll(list);
            }
        }
        writedwgcData(wb,dwgclist);

        //合同段
        List<Map<String, Object>> list = processhtdSheet(wb);
        writedhtdData(wb,list);

        FileOutputStream fileOut = new FileOutputStream(f);
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
        out.close();
        wb.close();


    }

    /**
     *
     * @param wb
     * @param list
     */
    private void writedhtdData(XSSFWorkbook wb, List<Map<String, Object>> list) {
        XSSFSheet sheet = wb.getSheet("合同段");
        createdHtdTable(wb,getNum(list.size()));
        int index = 0;
        int tableNum = 0;
        fillTitleHtdData(sheet, tableNum, list.get(0));
        for (Map<String, Object> datum : list) {
            if (index > 20){
                tableNum++;
                index = 0;
            }
            fillCommonHtdData(sheet,tableNum,index, datum);
            index++;
        }
    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param datum
     */
    private void fillTitleHtdData(XSSFSheet sheet, int tableNum, Map<String, Object> datum) {
        sheet.getRow(tableNum * 29 +1).getCell(1).setCellValue(datum.get("htd").toString());
        sheet.getRow(tableNum * 29 + 1).getCell(5).setCellValue(datum.get("jsxm").toString());
        sheet.getRow(tableNum * 29 + 2).getCell(1).setCellValue(datum.get("sgdw").toString());
        sheet.getRow(tableNum * 29 + 2).getCell(5).setCellValue(datum.get("jldw").toString());

    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param index
     * @param datum
     */
    private void fillCommonHtdData(XSSFSheet sheet, int tableNum, int index, Map<String, Object> datum) {
        sheet.getRow(index + 5).getCell(0).setCellValue(datum.get("htd").toString());
        sheet.getRow(index + 5).getCell(1).setCellValue(datum.get("dwgc").toString());
        sheet.getRow(index + 5).getCell(4).setCellValue(datum.get("zldj").toString());

    }

    /**
     *
     * @param wb
     * @param gettableNum
     */
    private void createdHtdTable(XSSFWorkbook wb, int gettableNum) {
        int record = 0;
        record = gettableNum;
        for(int i = 1; i < record; i++){
            if(i < record-1){
                RowCopy.copyRows(wb, "合同段", "合同段", 5, 28, (i - 1) * 24 + 29);
            }
            else{
                RowCopy.copyRows(wb, "合同段", "合同段", 5, 26, (i - 1) * 24 + 29);
            }
        }
        XSSFSheet sheet = wb.getSheet("合同段");
        int lastRowNum = sheet.getLastRowNum();
        int numMergedRegions = sheet.getNumMergedRegions();

        for (int i = numMergedRegions - 1; i >= 0; i--) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            int firstRow = mergedRegion.getFirstRow();
            int lastRow = mergedRegion.getLastRow();

            if (lastRow >= lastRowNum - 2 && lastRow <= lastRowNum) {
                sheet.removeMergedRegion(i);
            }
        }
        RowCopy.copyRows(wb, "source", "合同段", 0, 2,(record) * 24+2);
        wb.setPrintArea(wb.getSheetIndex("合同段"), 0, 7, 0, (record) * 24+4);
    }

    /**
     *
     * @param size
     * @return
     */
    private int getNum(int size) {
        return size%24 ==0 ? size/24 : size/24+1;
    }

    /**
     *
     * @param wb
     * @return
     */
    private List<Map<String,Object>> processhtdSheet(XSSFWorkbook wb) {
        XSSFSheet sheet = wb.getSheet("单位工程");
        List<Map<String,Object>> list = new ArrayList<>();
        Row row;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }

            if ("单位工程名称：".equals(row.getCell(0).toString()) ) {
                Map map = new HashMap();
                map.put("dwgc",row.getCell(1).getStringCellValue());
                map.put("jsxm",row.getCell(3).getStringCellValue());
                map.put("htd",sheet.getRow(i+6).getCell(0).getStringCellValue());
                map.put("sgdw",sheet.getRow(i+2).getCell(1).getStringCellValue());
                map.put("jldw",sheet.getRow(i+2).getCell(3).getStringCellValue());
                map.put("zldj",sheet.getRow(i+24).getCell(1).getStringCellValue());
                list.add(map);
            }
        }
        return list;

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void writedwgcData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("单位工程");
        createdwgcTable(wb,getdwgcNum(data));

        int index = 0;
        int tableNum = 0;
        String fbgc = data.get(0).get("dwgc").toString();
        for (Map<String, Object> datum : data) {
            if (fbgc.equals(datum.get("dwgc"))){
                fillTitleDwgcData(sheet,tableNum,datum);
                fillCommonDwgcData(sheet,tableNum,index,datum);
                index++;
            }else {
                fbgc = datum.get("dwgc").toString();
                tableNum ++;
                index = 0;
                fillTitleDwgcData(sheet,tableNum,datum);
                fillCommonDwgcData(sheet,tableNum,index,datum);
                index += 1;
            }

        }
        calculateDwgcSheet(sheet);
        for (int j = 0; j < wb.getNumberOfSheets(); j++) {
            JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
        }

    }

    /**
     *
     * @param sheet
     */
    private void calculateDwgcSheet(XSSFSheet sheet) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if ("单位工程质量等级".equals(row.getCell(0).toString())) {
                rowstart = sheet.getRow(i-18);
                rowend = sheet.getRow(i-1);
                /*row.getCell(1).setCellFormula("IF(COUNTIF("+rowstart.getCell(2).getReference() +":"+rowend.getCell(2).getReference()
                        +",\"<>合格\")=0,\"合格\", \"不合格\")");*///=IF(COUNTIF(C64:C81,"<>合格")=0, "合格", "不合格")
                row.getCell(1).setCellFormula("IF(COUNTIF("+rowstart.getCell(2).getReference()
                        +":"+rowend.getCell(2).getReference()+",\"合格\")=COUNTA("+rowstart.getCell(2).getReference()
                        +":"+rowend.getCell(2).getReference()+"),\"合格\", \"不合格\")");//=IF(COUNTIF(T7:T21, "合格") = COUNTA(T7:T21), "√", "")

            }
        }
    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param datum
     */
    private void fillTitleDwgcData(XSSFSheet sheet, int tableNum, Map<String, Object> datum) {
        sheet.getRow(tableNum * 28 +1).getCell(1).setCellValue(datum.get("dwgc").toString());
        sheet.getRow(tableNum * 28 +1).getCell(3).setCellValue(datum.get("jsxm").toString());
        sheet.getRow(tableNum * 28 +2).getCell(1).setCellValue(datum.get("jsxm").toString());
        sheet.getRow(tableNum * 28 +2).getCell(3).setCellValue(datum.get("gcbw").toString());
        sheet.getRow(tableNum * 28 +3).getCell(1).setCellValue(datum.get("sgbw").toString());
        sheet.getRow(tableNum * 28 +3).getCell(3).setCellValue(datum.get("jldw").toString());


    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param index
     * @param datum
     */
    private void fillCommonDwgcData(XSSFSheet sheet, int tableNum, int index, Map<String, Object> datum) {
        sheet.getRow(tableNum * 28 + index + 7).getCell(0).setCellValue(datum.get("htd").toString());
        sheet.getRow(tableNum * 28 + index + 7).getCell(1).setCellValue(datum.get("fbgc").toString());
        sheet.getRow(tableNum * 28 + index + 7).getCell(2).setCellValue(datum.get("sfhg").toString());

    }

    /**
     * 单位工程
     * @param wb
     * @param gettableNum
     */
    private void createdwgcTable(XSSFWorkbook wb, int gettableNum) {
        for(int i = 1; i < gettableNum; i++){
            RowCopy.copyRows(wb, "单位工程", "单位工程", 0, 27, i*28);
        }
        if(gettableNum > 1){
            wb.setPrintArea(wb.getSheetIndex("单位工程"), 0, 5, 0, gettableNum*28-1);
        }

    }

    /**
     *
     * @param data
     * @return
     */
    private int getdwgcNum(List<Map<String, Object>> data) {
        Map<String, Integer> resultMap = new HashMap<>();
        for (Map<String, Object> map : data) {
            String name = map.get("dwgc").toString();
            if (resultMap.containsKey(name)) {
                resultMap.put(name, resultMap.get(name) + 1);
            } else {
                resultMap.put(name, 1);
            }
        }
        int num = 0;
        for (Map.Entry<String, Integer> entry : resultMap.entrySet()) {
            int value = entry.getValue();
            if (value%18==0){
                num += value/18;
            }else {
                num += value/18+1;
            }
        }
        return num;

    }

    /**
     *
     * @param sheet
     * @return
     */
    private List<Map<String,Object>> processSheet(Sheet sheet) {
        List<Map<String,Object>> list = new ArrayList<>();

        Row row;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }

            if ("合同段：".equals(row.getCell(0).toString()) ) {
                Map map = new HashMap();
                map.put("htd",row.getCell(2).getStringCellValue());
                map.put("fbgc",row.getCell(8).getStringCellValue());
                map.put("jsxm",row.getCell(15).getStringCellValue());

                map.put("gcbw",sheet.getRow(i+1).getCell(2).getStringCellValue());
                map.put("sgbw",sheet.getRow(i+1).getCell(8).getStringCellValue());
                map.put("jldw",sheet.getRow(i+1).getCell(15).getStringCellValue());
                if (sheet.getRow(i+20).getCell(16).getStringCellValue().equals("√")){
                    map.put("sfhg","合格");
                }else {
                    map.put("sfhg","不合格");
                }

                map.put("dwgc",sheet.getSheetName().substring(sheet.getSheetName().indexOf("-")+1));
                list.add(map);
            }
        }
        return list;
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void writeLJData(XSSFWorkbook wb, List<Map<String, Object>> data,String sheetname,String proname,String htd) {
        /*Collections.sort(data, new Comparator<Map<String, Object>>() {
            @Override
            public int compare(Map<String, Object> map1, Map<String, Object> map2) {
                String ccname1 = (String) map1.get("ccname");
                String ccname2 = (String) map2.get("ccname");
                return ccname1.compareTo(ccname2);
            }
        });*/
        copySheet(wb,sheetname);
        XSSFPrintSetup ps = wb.getSheet(sheetname).getPrintSetup();
        ps.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)

        JjgHtd htdlist = jjgHtdService.selectInfo(proname,htd);
        XSSFSheet sheet = wb.getSheet(sheetname);
        createTable(wb,gettableNum(data),sheetname);

        int index = 0;
        int tableNum = 0;
        String fbgc = data.get(0).get("fbgc").toString();
        for (Map<String, Object> datum : data) {
            if (fbgc.equals(datum.get("fbgc"))){
                fillTitleData(sheet,tableNum,proname,htd,htdlist,datum.get("fbgc").toString());
                fillCommonData(sheet,tableNum,index,datum);
                index++;
            }else {
                fbgc = datum.get("fbgc").toString();
                tableNum ++;
                index = 0;
                fillTitleData(sheet,tableNum,proname,htd,htdlist,datum.get("fbgc").toString());
                fillCommonData(sheet,tableNum,index,datum);
                index += 1;
            }

        }
        //写完当前工作簿的数据后，就要插入公式计算了
        calculateFbgcSheet(sheet);
        for (int j = 0; j < wb.getNumberOfSheets(); j++) {
            JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
        }




    }

    /**
     * 计算分部工程的结果
     * @param sheet
     */
    private void calculateFbgcSheet(XSSFSheet sheet) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if ("实测项目是否全部合格".equals(row.getCell(0).toString())) {
                rowstart = sheet.getRow(i-15);
                rowend = sheet.getRow(i-1);
                row.getCell(8).setCellFormula("IF(COUNTIF("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+",\"合格\")=COUNTA("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+"),\"√\", \"\")");//=IF(COUNTIF(T7:T21, "合格") = COUNTA(T7:T21), "√", "")
                row.getCell(10).setCellFormula("IF(COUNTIF("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+",\"合格\")=COUNTA("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+"),\"\", \"×\")");//=IF(COUNTIF(T7:T21, "不合格") = COUNTA(T7:T21), "", "×")

                row.getCell(16).setCellFormula("IF(COUNTIF("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+",\"合格\")=COUNTA("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+"),\"√\", \"\")");//=IF(COUNTIF(T7:T21, "合格") = COUNTA(T7:T21), "√", "")
                row.getCell(19).setCellFormula("IF(COUNTIF("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+",\"合格\")=COUNTA("+rowstart.getCell(19).getReference()+":"+rowend.getCell(19).getReference()+"),\"\", \"×\")");//=IF(COUNTIF(T7:T21, "不合格") = COUNTA(T7:T21), "", "×")
            }
        }

    }

    /**
     *
     * @param wb
     * @param sheetname
     */
    private void copySheet(XSSFWorkbook wb,String sheetname) {
        String sourceSheetName = "模板";
        String targetSheetName = sheetname;
        XSSFSheet sourceSheet = wb.getSheet(sourceSheetName);

        // 创建新的工作表，并将源工作表中的内容复制到新工作表
        XSSFSheet targetSheet = wb.createSheet(targetSheetName);
        copySheetInfo(wb,sourceSheet, targetSheet);


    }


    /**
     *
     * @param wb
     * @param sourceSheet
     * @param targetSheet
     */
    private static void copySheetInfo(XSSFWorkbook wb, Sheet sourceSheet, Sheet targetSheet) {
        int maxColumnNum = 0;
        for (int i = 0; i <= sourceSheet.getLastRowNum(); i++) {
            Row sourceRow = sourceSheet.getRow(i);
            Row newRow = targetSheet.createRow(i);
            if (sourceRow != null) {
                copyRow(sourceRow, newRow, wb);
                if (sourceRow.getLastCellNum() > maxColumnNum) {
                    maxColumnNum = sourceRow.getLastCellNum();
                }
            }
        }

        // 复制列宽
        for (int j = 0; j < maxColumnNum; j++) {
            targetSheet.setColumnWidth(j, sourceSheet.getColumnWidth(j));
        }

        // 复制合并单元格
        for (int i = 0; i < sourceSheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = sourceSheet.getMergedRegion(i);
            targetSheet.addMergedRegion(mergedRegion);
        }

        // 复制打印区域
        PrintSetup sourcePrintSetup = sourceSheet.getPrintSetup();
        PrintSetup targetPrintSetup = targetSheet.getPrintSetup();
        targetPrintSetup.setLandscape(sourcePrintSetup.getLandscape());
        targetPrintSetup.setPaperSize(sourcePrintSetup.getPaperSize());
        targetPrintSetup.setFitWidth(sourcePrintSetup.getFitWidth());
        targetPrintSetup.setFitHeight(sourcePrintSetup.getFitHeight());
        targetPrintSetup.setScale(sourcePrintSetup.getScale());

        // 复制页眉
        Header sourceHeader = sourceSheet.getHeader();
        Header targetHeader = targetSheet.getHeader();
        targetHeader.setCenter(sourceHeader.getCenter());
        targetHeader.setLeft(sourceHeader.getLeft());
        targetHeader.setRight(sourceHeader.getRight());

        // 复制页脚
        Footer sourceFooter = sourceSheet.getFooter();
        Footer targetFooter = targetSheet.getFooter();
        targetFooter.setCenter(sourceFooter.getCenter());
        targetFooter.setLeft(sourceFooter.getLeft());
        targetFooter.setRight(sourceFooter.getRight());


    }

    /**
     *
     * @param sourceRow
     * @param newRow
     * @param wb
     */
    private static void copyRow(Row sourceRow, Row newRow, Workbook wb) {
        newRow.setHeight(sourceRow.getHeight());
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            Cell newCell = newRow.createCell(i);
            if (sourceCell != null) {
                CellStyle sourceCellStyle = sourceCell.getCellStyle();
                CellStyle newCellStyle = wb.createCellStyle();
                newCellStyle.cloneStyleFrom(sourceCellStyle);
                newCell.setCellStyle(newCellStyle);

                CellType cellType = sourceCell.getCellTypeEnum();

                switch (cellType) {
                    case BOOLEAN:
                        newCell.setCellValue(sourceCell.getBooleanCellValue());
                        break;
                    case STRING:
                        newCell.setCellValue(sourceCell.getRichStringCellValue());
                        break;
                    case NUMERIC:
                        newCell.setCellValue(sourceCell.getNumericCellValue());
                        break;
                    case FORMULA:
                        newCell.setCellFormula(sourceCell.getCellFormula());
                        break;
                    case BLANK:
                        // do nothing
                        break;
                    default:
                        // do nothing
                        break;
                }
            }
        }
    }


    /**
     * 分部
     * @param sheet
     * @param tableNum
     * @param index
     * @param datum
     */
    private void fillCommonData(XSSFSheet sheet, int tableNum, int index, Map<String, Object> datum) {
        sheet.getRow(tableNum * 22 + index + 6).getCell(1).setCellValue(1+index);
        sheet.getRow(tableNum * 22 + index + 6).getCell(2).setCellValue(datum.get("ccname").toString());
        sheet.getRow(tableNum * 22 + index + 6).getCell(6).setCellValue(String.valueOf(datum.get("yxps")));
        sheet.getRow(tableNum * 22 + index + 6).getCell(7).setCellValue(datum.get("filename").toString());
        sheet.getRow(tableNum * 22 + index + 6).getCell(17).setCellValue(datum.get("合格率").toString());
        if (datum.get("ccname").toString().contains("*")){
            Double value = Double.valueOf(datum.get("合格率").toString());
            if (value == 100){
                sheet.getRow(tableNum * 22 + index + 6).getCell(19).setCellValue("合格");
            }else {
                sheet.getRow(tableNum * 22 + index + 6).getCell(19).setCellValue("不合格");
            }
        }else if (datum.get("ccname").toString().contains("△")){
            Double value = Double.valueOf(datum.get("合格率").toString());
            if (value >= 95){
                sheet.getRow(tableNum * 22 + index + 6).getCell(19).setCellValue("合格");
            }else {
                sheet.getRow(tableNum * 22 + index + 6).getCell(19).setCellValue("不合格");
            }
        }else {
            Double value = Double.valueOf(datum.get("合格率").toString());
            if (value >= 80){
                sheet.getRow(tableNum * 22 + index + 6).getCell(19).setCellValue("合格");
            }else {
                sheet.getRow(tableNum * 22 + index + 6).getCell(19).setCellValue("不合格");
            }
        }

    }

    /**
     * 分部
     * @param sheet
     * @param tableNum
     * @param proname
     * @param htd
     * @param htdlist
     * @param fbgc
     */
    private void fillTitleData(XSSFSheet sheet, int tableNum, String proname, String htd,JjgHtd htdlist,String fbgc){
        sheet.getRow(tableNum * 22 +1).getCell(2).setCellValue(htd);
        sheet.getRow(tableNum * 22 +1).getCell(8).setCellValue(fbgc);
        sheet.getRow(tableNum * 22 +2).getCell(2).setCellValue(getgcbw(htdlist.getZhq(),htdlist.getZhz()));
        sheet.getRow(tableNum * 22 +2).getCell(8).setCellValue(htdlist.getSgdw());
        sheet.getRow(tableNum * 22 +2).getCell(15).setCellValue(htdlist.getJldw());
        sheet.getRow(tableNum * 22 +1).getCell(15).setCellValue(proname);
    }

    /**
     *分部
     * @param zhq
     * @param zhz
     * @return
     */
    private String getgcbw(String zhq, String zhz) {
        DecimalFormat decimalFormat = new DecimalFormat("#.###");
        double a = Double.valueOf(zhq);
        double b = Double.valueOf(zhz);

        int aa = (int) a / 1000;
        int bb = (int) b / 1000;
        double cc = a % 1000;
        double dd = b % 1000;
        String result = "K"+aa+"+"+decimalFormat.format(cc)+"--"+"K"+bb+"+"+decimalFormat.format(dd);

       return result;
    }

    /**
     * 分部
     * @param wb
     * @param gettableNum
     * @param sheetname
     */
    private void createTable(XSSFWorkbook wb,int gettableNum,String sheetname) {
        for(int i = 1; i < gettableNum; i++){
            RowCopy.copyRows(wb, sheetname, sheetname, 0, 21, i*22);
        }
        if(gettableNum > 1){
            wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 20, 0, gettableNum*22-1);
        }
    }

    /**
     * 分部
     * @param data
     * @return
     */
    private int gettableNum(List<Map<String, Object>> data) {
        Map<String, Integer> resultMap = new HashMap<>();
        for (Map<String, Object> map : data) {
            String name = map.get("fbgc").toString();
            if (resultMap.containsKey(name)) {
                resultMap.put(name, resultMap.get(name) + 1);
            } else {
                resultMap.put(name, 1);
            }
        }
        int num = 0;
        for (Map.Entry<String, Integer> entry : resultMap.entrySet()) {
            int value = entry.getValue();
            if (value%15==0){
                num += value/15;
            }else {
                num += value/15+1;
            }
        }
        return num;
    }

    /**
     * 查找文件
     * @param folderPath
     * @return
     */
    private static List<String> filterFiles(String folderPath) {
        List<String> matchingFiles = new ArrayList<>();

        File folder = new File(folderPath);
        File[] files = folder.listFiles();

        if (files != null) {
            for (File file : files) {
                //if (file.isFile() && !file.getName().contains("桥梁") && !file.getName().contains("隧道")) {
                    matchingFiles.add(file.getName());
                //}
            }
        }
        return matchingFiles;
    }


    @Autowired
    private JjgLqsQlService jjgLqsQlService;

    @Override
    public void generateBGZBG(String proname) throws IOException {

        List<JjgHtd> htdlist = jjgHtdService.gethtd(proname);

        List<Map<String,Object>> ljlist = new ArrayList<>();
        List<Map<String,Object>> qllist = new ArrayList<>();
        List<Map<String,Object>> fhllist = new ArrayList<>();

        List<Map<String, Object>> tsflist = new ArrayList<>();
        List<Map<String, Object>> pslist = new ArrayList<>();
        List<Map<String, Object>> xqlist = new ArrayList<>();
        List<Map<String, Object>> hdlist = new ArrayList<>();
        List<Map<String, Object>> zdlist = new ArrayList<>();

        List<Map<String, Object>> lmysdlist = new ArrayList<>();
        List<Map<String, Object>> lmwclist = new ArrayList<>();
        List<Map<String, Object>> lmwclcflist = new ArrayList<>();
        List<Map<String, Object>> lmssxslist = new ArrayList<>();
        List<Map<String, Object>> lmhntqdlist = new ArrayList<>();
        List<Map<String, Object>> lmhntxlbgclist = new ArrayList<>();
        List<Map<String, Object>> lmpzdlist = new ArrayList<>();
        List<Map<String, Object>> lmczlist = new ArrayList<>();
        List<Map<String, Object>> lmhplist = new ArrayList<>();
        List<Map<String, Object>> ldhdlist = new ArrayList<>();
        List<Map<String, Object>> khlist = new ArrayList<>();
        List<Map<String, Object>> zxfhdlist = new ArrayList<>();
        List<Map<String, Object>> lmhdlist = new ArrayList<>();
        List<Map<String, Object>> qmpzdlist = new ArrayList<>();
        List<Map<String, Object>> qmhphplist = new ArrayList<>();
        List<Map<String, Object>> qmkhlist = new ArrayList<>();
        List<Map<String, Object>> sdysdlist = new ArrayList<>();
        List<Map<String, Object>> sdczlist = new ArrayList<>();
        List<Map<String, Object>> sdssxslist = new ArrayList<>();
        List<Map<String, Object>> sdhntqdlist = new ArrayList<>();
        List<Map<String, Object>> sdhntxlbgclist = new ArrayList<>();
        List<Map<String, Object>> sdpzdlist = new ArrayList<>();
        List<Map<String, Object>> sdkhlist = new ArrayList<>();
        List<Map<String, Object>> sdzxfhdlist = new ArrayList<>();
        List<Map<String, Object>> sdldhdlist = new ArrayList<>();
        List<Map<String, Object>> sdhdlist = new ArrayList<>();
        List<Map<String, Object>> sdhplist = new ArrayList<>();
        List<Map<String, Object>> qlxblist = new ArrayList<>();
        List<Map<String, Object>> qlsblist = new ArrayList<>();
        List<Map<String, Object>> qmxlist = new ArrayList<>();
        List<Map<String, Object>> cqlist = new ArrayList<>();
        List<Map<String, Object>> ztlist = new ArrayList<>();
        List<Map<String, Object>> sdlmlist = new ArrayList<>();
        List<Map<String, Object>> bzlist = new ArrayList<>();
        List<Map<String, Object>> bxlist = new ArrayList<>();
        List<Map<String, Object>> jafhllist = new ArrayList<>();
        List<Map<String, Object>> qdcclist = new ArrayList<>();
        List<Map<String, Object>> sdcjqdlist = new ArrayList<>();
        List<Map<String, Object>> qmxhzlist = new ArrayList<>();
        List<Map<String, Object>> sdlmhzblist = new ArrayList<>();



        List<Map<String, Object>> ljhzlist = new ArrayList<>();
        List<Map<String, Object>> lmhzlist = new ArrayList<>();
        List<Map<String, Object>> xmcclist = new ArrayList<>();
        List<Map<String, Object>> qlgclist = new ArrayList<>();
        List<Map<String, Object>> sdgclist = new ArrayList<>();
        List<Map<String, Object>> jabzlist = new ArrayList<>();
        List<Map<String, Object>> ysdfxblist = new ArrayList<>();
        List<Map<String, Object>> qlszdfxblist = new ArrayList<>();
        List<Map<String, Object>> sdcqfxblist = new ArrayList<>();
        List<Map<String, Object>> ljhzblist = new ArrayList<>();
        List<Map<String, Object>> qlpdlist = new ArrayList<>();
        List<Map<String, Object>> sdpdlist = new ArrayList<>();
        List<Map<String, Object>> htdpdlist = new ArrayList<>();
        List<Map<String, Object>> xsxmpdlist = new ArrayList<>();

        for (JjgHtd jjgHtd : htdlist) {
            String htd = jjgHtd.getName();
            String lx = jjgHtd.getLx();
            CommonInfoVo commonInfoVo = new CommonInfoVo();
            commonInfoVo.setProname(proname);
            commonInfoVo.setHtd(htd);
            if (lx.contains("路基工程")){
                //表3.4.1-1 抽查统计表
                List<Map<String,Object>> ljhtdlist = getLjcjtjbData(proname,htd);
                ljlist.addAll(ljhtdlist);
                //表4.1.1-1
                List<Map<String, Object>> tsf = gettsfData(commonInfoVo);
                tsflist.addAll(tsf);
                //表4.1.1-2
                List<Map<String, Object>> ps = getpsData(commonInfoVo);
                pslist.addAll(ps);
                //表4.1.1-3
                List<Map<String, Object>> xq = getxqData(commonInfoVo);
                xqlist.addAll(xq);
                //表4.1.1-4
                List<Map<String, Object>> hd = gethdData(commonInfoVo);
                hdlist.addAll(hd);
                //表4.1.1-4
                List<Map<String, Object>> zd = getzdData(commonInfoVo);
                zdlist.addAll(zd);

                //表5.1.1-1  路基单位工程质量评定汇总表
                List<Map<String, Object>> ljhzb = getljhzbData(commonInfoVo);
                ljhzblist.addAll(ljhzb);


            }
            if (lx.contains("路面工程")){
                //表4.1.2-1
                List<Map<String, Object>> lmysd = getlmysdData(commonInfoVo);
                lmysdlist.addAll(lmysd);
                //表4.1.2-2
                List<Map<String, Object>> lmwc = getlmwcData(commonInfoVo);
                lmwclist.addAll(lmwc);

                List<Map<String, Object>> lmwclcf = getlmwclcfData(commonInfoVo);
                lmwclcflist.addAll(lmwclcf);
                //车辙
                List<Map<String, Object>> cz = getczData(commonInfoVo);
                lmczlist.addAll(cz);

                //表4.1.2-3
                List<Map<String, Object>> lmssxs = getlmssxsData(commonInfoVo);
                lmssxslist.addAll(lmssxs);

                //表4.1.2-4
                List<Map<String, Object>> hntqd = gethntqdData(commonInfoVo);
                lmhntqdlist.addAll(hntqd);

                //表4.1.2-5
                List<Map<String, Object>> hntxlbgc = gethntxlbgcData(commonInfoVo);
                lmhntxlbgclist.addAll(hntxlbgc);

                //表4.1.2-6 平整度
                List<Map<String, Object>> pzd = getpzdData(commonInfoVo);
                lmpzdlist.addAll(pzd);

                //表4.1.2-7 抗滑
                List<Map<String, Object>> kh = getkhData(commonInfoVo);
                khlist.addAll(kh);

                //表4.1.2-8(1) 钻芯法厚度
                List<Map<String, Object>> zxfhd = getzxfhdData(commonInfoVo);
                zxfhdlist.addAll(zxfhd);

                //表4.1.2-8(2) 路面雷达厚度
                List<Map<String, Object>> ldhd = getldhdData(commonInfoVo);
                ldhdlist.addAll(ldhd);

                //表4.1.2-8(3) 厚度
                List<Map<String, Object>> lmhd = getlmhdData(commonInfoVo);
                lmhdlist.addAll(lmhd);

                //表4.1.2-9 横坡
                List<Map<String, Object>> hp = gethpData(commonInfoVo);
                lmhplist.addAll(hp);

                //表4.1.2-10 汇总

                //桥面平整度
                List<Map<String, Object>> qmpzd = getqmpzdData(commonInfoVo);
                qmpzdlist.addAll(qmpzd);

                //桥面横坡
                List<Map<String, Object>> qmhp = getqmhpData(commonInfoVo);
                qmhphplist.addAll(qmhp);

                //桥面抗滑
                List<Map<String, Object>> qmkh = getqmkhData(commonInfoVo);
                qmkhlist.addAll(qmkh);

                //桥面汇总

                //隧道路面压实度
                List<Map<String, Object>> sdysd = getsdysdData(commonInfoVo);
                sdysdlist.addAll(sdysd);

                //隧道路面车辙
                List<Map<String, Object>> sdcz = getsdczData(commonInfoVo);
                sdczlist.addAll(sdcz);

                //隧道路面渗水系数
                List<Map<String, Object>> sdssxs = getsdssxsData(commonInfoVo);
                sdssxslist.addAll(sdssxs);

                //隧道混凝土路面强度
                List<Map<String, Object>> sdhntqd = getsdhntqdData(commonInfoVo);
                if (CollectionUtils.isNotEmpty(sdhntqd)){
                    sdhntqdlist.addAll(sdhntqd);
                }

                //隧道混凝土路面相邻板高差检
                List<Map<String, Object>> sdhntxlbgc = getsdhntxlbgcData(commonInfoVo);
                if (CollectionUtils.isNotEmpty(sdhntxlbgc)){
                    sdhntxlbgclist.addAll(sdhntxlbgc);
                }

                //隧道路面平整度
                List<Map<String, Object>> sdpzd = getsdpzdData(commonInfoVo);
                sdpzdlist.addAll(sdpzd);

                //隧道路面抗滑
                List<Map<String, Object>> sdkh = getsdkhData(commonInfoVo);
                sdkhlist.addAll(sdkh);

                //隧道路面钻芯法厚度
                List<Map<String, Object>> sdzxfhd = getsdzxfhdData(commonInfoVo);
                sdzxfhdlist.addAll(sdzxfhd);

                //隧道路面雷达厚度
                List<Map<String, Object>> sdldhd = getsdldhdData(commonInfoVo);
                sdldhdlist.addAll(sdldhd);

                //隧道路面厚度
                List<Map<String, Object>> sdhd = getsdhdData(commonInfoVo);
                sdhdlist.addAll(sdhd);


                //隧道路面横坡
                List<Map<String, Object>> sdhp = getsdhpData(commonInfoVo);
                sdhplist.addAll(sdhp);

            }
            if (lx.contains("桥梁工程")){

                List<Map<String,Object>> ljhtdlist = getLjqlcjData(proname, htd);
                qllist.addAll(ljhtdlist);

                //表4.1.3-1 桥梁下部检测结果汇总表
                List<Map<String,Object>> qlxb = getqlxbData(commonInfoVo);
                qlxblist.addAll(qlxb);

                //表4.1.3-2 桥梁上部检测结果汇总表
                List<Map<String,Object>> qlsb = getqlsbData(commonInfoVo);
                qlsblist.addAll(qlsb);

                List<Map<String,Object>> qmx = getqmxData(commonInfoVo);
                qmxlist.addAll(qmx);

                List<Map<String, Object>> qlpd = getqlpdData(commonInfoVo);
                qlpdlist.addAll(qlpd);

            }
            if (lx.contains("交安工程")){
                //表3.4.3-1 抽检
                List<Map<String,Object>> ljhtdlist = getLjjaData(proname, htd);
                fhllist.addAll(ljhtdlist);

                //表4.1.5-1  标志检测结果汇总表
                List<Map<String,Object>> bz = getbzData(commonInfoVo);
                bzlist.addAll(bz);

                //表4.1.5-2  标线检测结果汇总表
                List<Map<String,Object>> bx = getbxData(commonInfoVo);
                bxlist.addAll(bx);

                //表4.1.5-3  防护栏（波形梁）检测结果汇总表
                List<Map<String,Object>> fhl = getfhlData(commonInfoVo);
                jafhllist.addAll(fhl);

                //表4.1.5-4  防护栏（砼防护栏）检测结果汇总表
                List<Map<String,Object>> qdcc = getqdccData(commonInfoVo);
                qdcclist.addAll(qdcc);
            }
            if (lx.contains("隧道工程")){

                List<Map<String,Object>> sdcjlist = getsdcjData(commonInfoVo);
                sdcjqdlist.addAll(sdcjlist);

                //表4.1.4-1 衬砌检测结果汇总表
                List<Map<String,Object>> cq = getcqData(commonInfoVo);
                cqlist.addAll(cq);

                //4.1.4-2总体检测结果汇总表
                List<Map<String,Object>> zt = getztData(commonInfoVo);
                ztlist.addAll(zt);

                //表4.1.4-3隧道路面检测结果汇总表
                List<Map<String,Object>> sdlm = getsdlmData(commonInfoVo);
                sdlmlist.addAll(sdlm);

                List<Map<String, Object>> sdpd = getsdpdData(commonInfoVo);
                sdpdlist.addAll(sdpd);

            }

            List<Map<String,Object>> ljjchzlist = getljhzData(commonInfoVo);
            ljhzlist.addAll(ljjchzlist);

            List<Map<String,Object>> lmjchzlist = getlmhzData(commonInfoVo);
            lmhzlist.addAll(lmjchzlist);

            List<Map<String,Object>> lmqlhzlist = getlmqlhzData(commonInfoVo);
            qmxhzlist.addAll(lmqlhzlist);

            List<Map<String,Object>> sdlmhzlist = getsdlmhzData(commonInfoVo);
            sdlmhzblist.addAll(sdlmhzlist);

            List<Map<String,Object>> xmcc = getxmccData(commonInfoVo);
            xmcclist.addAll(xmcc);

            List<Map<String,Object>> qlgc = getqlgcData(commonInfoVo);
            qlgclist.addAll(qlgc);

            List<Map<String,Object>> sdgc = getsdgcData(commonInfoVo);
            sdgclist.addAll(sdgc);

            List<Map<String,Object>> jabz = getjabzData(commonInfoVo);
            jabzlist.addAll(jabz);

            List<Map<String,Object>> ysdfxb = getysdfxbData(commonInfoVo);
            ysdfxblist.addAll(ysdfxb);

            List<Map<String,Object>> qlszdfxb = getqlszdfxbData(commonInfoVo);
            qlszdfxblist.addAll(qlszdfxb);

            List<Map<String,Object>> sdcqfxb = getsdcqfxbData(commonInfoVo);
            sdcqfxblist.addAll(sdcqfxb);

            List<Map<String,Object>> htdpd = gethtdpdData(commonInfoVo);
            htdpdlist.addAll(htdpd);

        }
        List<Map<String,Object>> xsxmpd = getxsxmpdData(proname);
        xsxmpdlist.addAll(xsxmpd);

        XSSFWorkbook wb = null;
        File f = new File(filepath + File.separator + proname + File.separator + "报告中表格.xlsx");
        File fdir = new File(filepath + File.separator + proname);
        if (!fdir.exists()) {
            fdir.mkdirs();
        }
        File directory = new File("src/main/resources/static");
        String reportPath = directory.getCanonicalPath();
        String name = "报告中表格.xlsx";
        String path = reportPath + File.separator + name;
        Files.copy(Paths.get(path), new FileOutputStream(f));
        FileInputStream out = new FileInputStream(f);
        wb = new XSSFWorkbook(out);

        if (CollectionUtils.isNotEmpty(ljlist)){
            DBExcelData(wb,ljlist,proname);
        }
        if (CollectionUtils.isNotEmpty(ljlist)){
            DBExcelQlData(wb,qllist,proname);
        }
        if (CollectionUtils.isNotEmpty(ljlist)){
            DBExcelJaData(wb,fhllist,proname);
        }
        if (CollectionUtils.isNotEmpty(sdcjqdlist)){
            DBExcelSdcjData(wb,sdcjqdlist);
        }
        if (CollectionUtils.isNotEmpty(tsflist)){
            DBExceltsfData(wb,tsflist);
        }
        if (CollectionUtils.isNotEmpty(pslist)){
            DBExcelpsData(wb,pslist);
        }

        if (CollectionUtils.isNotEmpty(xqlist)){
            DBExcelxqData(wb,xqlist);
        }

        if (CollectionUtils.isNotEmpty(hdlist)){
            DBExcelhdData(wb,hdlist);
        }

        if (CollectionUtils.isNotEmpty(zdlist)){
            DBExcelzdData(wb,zdlist);
        }

        if (CollectionUtils.isNotEmpty(ljhzlist)){
            DBExcelljhzData(wb,ljhzlist);
        }
        if (CollectionUtils.isNotEmpty(lmysdlist)){
            DBExcellmysdData(wb,lmysdlist);
        }

        if (CollectionUtils.isNotEmpty(lmwclist)){
            DBExcellmwcdData(wb,lmwclist);
        }
        if (CollectionUtils.isNotEmpty(lmwclcflist)){
            DBExcellmwclcfdData(wb,lmwclcflist);
        }
        if (CollectionUtils.isNotEmpty(lmczlist)){
            DBExcellmczData(wb,lmczlist);
        }
        if (CollectionUtils.isNotEmpty(lmssxslist)){
            DBExcellmssxsData(wb,lmssxslist);
        }
        if (CollectionUtils.isNotEmpty(lmhntqdlist)){
            DBExcellmhntqdData(wb,lmhntqdlist);
        }
        if (CollectionUtils.isNotEmpty(lmhntxlbgclist)){
            DBExcellmhntxlbgcData(wb,lmhntxlbgclist);
        }
        if (CollectionUtils.isNotEmpty(lmpzdlist)){
            DBExcellmpzdData(wb,lmpzdlist);
        }
        if (CollectionUtils.isNotEmpty(khlist)){
            DBExcelkhData(wb,khlist);
        }
        if (CollectionUtils.isNotEmpty(zxfhdlist)){
            DBExcelzxfhdData(wb,zxfhdlist);
        }
        if (CollectionUtils.isNotEmpty(ldhdlist)){
            DBExcelldhdData(wb,ldhdlist);
        }
        if (CollectionUtils.isNotEmpty(lmhdlist)){
            DBExcellmhdData(wb,lmhdlist);
        }
        if (CollectionUtils.isNotEmpty(lmhplist)){
            DBExcellmhpData(wb,lmhplist);
        }
        if (CollectionUtils.isNotEmpty(lmhzlist)){
            DBExcellmhzData(wb,lmhzlist);
        }
        if (CollectionUtils.isNotEmpty(qmpzdlist)){
            DBExcelqmpzdData(wb,qmpzdlist);
        }
        if (CollectionUtils.isNotEmpty(qmhphplist)){
            DBExcelqmhpdData(wb,qmhphplist);
        }
        if (CollectionUtils.isNotEmpty(qmkhlist)){
            DBExcelqmkhData(wb,qmkhlist);
        }
        if (CollectionUtils.isNotEmpty(qmxhzlist)){
            DBExcelqmxhzData(wb,qmxhzlist);
        }
        if (CollectionUtils.isNotEmpty(sdysdlist)){
            DBExcelsdysdData(wb,sdysdlist);
        }
        if (CollectionUtils.isNotEmpty(sdczlist)){
            DBExcelsdczData(wb,sdczlist);
        }
        if (CollectionUtils.isNotEmpty(sdssxslist)){
            DBExcelsdssxsData(wb,sdssxslist);
        }
        if (CollectionUtils.isNotEmpty(sdhntqdlist)){
            DBExcelsdhntqdData(wb,sdhntqdlist);
        }
        if (CollectionUtils.isNotEmpty(sdhntxlbgclist)){
            DBExcelsdhntxlbgcData(wb,sdhntxlbgclist);
        }
        if (CollectionUtils.isNotEmpty(sdpzdlist)){
            DBExcelsdpzdData(wb,sdpzdlist);
        }
        if (CollectionUtils.isNotEmpty(sdkhlist)){
            DBExcelsdkhData(wb,sdkhlist);
        }
        if (CollectionUtils.isNotEmpty(sdzxfhdlist)){
            DBExcelsdzxfhdData(wb,sdzxfhdlist);
        }
        if (CollectionUtils.isNotEmpty(sdldhdlist)){
            DBExcelsdldhdData(wb,sdldhdlist);
        }
        if (CollectionUtils.isNotEmpty(sdhdlist)){
            DBExcelsdhdData(wb,sdhdlist);
        }
        if (CollectionUtils.isNotEmpty(sdhplist)){
            DBExcelsdhpData(wb,sdhplist);
        }
        if (CollectionUtils.isNotEmpty(sdlmhzblist)){
            DBExcelsdlmhzData(wb,sdlmhzblist);
        }
        if (CollectionUtils.isNotEmpty(xmcclist)){
            DBExcelxmccData(wb,xmcclist);
        }
        if (CollectionUtils.isNotEmpty(qlxblist)){
            DBExcelqlxbData(wb,qlxblist);
        }
        if (CollectionUtils.isNotEmpty(qlsblist)){
            DBExcelqlsbData(wb,qlsblist);
        }
        if (CollectionUtils.isNotEmpty(qmxlist)){
            DBExcelqmxData(wb,qmxlist);
        }
        if (CollectionUtils.isNotEmpty(qlgclist)){
            DBExcelqlpdData(wb,qlgclist);
        }
        if (CollectionUtils.isNotEmpty(cqlist)){
            DBExcelsdcqData(wb,cqlist);
        }
        if (CollectionUtils.isNotEmpty(ztlist)){
            DBExcelsdztData(wb,ztlist);
        }
        if (CollectionUtils.isNotEmpty(sdlmlist)){
            DBExcelsdlmData(wb,sdlmlist);
        }
        if (CollectionUtils.isNotEmpty(sdgclist)){
            DBExcelsdgcData(wb,sdgclist);
        }
        if (CollectionUtils.isNotEmpty(bzlist)){
            DBExceljabzData(wb,bzlist);
        }
        if (CollectionUtils.isNotEmpty(bxlist)){
            DBExceljabxData(wb,bxlist);
        }
        if (CollectionUtils.isNotEmpty(jafhllist)){
            DBExceljafhlData(wb,jafhllist);
        }
        if (CollectionUtils.isNotEmpty(qdcclist)){
            DBExcelqdccData(wb,qdcclist);
        }
        if (CollectionUtils.isNotEmpty(jabzlist)){
            DBExceljabzhzData(wb,jabzlist);
        }
        if (CollectionUtils.isNotEmpty(ysdfxblist)){
            DBExcelysdfxData(wb,ysdfxblist);
        }
        if (CollectionUtils.isNotEmpty(qlszdfxblist)){
            DBExcelqlszdfxbData(wb,qlszdfxblist);
        }
        if (CollectionUtils.isNotEmpty(sdcqfxblist)){
            DBExcelsdcqfxData(wb,sdcqfxblist);
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
     * @param wb
     * @param data
     */
    private void DBExcelsdcqfxData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.6-3");
        int index = 2;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());

            if (datum.containsKey("检测项目")) {
                sheet.getRow(index).getCell(1).setCellValue(datum.get("检测项目").toString());
            } else {
                sheet.getRow(index).getCell(1).setCellValue("");
            }
            if (datum.containsKey("检测总点数")) {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("检测总点数").toString()));
            } else {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("合格点数")) {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("合格点数").toString()));
            } else {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("合格率")) {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("合格率").toString()));
            } else {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("ds")) {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("ds").toString()));
            } else {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(0));
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelqlszdfxbData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.6-2");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());

            if (datum.containsKey("gdz")) {
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("gdz").toString()));
            } else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }
            if (datum.containsKey("max")) {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("max").toString()));
            } else {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("min")) {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("min").toString()));
            } else {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("jcds")) {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("jcds").toString()));
            } else {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("hgds")) {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("hgds").toString()));
            } else {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("hgl")) {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("hgl").toString()));
            } else {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(0));
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelysdfxData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.6-1");
        int index = 2;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());

            if (datum.containsKey("bzz")) {
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("bzz").toString()));
            } else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }
            if (datum.containsKey("jz")) {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("jz").toString()));
            } else {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("max")) {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("max").toString()));
            } else {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("min")) {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("min").toString()));
            } else {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("dbz")) {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("dbz").toString()));
            } else {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("jcds")) {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("jcds").toString()));
            } else {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("hgds")) {
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("hgds").toString()));
            } else {
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("hgl")) {
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("hgl").toString()));
            } else {
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(0));
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExceljabzhzData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.5-5");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(2).getCell(index).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("szdzds")){
                sheet.getRow(3).getCell(index).setCellValue(Double.valueOf(datum.get("szdzds").toString()));
            }else {
                sheet.getRow(3).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("szdhgds")){
                sheet.getRow(4).getCell(index).setCellValue(Double.valueOf(datum.get("szdhgds").toString()));
            }else {
                sheet.getRow(4).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("szdhgl")){
                sheet.getRow(5).getCell(index).setCellValue(Double.valueOf(datum.get("szdhgl").toString()));
            }else {
                sheet.getRow(5).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("jkzds")){
                sheet.getRow(6).getCell(index).setCellValue(Double.valueOf(datum.get("jkzds").toString()));
            }else {
                sheet.getRow(6).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("jkhgds")){
                sheet.getRow(7).getCell(index).setCellValue(Double.valueOf(datum.get("jkhgds").toString()));
            }else {
                sheet.getRow(7).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("jkhgl")){
                sheet.getRow(8).getCell(index).setCellValue(Double.valueOf(datum.get("jkhgl").toString()));
            }else {
                sheet.getRow(8).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("bzbhdzds")){
                sheet.getRow(9).getCell(index).setCellValue(Double.valueOf(datum.get("bzbhdzds").toString()));
            }else {
                sheet.getRow(9).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("bzbhdhgds")){
                sheet.getRow(10).getCell(index).setCellValue(Double.valueOf(datum.get("bzbhdhgds").toString()));
            }else {
                sheet.getRow(10).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("bzbhdhgl")){
                sheet.getRow(11).getCell(index).setCellValue(Double.valueOf(datum.get("bzbhdhgl").toString()));
            }else {
                sheet.getRow(11).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("xszds")){
                sheet.getRow(12).getCell(index).setCellValue(Double.valueOf(datum.get("xszds").toString()));
            }else {
                sheet.getRow(12).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("xshgds")){
                sheet.getRow(13).getCell(index).setCellValue(Double.valueOf(datum.get("xshgds").toString()));
            }else {
                sheet.getRow(13).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("xshgl")){
                sheet.getRow(14).getCell(index).setCellValue(Double.valueOf(datum.get("xshgl").toString()));
            }else {
                sheet.getRow(14).getCell(index).setCellValue(Double.valueOf(0));
            }


            if (datum.containsKey("xzds")){
                sheet.getRow(15).getCell(index).setCellValue(Double.valueOf(datum.get("xzds").toString()));
            }else {
                sheet.getRow(15).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("xhgds")){
                sheet.getRow(16).getCell(index).setCellValue(Double.valueOf(datum.get("xhgds").toString()));
            }else {
                sheet.getRow(16).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("xhgl")){
                sheet.getRow(17).getCell(index).setCellValue(Double.valueOf(datum.get("xhgl").toString()));
            }else {
                sheet.getRow(17).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("bxhdzds")){
                sheet.getRow(18).getCell(index).setCellValue(Double.valueOf(datum.get("bxhdzds").toString()));
            }else {
                sheet.getRow(18).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("bxhdhgds")){
                sheet.getRow(19).getCell(index).setCellValue(Double.valueOf(datum.get("bxhdhgds").toString()));
            }else {
                sheet.getRow(19).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("bxhdhgl")){
                sheet.getRow(20).getCell(index).setCellValue(Double.valueOf(datum.get("bxhdhgl").toString()));
            }else {
                sheet.getRow(20).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("jzds")){
                sheet.getRow(21).getCell(index).setCellValue(Double.valueOf(datum.get("jzds").toString()));
            }else {
                sheet.getRow(21).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("jhgds")){
                sheet.getRow(22).getCell(index).setCellValue(Double.valueOf(datum.get("jhgds").toString()));
            }else {
                sheet.getRow(22).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("jhgl")){
                sheet.getRow(23).getCell(index).setCellValue(Double.valueOf(datum.get("jhgl").toString()));
            }else {
                sheet.getRow(23).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("lzds")){
                sheet.getRow(24).getCell(index).setCellValue(Double.valueOf(datum.get("lzds").toString()));
            }else {
                sheet.getRow(24).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("lhgds")){
                sheet.getRow(25).getCell(index).setCellValue(Double.valueOf(datum.get("lhgds").toString()));
            }else {
                sheet.getRow(25).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("lhgl")){
                sheet.getRow(26).getCell(index).setCellValue(Double.valueOf(datum.get("lhgl").toString()));
            }else {
                sheet.getRow(26).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("szds")){
                sheet.getRow(27).getCell(index).setCellValue(Double.valueOf(datum.get("szds").toString()));
            }else {
                sheet.getRow(27).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("shgds")){
                sheet.getRow(28).getCell(index).setCellValue(Double.valueOf(datum.get("shgds").toString()));
            }else {
                sheet.getRow(28).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("shgl")){
                sheet.getRow(29).getCell(index).setCellValue(Double.valueOf(datum.get("shgl").toString()));
            }else {
                sheet.getRow(29).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("gzds")){
                sheet.getRow(30).getCell(index).setCellValue(Double.valueOf(datum.get("gzds").toString()));
            }else {
                sheet.getRow(30).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("ghgds")){
                sheet.getRow(31).getCell(index).setCellValue(Double.valueOf(datum.get("ghgds").toString()));
            }else {
                sheet.getRow(31).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("ghgl")){
                sheet.getRow(32).getCell(index).setCellValue(Double.valueOf(datum.get("ghgl").toString()));
            }else {
                sheet.getRow(32).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("qdzds")){
                sheet.getRow(33).getCell(index).setCellValue(Double.valueOf(datum.get("qdzds").toString()));
            }else {
                sheet.getRow(33).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qdhgs")){
                sheet.getRow(34).getCell(index).setCellValue(Double.valueOf(datum.get("qdhgs").toString()));
            }else {
                sheet.getRow(34).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qdhgl")){
                sheet.getRow(35).getCell(index).setCellValue(Double.valueOf(datum.get("qdhgl").toString()));
            }else {
                sheet.getRow(35).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("cczds")){
                sheet.getRow(36).getCell(index).setCellValue(Double.valueOf(datum.get("cczds").toString()));
            }else {
                sheet.getRow(36).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("cchgs")){
                sheet.getRow(37).getCell(index).setCellValue(Double.valueOf(datum.get("cchgs").toString()));
            }else {
                sheet.getRow(37).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("cchgl")){
                sheet.getRow(38).getCell(index).setCellValue(Double.valueOf(datum.get("cchgl").toString()));
            }else {
                sheet.getRow(38).getCell(index).setCellValue(Double.valueOf(0));
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelqdccData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.5-4");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());

            if (datum.containsKey("qdzds")) {
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("qdzds").toString()));
            } else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }
            if (datum.containsKey("qdhgs")) {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("qdhgs").toString()));
            } else {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qdhgl")) {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("qdhgl").toString()));
            } else {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("cczds")) {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("cczds").toString()));
            } else {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("cchgs")) {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("cchgs").toString()));
            } else {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("cchgl")) {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("cchgl").toString()));
            } else {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(0));
            }
            index++;
        }


    }


    /**
     *
     * @param wb
     * @param data
     */
    private void DBExceljafhlData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.5-3");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());

            if (datum.containsKey("jzds")) {
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("jzds").toString()));
            } else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }
            if (datum.containsKey("jhgds")) {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("jhgds").toString()));
            } else {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("jhgl")) {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("jhgl").toString()));
            } else {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("lzds")) {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("lzds").toString()));
            } else {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("lhgds")) {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("lhgds").toString()));
            } else {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("lhgl")) {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("lhgl").toString()));
            } else {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(0));
            }


            if (datum.containsKey("szds")) {
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("szds").toString()));
            } else {
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("shgds")) {
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("shgds").toString()));
            } else {
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("shgl")) {
                sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("shgl").toString()));
            } else {
                sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("gzds")) {
                sheet.getRow(index).getCell(10).setCellValue(Double.valueOf(datum.get("gzds").toString()));
            } else {
                sheet.getRow(index).getCell(10).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("ghgds")) {
                sheet.getRow(index).getCell(11).setCellValue(Double.valueOf(datum.get("ghgds").toString()));
            } else {
                sheet.getRow(index).getCell(11).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("ghgl")) {
                sheet.getRow(index).getCell(12).setCellValue(Double.valueOf(datum.get("ghgl").toString()));
            } else {
                sheet.getRow(index).getCell(12).setCellValue(Double.valueOf(0));
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExceljabxData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.5-2");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());

            if (datum.containsKey("xzds")) {
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("xzds").toString()));
            } else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }
            if (datum.containsKey("xhgds")) {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("xhgds").toString()));
            } else {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("xhgl")) {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("xhgl").toString()));
            } else {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("bxhdzds")) {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("bxhdzds").toString()));
            } else {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("bxhdhgds")) {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("bxhdhgds").toString()));
            } else {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("bxhdhgl")) {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("bxhdhgl").toString()));
            } else {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(0));
            }
            index++;
        }

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExceljabzData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        /**
         * [{htd=LJ-1, jkzds=26.0, jkhgds=24.0, jkhgl=92.31, szdzds=326.0, szdhgds=279.0,
         * szdhgl=85.58, hdzds=346.0, hdhgds=346.0, hdhgl=100.00, xszds=873, xshgds=873, xshgl=100.00}]
         */
        XSSFSheet sheet = wb.getSheet("表4.1.5-1");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());

            if (datum.containsKey("szdzds")) {
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("szdzds").toString()));
            } else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }
            if (datum.containsKey("szdhgds")) {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("szdhgds").toString()));
            } else {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("szdhgl")) {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("szdhgl").toString()));
            } else {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("jkzds")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("jkzds").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("jkhgds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("jkhgds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("jkhgl")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("jkhgl").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("bzbhdzds")) {
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("bzbhdzds").toString()));
            } else {
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("bzbhdhgds")) {
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("bzbhdhgds").toString()));
            } else {
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("bzbhdhgl")) {
                sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("bzbhdhgl").toString()));
            } else {
                sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("xszds")) {
                sheet.getRow(index).getCell(10).setCellValue(Double.valueOf(datum.get("xszds").toString()));
            } else {
                sheet.getRow(index).getCell(10).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("xshgds")) {
                sheet.getRow(index).getCell(11).setCellValue(Double.valueOf(datum.get("xshgds").toString()));
            } else {
                sheet.getRow(index).getCell(11).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("xshgl")) {
                sheet.getRow(index).getCell(12).setCellValue(Double.valueOf(datum.get("xshgl").toString()));
            } else {
                sheet.getRow(index).getCell(12).setCellValue(Double.valueOf(0));
            }
            index++;
        }

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdgcData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.4-4");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(2).getCell(index).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("cqqdjcds")){
                sheet.getRow(3).getCell(index).setCellValue(Double.valueOf(datum.get("cqqdjcds").toString()));
            }else {
                sheet.getRow(3).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("cqqdhgds")){
                sheet.getRow(4).getCell(index).setCellValue(Double.valueOf(datum.get("cqqdhgds").toString()));
            }else {
                sheet.getRow(4).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("cqqdhgl")){
                sheet.getRow(5).getCell(index).setCellValue(Double.valueOf(datum.get("cqqdhgl").toString()));
            }else {
                sheet.getRow(5).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("cqhdjcds")){
                sheet.getRow(6).getCell(index).setCellValue(Double.valueOf(datum.get("cqhdjcds").toString()));
            }else {
                sheet.getRow(6).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("cqhdhgds")){
                sheet.getRow(7).getCell(index).setCellValue(Double.valueOf(datum.get("cqhdhgds").toString()));
            }else {
                sheet.getRow(7).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("cqhdhgl")){
                sheet.getRow(8).getCell(index).setCellValue(Double.valueOf(datum.get("cqhdhgl").toString()));
            }else {
                sheet.getRow(8).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("dmpzdjcds")){
                sheet.getRow(9).getCell(index).setCellValue(Double.valueOf(datum.get("dmpzdjcds").toString()));
            }else {
                sheet.getRow(9).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("dmpzdhgds")){
                sheet.getRow(10).getCell(index).setCellValue(Double.valueOf(datum.get("dmpzdhgds").toString()));
            }else {
                sheet.getRow(10).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("dmpzdhgl")){
                sheet.getRow(11).getCell(index).setCellValue(Double.valueOf(datum.get("dmpzdhgl").toString()));
            }else {
                sheet.getRow(11).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("kdjcds")){
                sheet.getRow(12).getCell(index).setCellValue(Double.valueOf(datum.get("kdjcds").toString()));
            }else {
                sheet.getRow(12).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("kdhgds")){
                sheet.getRow(13).getCell(index).setCellValue(Double.valueOf(datum.get("kdhgds").toString()));
            }else {
                sheet.getRow(13).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("kdhgl")){
                sheet.getRow(14).getCell(index).setCellValue(Double.valueOf(datum.get("kdhgl").toString()));
            }else {
                sheet.getRow(14).getCell(index).setCellValue(Double.valueOf(0));
            }
            /*if (datum.containsKey("cqhdjcds")){
                sheet.getRow(15).getCell(index).setCellValue(Double.valueOf(datum.get("cqhdjcds").toString()));
            }else {
                sheet.getRow(15).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("cqhdhgds")){
                sheet.getRow(16).getCell(index).setCellValue(Double.valueOf(datum.get("cqhdhgds").toString()));
            }else {
                sheet.getRow(16).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("cqhdhgl")){
                sheet.getRow(17).getCell(index).setCellValue(Double.valueOf(datum.get("cqhdhgl").toString()));
            }else {
                sheet.getRow(17).getCell(index).setCellValue(Double.valueOf(0));
            }*/

            if (datum.containsKey("ysdjcds")){
                sheet.getRow(18).getCell(index).setCellValue(Double.valueOf(datum.get("ysdjcds").toString()));
            }else {
                sheet.getRow(18).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("ysdhgds")){
                sheet.getRow(19).getCell(index).setCellValue(Double.valueOf(datum.get("ysdhgds").toString()));
            }else {
                sheet.getRow(19).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("ysdhgl")){
                sheet.getRow(20).getCell(index).setCellValue(Double.valueOf(datum.get("ysdhgl").toString()));
            }else {
                sheet.getRow(20).getCell(index).setCellValue(Double.valueOf(0));
            }

            //车辙
           /* if (datum.containsKey("cqhdjcds")){
                sheet.getRow(21).getCell(index).setCellValue(Double.valueOf(datum.get("cqhdjcds").toString()));
            }else {
                sheet.getRow(21).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("cqhdhgds")){
                sheet.getRow(22).getCell(index).setCellValue(Double.valueOf(datum.get("cqhdhgds").toString()));
            }else {
                sheet.getRow(22).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("cqhdhgl")){
                sheet.getRow(23).getCell(index).setCellValue(Double.valueOf(datum.get("cqhdhgl").toString()));
            }else {
                sheet.getRow(23).getCell(index).setCellValue(Double.valueOf(0));
            }*/

            if (datum.containsKey("xsjcds")){
                sheet.getRow(24).getCell(index).setCellValue(Double.valueOf(datum.get("xsjcds").toString()));
            }else {
                sheet.getRow(24).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("xshgds")){
                sheet.getRow(25).getCell(index).setCellValue(Double.valueOf(datum.get("xshgds").toString()));
            }else {
                sheet.getRow(25).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("xshgl")){
                sheet.getRow(26).getCell(index).setCellValue(Double.valueOf(datum.get("xshgl").toString()));
            }else {
                sheet.getRow(26).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("qdjcds")){
                sheet.getRow(27).getCell(index).setCellValue(Double.valueOf(datum.get("qdjcds").toString()));
            }else {
                sheet.getRow(27).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qdhgds")){
                sheet.getRow(28).getCell(index).setCellValue(Double.valueOf(datum.get("qdhgds").toString()));
            }else {
                sheet.getRow(28).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qdhgl")){
                sheet.getRow(29).getCell(index).setCellValue(Double.valueOf(datum.get("qdhgl").toString()));
            }else {
                sheet.getRow(29).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("gcjcds")){
                sheet.getRow(30).getCell(index).setCellValue(Double.valueOf(datum.get("gcjcds").toString()));
            }else {
                sheet.getRow(30).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("gchgds")){
                sheet.getRow(31).getCell(index).setCellValue(Double.valueOf(datum.get("gchgds").toString()));
            }else {
                sheet.getRow(31).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("gchgl")){
                sheet.getRow(32).getCell(index).setCellValue(Double.valueOf(datum.get("gchgl").toString()));
            }else {
                sheet.getRow(32).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("hdjcds")){
                sheet.getRow(39).getCell(index).setCellValue(Double.valueOf(datum.get("hdjcds").toString()));
            }else {
                sheet.getRow(39).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("hdhgds")){
                sheet.getRow(40).getCell(index).setCellValue(Double.valueOf(datum.get("hdhgds").toString()));
            }else {
                sheet.getRow(40).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("hdhgl")){
                sheet.getRow(41).getCell(index).setCellValue(Double.valueOf(datum.get("hdhgl").toString()));
            }else {
                sheet.getRow(41).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("hpjcds")){
                sheet.getRow(42).getCell(index).setCellValue(Double.valueOf(datum.get("hpjcds").toString()));
            }else {
                sheet.getRow(42).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("hphgds")){
                sheet.getRow(43).getCell(index).setCellValue(Double.valueOf(datum.get("hphgds").toString()));
            }else {
                sheet.getRow(43).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("hphgl")){
                sheet.getRow(44).getCell(index).setCellValue(Double.valueOf(datum.get("hphgl").toString()));
            }else {
                sheet.getRow(44).getCell(index).setCellValue(Double.valueOf(0));
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdlmData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.4-3");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());

            if (datum.containsKey("ysdjcds")){
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("ysdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }
            if (datum.containsKey("ysdhgds")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("ysdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("ysdhgl")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("ysdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(0));
            }

            //车辙
           /* if (datum.containsKey("cqhdjcds")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("cqhdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("cqhdhgds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("cqhdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("cqhdhgl")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("cqhdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(0));
            }*/

            if (datum.containsKey("xsjcds")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("xsjcds").toString()));
            }else {
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("xshgds")){
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("xshgds").toString()));
            }else {
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("xshgl")){
                sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("xshgl").toString()));
            }else {
                sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("qdjcds")){
                sheet.getRow(index).getCell(10).setCellValue(Double.valueOf(datum.get("qdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(10).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qdhgds")){
                sheet.getRow(index).getCell(11).setCellValue(Double.valueOf(datum.get("qdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(11).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qdhgl")){
                sheet.getRow(index).getCell(12).setCellValue(Double.valueOf(datum.get("qdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(12).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("gcjcds")){
                sheet.getRow(index).getCell(13).setCellValue(Double.valueOf(datum.get("gcjcds").toString()));
            }else {
                sheet.getRow(index).getCell(13).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("gchgds")){
                sheet.getRow(index).getCell(14).setCellValue(Double.valueOf(datum.get("gchgds").toString()));
            }else {
                sheet.getRow(index).getCell(14).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("gchgl")){
                sheet.getRow(index).getCell(15).setCellValue(Double.valueOf(datum.get("gchgl").toString()));
            }else {
                sheet.getRow(index).getCell(15).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("hdjcds")){
                sheet.getRow(index).getCell(22).setCellValue(Double.valueOf(datum.get("hdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(22).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("hdhgds")){
                sheet.getRow(index).getCell(23).setCellValue(Double.valueOf(datum.get("hdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(23).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("hdhgl")){
                sheet.getRow(index).getCell(24).setCellValue(Double.valueOf(datum.get("hdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(24).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("hpjcds")){
                sheet.getRow(index).getCell(25).setCellValue(Double.valueOf(datum.get("hpjcds").toString()));
            }else {
                sheet.getRow(index).getCell(25).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("hphgds")){
                sheet.getRow(index).getCell(26).setCellValue(Double.valueOf(datum.get("hphgds").toString()));
            }else {
                sheet.getRow(index).getCell(26).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("hphgl")){
                sheet.getRow(index).getCell(27).setCellValue(Double.valueOf(datum.get("hphgl").toString()));
            }else {
                sheet.getRow(index).getCell(27).setCellValue(Double.valueOf(0));
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdztData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.4-2");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("kdjcds")){
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("kdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }
            if (datum.containsKey("kdhgds")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("kdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("kdhgl")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("kdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(0));
            }
            /*if (datum.containsKey("cqhdjcds")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("cqhdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("cqhdhgds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("cqhdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("cqhdhgl")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("cqhdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(0));
            }*/
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdcqData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.4-1");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("cqqdjcds")){
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("cqqdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }
            if (datum.containsKey("cqqdhgds")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("cqqdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("cqqdhgl")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("cqqdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("cqhdjcds")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("cqhdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("cqhdhgds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("cqhdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("cqhdhgl")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("cqhdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("dmpzdjcds")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("dmpzdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("dmpzdhgds")){
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("dmpzdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("dmpzdhgl")){
                sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("dmpzdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(0));
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelqlpdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.3-4");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(2).getCell(index).setCellValue(datum.get("htd").toString());

            if (datum.containsKey("tqdjcds")) {
                sheet.getRow(3).getCell(index).setCellValue(Double.valueOf(datum.get("tqdjcds").toString()));
            } else {
                sheet.getRow(3).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("tqdhgds")) {
                sheet.getRow(4).getCell(index).setCellValue(Double.valueOf(datum.get("tqdhgds").toString()));
            } else {
                sheet.getRow(4).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("tqdhgl")) {
                sheet.getRow(5).getCell(index).setCellValue(Double.valueOf(datum.get("tqdhgl").toString()));
            } else {
                sheet.getRow(5).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("jgccjcds")) {
                sheet.getRow(6).getCell(index).setCellValue(Double.valueOf(datum.get("jgccjcds").toString()));
            } else {
                sheet.getRow(6).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("jgcchgds")) {
                sheet.getRow(7).getCell(index).setCellValue(Double.valueOf(datum.get("jgcchgds").toString()));
            } else {
                sheet.getRow(7).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("jgcchgl")) {
                sheet.getRow(8).getCell(index).setCellValue(Double.valueOf(datum.get("jgcchgl").toString()));
            } else {
                sheet.getRow(8).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("bhchdjcds")) {
                sheet.getRow(9).getCell(index).setCellValue(Double.valueOf(datum.get("bhchdjcds").toString()));
            } else {
                sheet.getRow(9).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("bhchdhgds")) {
                sheet.getRow(10).getCell(index).setCellValue(Double.valueOf(datum.get("bhchdhgds").toString()));
            } else {
                sheet.getRow(10).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("bhchdhgl")) {
                sheet.getRow(11).getCell(index).setCellValue(Double.valueOf(datum.get("bhchdhgl").toString()));
            } else {
                sheet.getRow(11).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("szdjcds")) {
                sheet.getRow(12).getCell(index).setCellValue(Double.valueOf(datum.get("szdjcds").toString()));
            } else {
                sheet.getRow(12).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("szdhgds")) {
                sheet.getRow(13).getCell(index).setCellValue(Double.valueOf(datum.get("szdhgds").toString()));
            } else {
                sheet.getRow(13).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("szdhgl")) {
                sheet.getRow(14).getCell(index).setCellValue(Double.valueOf(datum.get("szdhgl").toString()));
            } else {
                sheet.getRow(14).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("qlsbqdjcds")) {
                sheet.getRow(15).getCell(index).setCellValue(Double.valueOf(datum.get("qlsbqdjcds").toString()));
            } else {
                sheet.getRow(15).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("qlsbqdhgds")) {
                sheet.getRow(16).getCell(index).setCellValue(Double.valueOf(datum.get("qlsbqdhgds").toString()));
            } else {
                sheet.getRow(16).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("qlsbqdhgl")) {
                sheet.getRow(17).getCell(index).setCellValue(Double.valueOf(datum.get("qlsbqdhgl").toString()));
            } else {
                sheet.getRow(17).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("qlsbjgccjcds")) {
                sheet.getRow(18).getCell(index).setCellValue(Double.valueOf(datum.get("qlsbjgccjcds").toString()));
            } else {
                sheet.getRow(18).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("qlsbjgcchgds")) {
                sheet.getRow(19).getCell(index).setCellValue(Double.valueOf(datum.get("qlsbjgcchgds").toString()));
            } else {
                sheet.getRow(19).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("qlsbjgcchgl")) {
                sheet.getRow(20).getCell(index).setCellValue(Double.valueOf(datum.get("qlsbjgcchgl").toString()));
            } else {
                sheet.getRow(20).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("qlsbbhchdjcds")){
                sheet.getRow(21).getCell(index).setCellValue(Double.valueOf(datum.get("qlsbbhchdjcds").toString()));
            }else {
                sheet.getRow(21).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qlsbbhchdhgds")){
                sheet.getRow(22).getCell(index).setCellValue(Double.valueOf(datum.get("qlsbbhchdhgds").toString()));
            }else {
                sheet.getRow(22).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qlsbbhchdhgl")){
                sheet.getRow(23).getCell(index).setCellValue(Double.valueOf(datum.get("qlsbbhchdhgl").toString()));
            }else {
                sheet.getRow(23).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("qmhpzds")){
                sheet.getRow(27).getCell(index).setCellValue(Double.valueOf(datum.get("qmhpzds").toString()));
            }else {
                sheet.getRow(27).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qmhphgds")){
                sheet.getRow(28).getCell(index).setCellValue(Double.valueOf(datum.get("qmhphgds").toString()));
            }else {
                sheet.getRow(28).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qmhphgl")){
                sheet.getRow(29).getCell(index).setCellValue(Double.valueOf(datum.get("qmhphgl").toString()));
            }else {
                sheet.getRow(29).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("qmkhgzsdzds")){
                sheet.getRow(30).getCell(index).setCellValue(Double.valueOf(datum.get("qmkhgzsdzds").toString()));
            }else {
                sheet.getRow(30).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qmkhgzsdhgds")){
                sheet.getRow(31).getCell(index).setCellValue(Double.valueOf(datum.get("qmkhgzsdhgds").toString()));
            }else {
                sheet.getRow(31).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qmkhgzsdhgl")){
                sheet.getRow(32).getCell(index).setCellValue(Double.valueOf(datum.get("qmkhgzsdhgl").toString()));
            }else {
                sheet.getRow(32).getCell(index).setCellValue(Double.valueOf(0));
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelqmxData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.3-3");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());

            /*if (datum.containsKey("qlsbqdjcds")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("qlsbqdjcds").toString());
            }else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }
            if (datum.containsKey("qlsbqdhgds")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("qlsbqdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qlsbqdhgl")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("qlsbqdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(0));
            }*/

            if (datum.containsKey("qmhpzds")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("qmhpzds").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qmhphgds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("qmhphgds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qmhphgl")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("qmhphgl").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("qmkhgzsdzds")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("qmkhgzsdzds").toString()));
            }else {
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qmkhgzsdhgds")){
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("qmkhgzsdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qmkhgzsdhgl")){
                sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("qmkhgzsdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(0));
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelqlsbData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.3-2");
        XSSFSheet sheet = wb.getSheet("表4.1.3-2");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("qlsbqdjcds")){
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("qlsbqdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }
            if (datum.containsKey("qlsbqdhgds")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("qlsbqdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qlsbqdhgl")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("qlsbqdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("qlsbjgccjcds")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("qlsbjgccjcds").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qlsbjgcchgds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("qlsbjgcchgds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qlsbjgcchgl")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("qlsbjgcchgl").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("qlsbbhchdjcds")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("qlsbbhchdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qlsbbhchdhgds")){
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("qlsbbhchdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("qlsbbhchdhgl")){
                sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("qlsbbhchdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(0));
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelqlxbData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.3-1");
        XSSFSheet sheet = wb.getSheet("表4.1.3-1");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("tqdjcds")){
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("tqdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }
            if (datum.containsKey("tqdhgds")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("tqdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("tqdhgl")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("tqdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("jgccjcds")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("jgccjcds").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("jgcchgds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("jgcchgds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("jgcchgl")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("jgcchgl").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("bhchdjcds")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("bhchdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("bhchdhgds")){
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("bhchdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("bhchdhgl")){
                sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("bhchdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("szdjcds")){
                sheet.getRow(index).getCell(10).setCellValue(Double.valueOf(datum.get("szdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(10).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("szdhgds")){
                sheet.getRow(index).getCell(11).setCellValue(Double.valueOf(datum.get("szdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(11).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("szdhgl")){
                sheet.getRow(index).getCell(12).setCellValue(Double.valueOf(datum.get("szdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(12).setCellValue(Double.valueOf(0));
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelxmccData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.2-24");
        int index = 4;
        for (Map<String, Object> datum : data) {
            sheet.getRow(2).getCell(index).setCellValue(datum.get("htd").toString());

            if (datum.containsKey("ysdzds")) {
                sheet.getRow(3).getCell(index).setCellValue(Double.valueOf(datum.get("ysdzds").toString()));
            } else {
                sheet.getRow(3).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("ysdhgds")) {
                sheet.getRow(4).getCell(index).setCellValue(Double.valueOf(datum.get("ysdhgds").toString()));
            } else {
                sheet.getRow(4).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("ysdhgl")) {
                sheet.getRow(5).getCell(index).setCellValue(Double.valueOf(datum.get("ysdhgl").toString()));
            } else {
                sheet.getRow(5).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("wczds")) {
                sheet.getRow(6).getCell(index).setCellValue(Double.valueOf(datum.get("wczds").toString()));
            } else {
                sheet.getRow(6).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("wchgds")) {
                sheet.getRow(7).getCell(index).setCellValue(Double.valueOf(datum.get("wchgds").toString()));
            } else {
                sheet.getRow(7).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("wchgl")) {
                sheet.getRow(8).getCell(index).setCellValue(Double.valueOf(datum.get("wchgl").toString()));
            } else {
                sheet.getRow(8).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("czzds")) {
                sheet.getRow(9).getCell(index).setCellValue(Double.valueOf(datum.get("czzds").toString()));
            } else {
                sheet.getRow(9).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("czhgds")) {
                sheet.getRow(10).getCell(index).setCellValue(Double.valueOf(datum.get("czhgds").toString()));
            } else {
                sheet.getRow(10).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("czhgl")) {
                sheet.getRow(11).getCell(index).setCellValue(Double.valueOf(datum.get("czhgl").toString()));
            } else {
                sheet.getRow(11).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("ssxszds")){
                sheet.getRow(12).getCell(index).setCellValue(Double.valueOf(datum.get("ssxszds").toString()));
            }else {
                sheet.getRow(12).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("ssxshgds")){
                sheet.getRow(13).getCell(index).setCellValue(Double.valueOf(datum.get("ssxshgds").toString()));
            }else {
                sheet.getRow(13).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("ssxshgl")){
                sheet.getRow(14).getCell(index).setCellValue(Double.valueOf(datum.get("ssxshgl").toString()));
            }else {
                sheet.getRow(14).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("lqpzdzds")){
                sheet.getRow(15).getCell(index).setCellValue(Double.valueOf(datum.get("lqpzdzds").toString()));
            }else {
                sheet.getRow(15).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("lqpzdhgds")){
                sheet.getRow(16).getCell(index).setCellValue(Double.valueOf(datum.get("lqpzdhgds").toString()));
            }else {
                sheet.getRow(16).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("lqpzdhgl")){
                sheet.getRow(17).getCell(index).setCellValue(Double.valueOf(datum.get("lqpzdhgl").toString()));
            }else {
                sheet.getRow(17).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("mcxszds")){
                sheet.getRow(18).getCell(index).setCellValue(Double.valueOf(datum.get("mcxszds").toString()));
            }else {
                sheet.getRow(18).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("mcxshgds")){
                sheet.getRow(19).getCell(index).setCellValue(Double.valueOf(datum.get("mcxshgds").toString()));
            }else {
                sheet.getRow(19).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("mcxshgl")){
                sheet.getRow(20).getCell(index).setCellValue(Double.valueOf(datum.get("mcxshgl").toString()));
            }else {
                sheet.getRow(20).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("gzsdzds")){
                sheet.getRow(21).getCell(index).setCellValue(Double.valueOf(datum.get("gzsdzds").toString()));
            }else {
                sheet.getRow(21).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("gzsdhgds")){
                sheet.getRow(22).getCell(index).setCellValue(Double.valueOf(datum.get("gzsdhgds").toString()));
            }else {
                sheet.getRow(22).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("gzsdhgl")){
                sheet.getRow(23).getCell(index).setCellValue(Double.valueOf(datum.get("gzsdhgl").toString()));
            }else {
                sheet.getRow(23).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("lqhdzds")){
                sheet.getRow(24).getCell(index).setCellValue(Double.valueOf(datum.get("lqhdzds").toString()));
            }else {
                sheet.getRow(24).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("lqhdhgds")){
                sheet.getRow(25).getCell(index).setCellValue(Double.valueOf(datum.get("lqhdhgds").toString()));
            }else {
                sheet.getRow(25).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("lqhdhgl")){
                sheet.getRow(26).getCell(index).setCellValue(Double.valueOf(datum.get("lqhdhgl").toString()));
            }else {
                sheet.getRow(26).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("lqhpzds")){
                sheet.getRow(27).getCell(index).setCellValue(Double.valueOf(datum.get("lqhpzds").toString()));
            }else {
                sheet.getRow(27).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("lqhphgds")){
                sheet.getRow(28).getCell(index).setCellValue(Double.valueOf(datum.get("lqhphgds").toString()));
            }else {
                sheet.getRow(28).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("lqhphgl")){
                sheet.getRow(29).getCell(index).setCellValue(Double.valueOf(datum.get("lqhphgl").toString()));
            }else {
                sheet.getRow(29).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("hntqdzds")){
                sheet.getRow(30).getCell(index).setCellValue(Double.valueOf(datum.get("hntqdzds").toString()));
            }else {
                sheet.getRow(30).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("hntqdhgds")){
                sheet.getRow(31).getCell(index).setCellValue(Double.valueOf(datum.get("hntqdhgds").toString()));
            }else {
                sheet.getRow(31).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("hntqdhgl")){
                sheet.getRow(32).getCell(index).setCellValue(Double.valueOf(datum.get("hntqdhgl").toString()));
            }else {
                sheet.getRow(32).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("tlmxlbgczds")){
                sheet.getRow(33).getCell(index).setCellValue(Double.valueOf(datum.get("tlmxlbgczds").toString()));
            }else {
                sheet.getRow(33).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("tlmxlbgchgds")){
                sheet.getRow(34).getCell(index).setCellValue(Double.valueOf(datum.get("tlmxlbgchgds").toString()));
            }else {
                sheet.getRow(34).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("tlmxlbgchgl")){
                sheet.getRow(35).getCell(index).setCellValue(Double.valueOf(datum.get("tlmxlbgchgl").toString()));
            }else {
                sheet.getRow(35).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("hntpzdzds")){
                sheet.getRow(36).getCell(index).setCellValue(Double.valueOf(datum.get("hntpzdzds").toString()));
            }else {
                sheet.getRow(36).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("hntpzdhgds")){
                sheet.getRow(37).getCell(index).setCellValue(Double.valueOf(datum.get("hntpzdhgds").toString()));
            }else {
                sheet.getRow(37).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("hntpzdhgl")){
                sheet.getRow(38).getCell(index).setCellValue(Double.valueOf(datum.get("hntpzdhgl").toString()));
            }else {
                sheet.getRow(38).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("khzds")){
                sheet.getRow(39).getCell(index).setCellValue(Double.valueOf(datum.get("khzds").toString()));
            }else {
                sheet.getRow(39).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("khhgds")){
                sheet.getRow(40).getCell(index).setCellValue(Double.valueOf(datum.get("khhgds").toString()));
            }else {
                sheet.getRow(40).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("khhgl")){
                sheet.getRow(41).getCell(index).setCellValue(Double.valueOf(datum.get("khhgl").toString()));
            }else {
                sheet.getRow(41).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("hnthdzds") && datum.get("hnthdzds") !=null){
                sheet.getRow(42).getCell(index).setCellValue(Double.valueOf(datum.get("hnthdzds").toString()));
            }else {
                sheet.getRow(42).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("hnthgds")&& datum.get("hnthgds") !=null){
                sheet.getRow(43).getCell(index).setCellValue(Double.valueOf(datum.get("hnthgds").toString()));
            }else {
                sheet.getRow(43).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("hnthgl")&& datum.get("hnthgl") !=null){
                sheet.getRow(44).getCell(index).setCellValue(Double.valueOf(datum.get("hnthgl").toString()));
            }else {
                sheet.getRow(44).getCell(index).setCellValue(0);
            }


            if (datum.containsKey("hnthpzds")){
                sheet.getRow(45).getCell(index).setCellValue(Double.valueOf(datum.get("hnthpzds").toString()));
            }else {
                sheet.getRow(45).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("hnthphgds")){
                sheet.getRow(46).getCell(index).setCellValue(Double.valueOf(datum.get("hnthphgds").toString()));
            }else {
                sheet.getRow(46).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("hnthphdl")){
                sheet.getRow(47).getCell(index).setCellValue(Double.valueOf(datum.get("hnthphdl").toString()));
            }else {
                sheet.getRow(47).getCell(index).setCellValue(0);
            }
            index++;

        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdlmhzData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.2-23");
        int index = 4;
        for (Map<String, Object> datum : data) {
            sheet.getRow(2).getCell(index).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("ysdzs")){
                sheet.getRow(3).getCell(index).setCellValue(Double.valueOf(datum.get("ysdzs").toString()));
            }else {
                sheet.getRow(3).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("ysdhgs")){
                sheet.getRow(4).getCell(index).setCellValue(Double.valueOf(datum.get("ysdhgs").toString()));
            }else {
                sheet.getRow(4).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("ysdhgl")){
                sheet.getRow(5).getCell(index).setCellValue(Double.valueOf(datum.get("ysdhgl").toString()));
            }else {
                sheet.getRow(5).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("sdczjcds")){
                sheet.getRow(6).getCell(index).setCellValue(Double.valueOf(datum.get("sdczjcds").toString()));
            }else {
                sheet.getRow(6).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("sdczhgds")){
                sheet.getRow(7).getCell(index).setCellValue(Double.valueOf(datum.get("sdczhgds").toString()));
            }else {
                sheet.getRow(7).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("sdczhgl")){
                sheet.getRow(8).getCell(index).setCellValue(Double.valueOf(datum.get("sdczhgl").toString()));
            }else {
                sheet.getRow(8).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("sdssxsjcds")){
                sheet.getRow(9).getCell(index).setCellValue(Double.valueOf(datum.get("sdssxsjcds").toString()));
            }else {
                sheet.getRow(9).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("sdssxshgds")){
                sheet.getRow(10).getCell(index).setCellValue(Double.valueOf(datum.get("sdssxshgds").toString()));
            }else {
                sheet.getRow(10).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("sdssxshgl")){
                sheet.getRow(11).getCell(index).setCellValue(Double.valueOf(datum.get("sdssxshgl").toString()));
            }else {
                sheet.getRow(11).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("sdpzdlqjcds")){
                sheet.getRow(12).getCell(index).setCellValue(Double.valueOf(datum.get("sdpzdlqjcds").toString()));
            }else {
                sheet.getRow(12).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("sdpzdlqhgds")){
                sheet.getRow(13).getCell(index).setCellValue(Double.valueOf(datum.get("sdpzdlqhgds").toString()));
            }else {
                sheet.getRow(13).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("sdpzdlqhgl")){
                sheet.getRow(14).getCell(index).setCellValue(Double.valueOf(datum.get("sdpzdlqhgl").toString()));
            }else {
                sheet.getRow(14).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("khlqmcxsjcds")){
                sheet.getRow(15).getCell(index).setCellValue(Double.valueOf(datum.get("khlqmcxsjcds").toString()));
            }else {
                sheet.getRow(15).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("khlqmcxshgds")){
                sheet.getRow(16).getCell(index).setCellValue(Double.valueOf(datum.get("khlqmcxshgds").toString()));
            }else {
                sheet.getRow(16).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("khlqmcxshgl")){
                sheet.getRow(17).getCell(index).setCellValue(Double.valueOf(datum.get("khlqmcxshgl").toString()));
            }else {
                sheet.getRow(17).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("khlqgzsdjcds")){
                sheet.getRow(18).getCell(index).setCellValue(Double.valueOf(datum.get("khlqgzsdjcds").toString()));
            }else {
                sheet.getRow(18).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("khlqgzsdhgds")){
                sheet.getRow(19).getCell(index).setCellValue(Double.valueOf(datum.get("khlqgzsdhgds").toString()));
            }else {
                sheet.getRow(19).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("khlqgzsdhgl")){
                sheet.getRow(20).getCell(index).setCellValue(Double.valueOf(datum.get("khlqgzsdhgl").toString()));
            }else {
                sheet.getRow(20).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("hdlqzds")){
                sheet.getRow(21).getCell(index).setCellValue(Double.valueOf(datum.get("hdlqzds").toString()));
            }else {
                sheet.getRow(21).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("hdlqhgds")){
                sheet.getRow(22).getCell(index).setCellValue(Double.valueOf(datum.get("hdlqhgds").toString()));
            }else {
                sheet.getRow(22).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("hdlqhgl")){
                sheet.getRow(23).getCell(index).setCellValue(Double.valueOf(datum.get("hdlqhgl").toString()));
            }else {
                sheet.getRow(23).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("hplqzds")){
                sheet.getRow(24).getCell(index).setCellValue(Double.valueOf(datum.get("hplqzds").toString()));
            }else {
                sheet.getRow(24).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("hplqhgds")){
                sheet.getRow(25).getCell(index).setCellValue(Double.valueOf(datum.get("hplqhgds").toString()));
            }else {
                sheet.getRow(25).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("hplqhgl")){
                sheet.getRow(26).getCell(index).setCellValue(Double.valueOf(datum.get("hplqhgl").toString()));
            }else {
                sheet.getRow(26).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("sdpzdhntjcds")){
                sheet.getRow(33).getCell(index).setCellValue(Double.valueOf(datum.get("sdpzdhntjcds").toString()));
            }else {
                sheet.getRow(33).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("sdpzdhnthgds")){
                sheet.getRow(34).getCell(index).setCellValue(Double.valueOf(datum.get("sdpzdhnthgds").toString()));
            }else {
                sheet.getRow(34).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("sdpzdhnthgl")){
                sheet.getRow(35).getCell(index).setCellValue(Double.valueOf(datum.get("sdpzdhnthgl").toString()));
            }else {
                sheet.getRow(35).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("khhntmcxsjcds")){
                sheet.getRow(36).getCell(index).setCellValue(Double.valueOf(datum.get("khhntmcxsjcds").toString()));
            }else {
                sheet.getRow(36).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("khhntmcxshgds")){
                sheet.getRow(37).getCell(index).setCellValue(Double.valueOf(datum.get("khhntmcxshgds").toString()));
            }else {
                sheet.getRow(37).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("khhntmcxshgl")){
                sheet.getRow(38).getCell(index).setCellValue(Double.valueOf(datum.get("khhntmcxshgl").toString()));
            }else {
                sheet.getRow(38).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("khhntgzsdjcds")){
                sheet.getRow(39).getCell(index).setCellValue(Double.valueOf(datum.get("khhntgzsdjcds").toString()));
            }else {
                sheet.getRow(39).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("khhntgzsdhgds")){
                sheet.getRow(40).getCell(index).setCellValue(Double.valueOf(datum.get("khhntgzsdhgds").toString()));
            }else {
                sheet.getRow(40).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("khhntgzsdhgl")){
                sheet.getRow(41).getCell(index).setCellValue(Double.valueOf(datum.get("khhntgzsdhgl").toString()));
            }else {
                sheet.getRow(41).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("hdhntzds")){
                sheet.getRow(42).getCell(index).setCellValue(Double.valueOf(datum.get("hdhntzds").toString()));
            }else {
                sheet.getRow(42).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("hdhnthgds")){
                sheet.getRow(43).getCell(index).setCellValue(Double.valueOf(datum.get("hdhnthgds").toString()));
            }else {
                sheet.getRow(43).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("hdhnthgl")){
                sheet.getRow(44).getCell(index).setCellValue(Double.valueOf(datum.get("hdhnthgl").toString()));
            }else {
                sheet.getRow(44).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("hphntzds")){
                sheet.getRow(45).getCell(index).setCellValue(Double.valueOf(datum.get("hphntzds").toString()));
            }else {
                sheet.getRow(45).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("hphnthgds")){
                sheet.getRow(46).getCell(index).setCellValue(Double.valueOf(datum.get("hphnthgds").toString()));
            }else {
                sheet.getRow(46).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("hphnthgl")){
                sheet.getRow(47).getCell(index).setCellValue(Double.valueOf(datum.get("hphnthgl").toString()));
            }else {
                sheet.getRow(47).getCell(index).setCellValue(0);
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdhpData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-22");
        XSSFSheet sheet = wb.getSheet("表4.1.2-22");
        int index = 3;
        for (Map<String, Object> datum : data) {

            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("sdhplmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("sdhplmlx").toString());
            }
            if (datum.containsKey("sdhpzds")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("sdhpzds").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("sdhphgds")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("sdhphgds").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("sdhphgl")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("sdhphgl").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(0));
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdhdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-21(3)");
        XSSFSheet sheet = wb.getSheet("表4.1.2-21(3)");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("sdlmhdlmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("sdlmhdlmlx").toString());
            }

            if (datum.containsKey("sdlmhdjcds")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("sdlmhdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("sdlmhdhgs")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("sdlmhdhgs").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("sdlmhdhgl")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("sdlmhdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(0));
            }

            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdldhdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-21(2)");
        XSSFSheet sheet = wb.getSheet("表4.1.2-21(2)");
        int index = 3;
        for (Map<String, Object> datum : data) {

            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("sdldhdlmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("sdldhdlmlx").toString());
            }
            if (datum.containsKey("sdldhdlb")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("sdldhdlb").toString());
            }
            if (datum.containsKey("sdldhdbhfw")){
                sheet.getRow(index).getCell(3).setCellValue(datum.get("sdldhdbhfw").toString());
            }
            if (datum.containsKey("sdldhdsjz")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("sdldhdsjz").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("sdldhdjcds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("sdldhdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("sdldhdhgds")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("sdldhdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("sdldhdhgl")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("sdldhdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(0));
            }

            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdzxfhdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.2-21(1)");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("sdzxfhdlmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("sdzxfhdlmlx").toString());
            }
            if (datum.containsKey("sdzxfhdlb")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("sdzxfhdlb").toString());

            }
            if (datum.containsKey("sdzxfhdpjzfw")){
                sheet.getRow(index).getCell(3).setCellValue(datum.get("sdzxfhdpjzfw").toString());
            }
            if (datum.containsKey("sdzxfhdzdbz")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("sdzxfhdzdbz").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("sdzxfhdydbz")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("sdzxfhdydbz").toString()));

            }else {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("sdzxfhdsjz")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("sdzxfhdsjz").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("sdzxfhdjcds")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("sdzxfhdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("sdzxfhdhgs")){
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("sdzxfhdhgs").toString()));
            }else {
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("sdzxfhdhgl")){
                sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("sdzxfhdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(0));
            }
            index++;
        }

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdkhData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-20");
        XSSFSheet sheet = wb.getSheet("表4.1.2-20");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("sdkhlmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("sdkhlmlx").toString());
            }
            if (datum.containsKey("sdkhkpzb")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("sdkhkpzb").toString());
            }
            if (datum.containsKey("sdkhsjz")){
                sheet.getRow(index).getCell(3).setCellValue(datum.get("sdkhsjz").toString());
            }
            if (datum.containsKey("sdkhbhfw")){
                sheet.getRow(index).getCell(4).setCellValue(datum.get("sdkhbhfw").toString());
            }

            if (datum.containsKey("sdkhjcds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("sdkhjcds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(0));

            }
            if (datum.containsKey("sdkhhgds")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("sdkhhgds").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("sdkhhgl")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("sdkhhgl").toString()));
            }else {
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(0));
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdpzdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-19");
        XSSFSheet sheet = wb.getSheet("表4.1.2-19");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("sdpzdzb")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("sdpzdzb").toString());
            }
            if (datum.containsKey("sdpzdlmlx")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("sdpzdlmlx").toString());
            }
            if (datum.containsKey("sdpzdgdz")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("sdpzdgdz").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(0);
            }
            if (datum.containsKey("sdpzdbhfw")){
                sheet.getRow(index).getCell(4).setCellValue(datum.get("sdpzdbhfw").toString());
            }
            if (datum.containsKey("sdpzdjcds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("sdpzdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }

            if (datum.containsKey("sdpzdhgds")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("sdpzdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(0);
            }
            if (datum.containsKey("sdpzdhgl")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("sdpzdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(7).setCellValue(0);
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdhntxlbgcData(XSSFWorkbook wb, List<Map<String, Object>> data) {

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdhntqdData(XSSFWorkbook wb, List<Map<String, Object>> data) {

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdssxsData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-16");
        XSSFSheet sheet = wb.getSheet("表4.1.2-16");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("sdssxsgdz")){
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("sdssxsgdz").toString()));
            }else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }

            if (datum.containsKey("sdssxsmin")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("sdssxsmin").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(0);
            }

            if (datum.containsKey("sdssxsmax")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("sdssxsmax").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(0);
            }

            if (datum.containsKey("sdssxsssjcds")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("sdssxsssjcds").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(0);
            }
            if (datum.containsKey("sdssxssshgds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("sdssxssshgds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }
            if (datum.containsKey("sdssxssshgl")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("sdssxssshgl").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(0);
            }
            index++;
        }

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdczData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-15(2)");
        XSSFSheet sheet = wb.getSheet("表4.1.2-15(2)");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("sdczzb")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("sdczzb").toString());
            }
            if (datum.containsKey("sdczlmlx")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("sdczlmlx").toString());
            }
            if (datum.containsKey("sdczgdz")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("sdczgdz").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(0);
            }
            if (datum.containsKey("sdczbhfw")){
                sheet.getRow(index).getCell(4).setCellValue(datum.get("sdczbhfw").toString());
            }
            if (datum.containsKey("sdczzds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("sdczzds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }
            if (datum.containsKey("sdczhgds")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("sdczhgds").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(0);
            }
            if (datum.containsKey("sdczhgl")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("sdczhgl").toString()));
            }else {
                sheet.getRow(index).getCell(7).setCellValue(0);
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelsdysdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        System.out.println(data);
        createRow(wb,data.size(),"表4.1.2-15");
        XSSFSheet sheet = wb.getSheet("表4.1.2-15");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("ysdlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("ysdlx").toString());
            }
            if (datum.containsKey("ysdsczbh")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("ysdsczbh").toString());
            }

            if (datum.containsKey("ysdzdbz")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("ysdzdbz").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(0);
            }

            if (datum.containsKey("ysdydbz")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("ysdydbz").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(0);
            }
            if (datum.containsKey("ysdgdz")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("ysdgdz").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }

            if (datum.containsKey("ysdzs")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("ysdzs").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(0);
            }

            if (datum.containsKey("ysdhgs")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("ysdhgs").toString()));
            }else {
                sheet.getRow(index).getCell(7).setCellValue(0);
            }

            if (datum.containsKey("ysdhgl")){
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("ysdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(8).setCellValue(0);
            }
            index ++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelqmxhzData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.2-14");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(2).getCell(index).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("qmpzdjcds")){
                sheet.getRow(3).getCell(index).setCellValue(Double.valueOf(datum.get("qmpzdjcds").toString()));
            }else {
                sheet.getRow(3).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("qmpzdhgds")){
                sheet.getRow(4).getCell(index).setCellValue(Double.valueOf(datum.get("qmpzdhgds").toString()));
            }else {
                sheet.getRow(4).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("qmpzdhgl")){
                sheet.getRow(5).getCell(index).setCellValue(Double.valueOf(datum.get("qmpzdhgl").toString()));
            }else {
                sheet.getRow(5).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("qmhpzds")){
                sheet.getRow(6).getCell(index).setCellValue(Double.valueOf(datum.get("qmhpzds").toString()));
            }else {
                sheet.getRow(6).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("qmhphgds")){
                sheet.getRow(7).getCell(index).setCellValue(Double.valueOf(datum.get("qmhphgds").toString()));
            }else {
                sheet.getRow(7).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("qmhphgl")){
                sheet.getRow(8).getCell(index).setCellValue(Double.valueOf(datum.get("qmhphgl").toString()));
            }else {
                sheet.getRow(8).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("mcxszds")){
                sheet.getRow(9).getCell(index).setCellValue(Double.valueOf(datum.get("mcxszds").toString()));
            }else {
                sheet.getRow(9).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("mcxshgds")){
                sheet.getRow(10).getCell(index).setCellValue(Double.valueOf(datum.get("mcxshgds").toString()));
            }else {
                sheet.getRow(10).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("mcxshgl")){
                sheet.getRow(11).getCell(index).setCellValue(Double.valueOf(datum.get("mcxshgl").toString()));
            }else {
                sheet.getRow(11).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("gzsdzds")){
                sheet.getRow(12).getCell(index).setCellValue(Double.valueOf(datum.get("gzsdzds").toString()));
            }else {
                sheet.getRow(12).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("gzsdhgds")){
                sheet.getRow(13).getCell(index).setCellValue(Double.valueOf(datum.get("gzsdhgds").toString()));
            }else {
                sheet.getRow(13).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("gzsdhgl")){
                sheet.getRow(14).getCell(index).setCellValue(Double.valueOf(datum.get("gzsdhgl").toString()));
            }else {
                sheet.getRow(14).getCell(index).setCellValue(0);
            }
        }
    }


    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelqmkhData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.2-13");
        int index = 3;

        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("qmkhlmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("qmkhlmlx").toString());
            }
            if (datum.containsKey("qmkhzb")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("qmkhzb").toString());
            }
            if (datum.containsKey("qmkhsjz")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("qmkhsjz").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(0);
            }
            if (datum.containsKey("qmkhbhfw")){
                sheet.getRow(index).getCell(4).setCellValue(datum.get("qmkhbhfw").toString());
            }

            if (datum.containsKey("qmkhjcds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("qmkhjcds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }

            if (datum.containsKey("qmkhhgds")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("qmkhhgds").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(0);
            }

            if (datum.containsKey("qmkhhgl")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("qmkhhgl").toString()));
            }else {
                sheet.getRow(index).getCell(7).setCellValue(0);
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelqmhpdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.2-12");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("qmhplmlx")) {
                sheet.getRow(index).getCell(1).setCellValue(datum.get("qmhplmlx").toString());
            }
            if (datum.containsKey("qmhpzds")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("qmhpzds").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(0);
            }
            if (datum.containsKey("qmhphgds")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("qmhphgds").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(0);
            }

            if (datum.containsKey("qmhphgl")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("qmhphgl").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(0);
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelqmpzdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.2-11");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("qmpzdzb")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("qmpzdzb").toString());
            }
            if (datum.containsKey("qmpzdlmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("qmpzdlmlx").toString());
            }
            if (datum.containsKey("qmpzdgdz")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("qmpzdgdz").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(0);
            }
            if (datum.containsKey("qmpzdbhfw")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("qmpzdbhfw").toString()));
            }

            if (datum.containsKey("qmpzdjcds")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("qmpzdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(0);
            }

            if (datum.containsKey("qmpzdhgds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("qmpzdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }

            if (datum.containsKey("qmpzdhgl")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("qmpzdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(0);
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmhzData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.2-10");
        int index = 4;
        for (Map<String, Object> datum : data) {
            sheet.getRow(2).getCell(index).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("ysdzds")){
                sheet.getRow(3).getCell(index).setCellValue(Double.valueOf(datum.get("ysdzds").toString()));
            }else {
                sheet.getRow(3).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("ysdhgds")){
                sheet.getRow(4).getCell(index).setCellValue(Double.valueOf(datum.get("ysdhgds").toString()));
            }else {
                sheet.getRow(4).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("ysdhgl")){
                sheet.getRow(5).getCell(index).setCellValue(Double.valueOf(datum.get("ysdhgl").toString()));
            }else {
                sheet.getRow(5).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("lmwczds")){
                sheet.getRow(6).getCell(index).setCellValue(Double.valueOf(datum.get("lmwczds").toString()));
            }else {
                sheet.getRow(6).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("lmwchgds")){
                sheet.getRow(7).getCell(index).setCellValue(Double.valueOf(datum.get("lmwchgds").toString()));
            }else {
                sheet.getRow(7).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("lmwchgl")){
                sheet.getRow(8).getCell(index).setCellValue(Double.valueOf(datum.get("lmwchgl").toString()));
            }else {
                sheet.getRow(8).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("czzds")){
                sheet.getRow(9).getCell(index).setCellValue(Double.valueOf(datum.get("czzds").toString()));
            }else {
                sheet.getRow(9).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("czhgds")){
                sheet.getRow(10).getCell(index).setCellValue(Double.valueOf(datum.get("czhgds").toString()));
            }else {
                sheet.getRow(10).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("czhgl")){
                sheet.getRow(11).getCell(index).setCellValue(Double.valueOf(datum.get("czhgl").toString()));
            }else {
                sheet.getRow(11).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("lmssxsssjcds")){
                sheet.getRow(12).getCell(index).setCellValue(Double.valueOf(datum.get("lmssxsssjcds").toString()));
            }else {
                sheet.getRow(12).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("lmssxssshgds")){
                sheet.getRow(13).getCell(index).setCellValue(Double.valueOf(datum.get("lmssxssshgds").toString()));
            }else {
                sheet.getRow(13).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("lmssxssshgl")){
                sheet.getRow(14).getCell(index).setCellValue(Double.valueOf(datum.get("lmssxssshgl").toString()));
            }else {
                sheet.getRow(14).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("pzdlqjcds")){
                sheet.getRow(15).getCell(index).setCellValue(Double.valueOf(datum.get("pzdlqjcds").toString()));
            }else {
                sheet.getRow(15).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("pzdlqhgds")){
                sheet.getRow(16).getCell(index).setCellValue(Double.valueOf(datum.get("pzdlqhgds").toString()));
            }else {
                sheet.getRow(16).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("pzdlqhgl")){
                sheet.getRow(17).getCell(index).setCellValue(Double.valueOf(datum.get("pzdlqhgds").toString()));
            }else {
                sheet.getRow(17).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("khlqmcxsjcds")){
                sheet.getRow(18).getCell(index).setCellValue(Double.valueOf(datum.get("khlqmcxsjcds").toString()));
            }else {
                sheet.getRow(18).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("khlqmcxshgds")){
                sheet.getRow(19).getCell(index).setCellValue(Double.valueOf(datum.get("khlqmcxshgds").toString()));
            }else {
                sheet.getRow(19).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("khlqmcxshgl")){
                sheet.getRow(20).getCell(index).setCellValue(Double.valueOf(datum.get("khlqmcxshgl").toString()));
            }else {
                sheet.getRow(20).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("khlqgzsdjcds")){
                sheet.getRow(21).getCell(index).setCellValue(Double.valueOf(datum.get("khlqgzsdjcds").toString()));
            }else {
                sheet.getRow(21).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("khlqgzsdhgds")){
                sheet.getRow(22).getCell(index).setCellValue(Double.valueOf(datum.get("khlqgzsdhgds").toString()));
            }else {
                sheet.getRow(22).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("khlqgzsdhgl")){
                sheet.getRow(23).getCell(index).setCellValue(Double.valueOf(datum.get("khlqgzsdhgl").toString()));
            }else {
                sheet.getRow(23).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("lmhdlqjcds")){
                sheet.getRow(24).getCell(index).setCellValue(Double.valueOf(datum.get("lmhdlqjcds").toString()));
            }else {
                sheet.getRow(24).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("lmhdlqhgs")){
                sheet.getRow(25).getCell(index).setCellValue(Double.valueOf(datum.get("lmhdlqhgs").toString()));
            }else {
                sheet.getRow(25).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("lmhdlqhgl")){
                sheet.getRow(26).getCell(index).setCellValue(Double.valueOf(datum.get("lmhdlqhgl").toString()));
            }else {
                sheet.getRow(26).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("hplqzds")){
                sheet.getRow(27).getCell(index).setCellValue(Double.valueOf(datum.get("hplqzds").toString()));
            }else {
                sheet.getRow(27).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("hplqhgds")){
                sheet.getRow(28).getCell(index).setCellValue(Double.valueOf(datum.get("hplqhgds").toString()));
            }else {
                sheet.getRow(28).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("hplqhgl")){
                sheet.getRow(29).getCell(index).setCellValue(Double.valueOf(datum.get("hplqhgl").toString()));
            }else {
                sheet.getRow(29).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("hntqdzds")){
                sheet.getRow(30).getCell(index).setCellValue(Double.valueOf(datum.get("hntqdzds").toString()));
            }else {
                sheet.getRow(30).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("hntqdhgds")){
                sheet.getRow(31).getCell(index).setCellValue(Double.valueOf(datum.get("hntqdhgds").toString()));
            }else {
                sheet.getRow(31).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("hntqdhgl")){
                sheet.getRow(32).getCell(index).setCellValue(Double.valueOf(datum.get("hntqdhgl").toString()));
            }else {
                sheet.getRow(32).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("hntxlbgczds")){
                sheet.getRow(33).getCell(index).setCellValue(Double.valueOf(datum.get("hntxlbgczds").toString()));
            }else {
                sheet.getRow(33).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("hntxlbgchgds")){
                sheet.getRow(34).getCell(index).setCellValue(Double.valueOf(datum.get("hntxlbgchgds").toString()));
            }else {
                sheet.getRow(34).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("hntxlbgchgl")){
                sheet.getRow(35).getCell(index).setCellValue(Double.valueOf(datum.get("hntxlbgchgl").toString()));
            }else {
                sheet.getRow(35).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("pzdhntjcds")){
                sheet.getRow(36).getCell(index).setCellValue(Double.valueOf(datum.get("pzdhntjcds").toString()));
            }else {
                sheet.getRow(36).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("pzdhnthgds")){
                sheet.getRow(37).getCell(index).setCellValue(Double.valueOf(datum.get("pzdhnthgds").toString()));
            }else {
                sheet.getRow(37).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("pzdhnthgl")){
                sheet.getRow(38).getCell(index).setCellValue(Double.valueOf(datum.get("pzdhnthgl").toString()));
            }else {
                sheet.getRow(38).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("khhntmcxsjcds")){
                sheet.getRow(39).getCell(index).setCellValue(Double.valueOf(datum.get("khhntmcxsjcds").toString()));
            }else {
                sheet.getRow(39).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("khhntmcxshgds")){
                sheet.getRow(40).getCell(index).setCellValue(Double.valueOf(datum.get("khhntmcxshgds").toString()));
            }else {
                sheet.getRow(40).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("khhntmcxshgl")){
                sheet.getRow(41).getCell(index).setCellValue(Double.valueOf(datum.get("khhntmcxshgl").toString()));
            }else {
                sheet.getRow(41).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("khhntgzsdjcds")){
                sheet.getRow(42).getCell(index).setCellValue(Double.valueOf(datum.get("khhntgzsdjcds").toString()));
            }else {
                sheet.getRow(42).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("khhntgzsdhgds")){
                sheet.getRow(43).getCell(index).setCellValue(Double.valueOf(datum.get("khhntgzsdhgds").toString()));
            }else {
                sheet.getRow(43).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("khhntgzsdhgl")){
                sheet.getRow(44).getCell(index).setCellValue(Double.valueOf(datum.get("khhntgzsdhgl").toString()));
            }else {
                sheet.getRow(44).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("lmhdhntjcds")){
                sheet.getRow(45).getCell(index).setCellValue(Double.valueOf(datum.get("lmhdhntjcds").toString()));
            }else {
                sheet.getRow(45).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("lmhdhnthgs")){
                sheet.getRow(46).getCell(index).setCellValue(Double.valueOf(datum.get("lmhdhnthgs").toString()));
            }else {
                sheet.getRow(46).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("lmhdhnthgl")){
                sheet.getRow(47).getCell(index).setCellValue(Double.valueOf(datum.get("lmhdhnthgl").toString()));
            }else {
                sheet.getRow(47).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("hphntzds")){
                sheet.getRow(48).getCell(index).setCellValue(Double.valueOf(datum.get("hphntzds").toString()));
            }else {
                sheet.getRow(48).getCell(index).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("hphnthgds")){
                sheet.getRow(49).getCell(index).setCellValue(Double.valueOf(datum.get("hphnthgds").toString()));
            }else {
                sheet.getRow(49).getCell(index).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("hphnthgl")){
                sheet.getRow(50).getCell(index).setCellValue(Double.valueOf(datum.get("hphnthgl").toString()));
            }else {
                sheet.getRow(50).getCell(index).setCellValue(Double.valueOf(0));
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmhpData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-9");
        XSSFSheet sheet = wb.getSheet("表4.1.2-9");
        int index = 3;
        for (Map<String, Object> datum : data) {

            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("hplmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("hplmlx").toString());
            }
            if (datum.containsKey("hpzds")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("hpzds").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("hphgds")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("hphgds").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("hphgl")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("hphgl").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(0));
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmhdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-8(3)");
        XSSFSheet sheet = wb.getSheet("表4.1.2-8(3)");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("lmhdlmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("lmhdlmlx").toString());
            }

            if (datum.containsKey("lmhdjcds")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("lmhdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("lmhdhgs")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("lmhdhgs").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("lmhdhgl")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("lmhdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(0));
            }

            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelldhdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-8(2)");
        XSSFSheet sheet = wb.getSheet("表4.1.2-8(2)");
        int index = 3;
        for (Map<String, Object> datum : data) {

            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("ldhdlmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("ldhdlmlx").toString());
            }
            if (datum.containsKey("ldhdlb")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("ldhdlb").toString());
            }
            if (datum.containsKey("ldhdbhfw")){
                sheet.getRow(index).getCell(3).setCellValue(datum.get("ldhdbhfw").toString());
            }
            if (datum.containsKey("ldhdsjz")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("ldhdsjz").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("ldhdjcds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("ldhdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("ldhdhgds")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("ldhdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("ldhdhgl")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("ldhdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(0));
            }

            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelzxfhdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        //createRow(wb,data.size(),"表4.1.2-8(1)");
        XSSFSheet sheet = wb.getSheet("表4.1.2-8(1)");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("zxfhdlmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("zxfhdlmlx").toString());
            }
            if (datum.containsKey("zxfhdlb")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("zxfhdlb").toString());

            }
            if (datum.containsKey("zxfhdpjzfw")){
                sheet.getRow(index).getCell(3).setCellValue(datum.get("zxfhdpjzfw").toString());
            }
            if (datum.containsKey("zxfhdzdbz")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("zxfhdzdbz").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("zxfhdydbz")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("zxfhdydbz").toString()));

            }else {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("zxfhdsjz")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("zxfhdsjz").toString()));

            }else {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("zxfhdjcds")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("zxfhdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("zxfhdhgs")){
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("zxfhdhgs").toString()));
            }else {
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(0));
            }

            if (datum.containsKey("zxfhdhgl")){
                sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("zxfhdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(0));
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelkhData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-7");
        XSSFSheet sheet = wb.getSheet("表4.1.2-7");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("khlmlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("khlmlx").toString());
            }
            if (datum.containsKey("khkpzb")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("khkpzb").toString());
            }
            if (datum.containsKey("khsjz")){
                sheet.getRow(index).getCell(3).setCellValue(datum.get("khsjz").toString());
            }
            if (datum.containsKey("khbhfw")){
                sheet.getRow(index).getCell(4).setCellValue(datum.get("khbhfw").toString());
            }

            if (datum.containsKey("khjcds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("khjcds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(0));

            }
            if (datum.containsKey("khhgds")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("khhgds").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(0));
            }
            if (datum.containsKey("khhgl")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("khhgl").toString()));
            }else {
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(0));
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmpzdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-6");
        XSSFSheet sheet = wb.getSheet("表4.1.2-6");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("pzdzb")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("pzdzb").toString());
            }
            if (datum.containsKey("pzdlmlx")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("pzdlmlx").toString());
            }
            if (datum.containsKey("pzdgdz")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("pzdgdz").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(0);
            }
            if (datum.containsKey("pzdbhfw")){
                sheet.getRow(index).getCell(4).setCellValue(datum.get("pzdbhfw").toString());
            }
            if (datum.containsKey("pzdjcds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("pzdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }

            if (datum.containsKey("pzdhgds")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("pzdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(0);
            }
            if (datum.containsKey("pzdhgl")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("pzdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(7).setCellValue(0);
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmhntxlbgcData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-5");
        XSSFSheet sheet = wb.getSheet("表4.1.2-5");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("hntxlbgcgdz")){
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("hntxlbgcgdz").toString()));
            }else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }
            if (datum.containsKey("hntxlbgcmin")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("hntxlbgcmin").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(0);
            }
            if (datum.containsKey("hntxlbgcmax")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("hntxlbgcmax").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(0);
            }
            if (datum.containsKey("hntxlbgczds")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("hntxlbgczds").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(0);
            }
            if (datum.containsKey("hntxlbgchgds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("hntxlbgchgds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }
            if (datum.containsKey("hntxlbgchgl")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("hntxlbgchgl").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(0);
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmhntqdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-4");
        XSSFSheet sheet = wb.getSheet("表4.1.2-4");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("hntqdgdz")){
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("hntqdgdz").toString()));
            }else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }
            if (datum.containsKey("hntqdmin")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("hntqdmin").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(0);
            }
            if (datum.containsKey("hntqdmax")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("hntqdmax").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(0);
            }
            if (datum.containsKey("hntqdpjz")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("hntqdpjz").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(0);
            }
            if (datum.containsKey("hntqdzds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("hntqdzds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }
            if (datum.containsKey("hntqdhgds")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("hntqdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(0);
            }
            if (datum.containsKey("hntqdhgl")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("hntqdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(7).setCellValue(0);
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmssxsData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-3");
        XSSFSheet sheet = wb.getSheet("表4.1.2-3");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("lmssxsgdz")){
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("lmssxsgdz").toString()));
            }else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }

            if (datum.containsKey("lmssxsmin")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("lmssxsmin").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(0);
            }

            if (datum.containsKey("lmssxsmax")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("lmssxsmax").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(0);
            }

            if (datum.containsKey("lmssxsssjcds")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("lmssxsssjcds").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(0);
            }
            if (datum.containsKey("lmssxssshgds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("lmssxssshgds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }
            if (datum.containsKey("lmssxssshgl")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("lmssxssshgl").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(0);
            }
            index++;
        }

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmczData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-2(2)");
        XSSFSheet sheet = wb.getSheet("表4.1.2-2(2)");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("czzb")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("czzb").toString());
            }
            if (datum.containsKey("czlmlx")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("czlmlx").toString());
            }
            if (datum.containsKey("czgdz")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("czgdz").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(0);
            }
            if (datum.containsKey("czbhfw")){
                sheet.getRow(index).getCell(4).setCellValue(datum.get("czbhfw").toString());
            }
            if (datum.containsKey("czzds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("czzds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }
            if (datum.containsKey("czhgds")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("czhgds").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(0);
            }
            if (datum.containsKey("czhgl")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("czhgl").toString()));
            }else {
                sheet.getRow(index).getCell(7).setCellValue(0);
            }
            index++;
        }


    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmwclcfdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-21");
        XSSFSheet sheet = wb.getSheet("表4.1.2-21");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("lmwclcfdbz")){
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("lmwclcfdbz").toString()));
            }else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }
            if (datum.containsKey("lmwclcfmax")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("lmwclcfmax").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(0);
            }

            if (datum.containsKey("lmwclcfmin")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("lmwclcfmin").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(0);
            }

            if (datum.containsKey("lmwclcfzds")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("lmwclcfzds").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(0);
            }
            if (datum.containsKey("lmwclcfhgds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("lmwclcfhgds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }
            if (datum.containsKey("lmwclcfhgl")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("lmwclcfhgl").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(0);
            }
            index++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmwcdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-211");
        XSSFSheet sheet = wb.getSheet("表4.1.2-211");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("lmwcdbz")){
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("lmwcdbz").toString()));
            }else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }

            if (datum.containsKey("lmwcmax")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("lmwcmax").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(0);
            }

            if (datum.containsKey("lmwcmin")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("lmwcmin").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(0);
            }

            if (datum.containsKey("lmwczds")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("lmwczds").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(0);
            }

            if (datum.containsKey("lmwchgds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("lmwchgds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }

            if (datum.containsKey("lmwchgl")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("lmwchgl").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(0);
            }
            index++;
        }

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcellmysdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.2-1");
        XSSFSheet sheet = wb.getSheet("表4.1.2-1");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("ysdlx")){
                sheet.getRow(index).getCell(1).setCellValue(datum.get("ysdlx").toString());
            }
            if (datum.containsKey("ysdsczbh")){
                sheet.getRow(index).getCell(2).setCellValue(datum.get("ysdsczbh").toString());
            }

            if (datum.containsKey("ysdzdbz")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("ysdzdbz").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(0);
            }

            if (datum.containsKey("ysdydbz")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("ysdydbz").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(0);
            }
            if (datum.containsKey("ysdgdz")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("ysdgdz").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }

            if (datum.containsKey("ysdzs")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("ysdzs").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(0);
            }

            if (datum.containsKey("ysdhgs")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("ysdhgs").toString()));
            }else {
                sheet.getRow(index).getCell(7).setCellValue(0);
            }

            if (datum.containsKey("ysdhgl")){
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("ysdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(8).setCellValue(0);
            }
            index ++;
        }


    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelljhzData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        XSSFSheet sheet = wb.getSheet("表4.1.1-6");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(2).getCell(index).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("ysdjcds")){
                sheet.getRow(3).getCell(index).setCellValue(Double.valueOf(datum.get("ysdjcds").toString()));
            }else {
                sheet.getRow(3).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("ysdhgds")){
                sheet.getRow(4).getCell(index).setCellValue(Double.valueOf(datum.get("ysdhgds").toString()));
            }else {
                sheet.getRow(4).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("ysdhgl")){
                sheet.getRow(5).getCell(index).setCellValue(Double.valueOf(datum.get("ysdhgl").toString()));
            }else {
                sheet.getRow(5).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("cjjcds")){
                sheet.getRow(6).getCell(index).setCellValue(Double.valueOf(datum.get("cjjcds").toString()));
            }else {
                sheet.getRow(6).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("cjhgds")){
                sheet.getRow(7).getCell(index).setCellValue(Double.valueOf(datum.get("cjhgds").toString()));
            }else {
                sheet.getRow(7).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("cjhgl")){
                sheet.getRow(8).getCell(index).setCellValue(Double.valueOf(datum.get("cjhgl").toString()));
            }else {
                sheet.getRow(8).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("wcjcds")){
                sheet.getRow(9).getCell(index).setCellValue(Double.valueOf(datum.get("wcjcds").toString()));
            }else {
                sheet.getRow(9).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("wchgds")){
                sheet.getRow(10).getCell(index).setCellValue(Double.valueOf(datum.get("wchgds").toString()));
            }else {
                sheet.getRow(10).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("wchgl")){
                sheet.getRow(11).getCell(index).setCellValue(Double.valueOf(datum.get("wchgl").toString()));
            }else {
                sheet.getRow(11).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("bpjcds")){
                sheet.getRow(12).getCell(index).setCellValue(Double.valueOf(datum.get("bpjcds").toString()));
            }else {
                sheet.getRow(12).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("bphgds")){
                sheet.getRow(13).getCell(index).setCellValue(Double.valueOf(datum.get("bphgds").toString()));
            }else {
                sheet.getRow(13).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("bphgl")){
                sheet.getRow(14).getCell(index).setCellValue(Double.valueOf(datum.get("bphgl").toString()));
            }else {
                sheet.getRow(14).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("psdmccjcds")){
                sheet.getRow(15).getCell(index).setCellValue(Double.valueOf(datum.get("psdmccjcds").toString()));
            }else {
                sheet.getRow(15).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("psdmcchgds")){
                sheet.getRow(16).getCell(index).setCellValue(Double.valueOf(datum.get("psdmcchgds").toString()));
            }else {
                sheet.getRow(16).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("psdmcchgl")){
                sheet.getRow(17).getCell(index).setCellValue(Double.valueOf(datum.get("psdmcchgl").toString()));
            }else {
                sheet.getRow(17).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("pspqhdjcds")){
                sheet.getRow(18).getCell(index).setCellValue(Double.valueOf(datum.get("pspqhdjcds").toString()));
            }else {
                sheet.getRow(18).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("pspqhdhgds")){
                sheet.getRow(19).getCell(index).setCellValue(Double.valueOf(datum.get("pspqhdhgds").toString()));
            }else {
                sheet.getRow(19).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("pspqhdhgl")){
                sheet.getRow(20).getCell(index).setCellValue(Double.valueOf(datum.get("pspqhdhgl").toString()));
            }else {
                sheet.getRow(20).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("xqtqdjcds")){
                sheet.getRow(21).getCell(index).setCellValue(Double.valueOf(datum.get("xqtqdjcds").toString()));
            }else {
                sheet.getRow(21).getCell(index).setCellValue(0);
            }
            if (datum.containsKey("xqtqdhgds")){
                sheet.getRow(22).getCell(index).setCellValue(Double.valueOf(datum.get("xqtqdhgds").toString()));
            }else {
                sheet.getRow(22).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("xqtqdhgl")){
                sheet.getRow(23).getCell(index).setCellValue(Double.valueOf(datum.get("xqtqdhgl").toString()));
            }else {
                sheet.getRow(23).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("xqjgccjcds")){
                sheet.getRow(24).getCell(index).setCellValue(Double.valueOf(datum.get("xqjgccjcds").toString()));
            }else {
                sheet.getRow(24).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("xqjgcchgds")){
                sheet.getRow(25).getCell(index).setCellValue(Double.valueOf(datum.get("xqjgcchgds").toString()));
            }else {
                sheet.getRow(25).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("xqjgcchgl")){
                sheet.getRow(26).getCell(index).setCellValue(Double.valueOf(datum.get("xqjgcchgl").toString()));
            }else {
                sheet.getRow(26).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("hdtqdjcds")){
                sheet.getRow(27).getCell(index).setCellValue(Double.valueOf(datum.get("hdtqdjcds").toString()));
            }else {
                sheet.getRow(27).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("hdtqdhgds")){
                sheet.getRow(28).getCell(index).setCellValue(Double.valueOf(datum.get("hdtqdhgds").toString()));
            }else {
                sheet.getRow(28).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("hdtqdhgl")){
                sheet.getRow(29).getCell(index).setCellValue(Double.valueOf(datum.get("hdtqdhgl").toString()));
            }else {
                sheet.getRow(29).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("hdjgccjcds")){
                sheet.getRow(30).getCell(index).setCellValue(Double.valueOf(datum.get("hdjgccjcds").toString()));
            }else {
                sheet.getRow(30).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("hdjgcchgds")){
                sheet.getRow(31).getCell(index).setCellValue(Double.valueOf(datum.get("hdjgcchgds").toString()));
            }else {
                sheet.getRow(31).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("hdjgcchgl")){
                sheet.getRow(32).getCell(index).setCellValue(Double.valueOf(datum.get("hdjgcchgl").toString()));
            }else {
                sheet.getRow(32).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("zdtqdjcds")){
                sheet.getRow(33).getCell(index).setCellValue(Double.valueOf(datum.get("zdtqdjcds").toString()));
            }else {
                sheet.getRow(33).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("zdtqdhgds")){
                sheet.getRow(34).getCell(index).setCellValue(Double.valueOf(datum.get("zdtqdhgds").toString()));
            }else {
                sheet.getRow(34).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("zdtqdhgl")){
                sheet.getRow(35).getCell(index).setCellValue(Double.valueOf(datum.get("zdtqdhgl").toString()));
            }else {
                sheet.getRow(35).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("zddmccjcds")){
                sheet.getRow(36).getCell(index).setCellValue(Double.valueOf(datum.get("zddmccjcds").toString()));
            }else {
                sheet.getRow(36).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("zddmcchgds")){
                sheet.getRow(37).getCell(index).setCellValue(Double.valueOf(datum.get("zddmcchgds").toString()));
            }else {
                sheet.getRow(37).getCell(index).setCellValue(0);
            }

            if (datum.containsKey("zddmcchgl")){
                sheet.getRow(38).getCell(index).setCellValue(Double.valueOf(datum.get("zddmcchgl").toString()));
            }else {
                sheet.getRow(38).getCell(index).setCellValue(0);
            }
            index ++;
        }

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelzdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.1-5");
        XSSFSheet sheet = wb.getSheet("表4.1.1-5");
        int index = 3;
        for (Map<String, Object> datum : data) {

            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("zdtqdjcds")){
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("zdtqdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }

            if (datum.containsKey("zdtqdhgds")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("zdtqdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(0);
            }

            if (datum.containsKey("zdtqdhgl")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("zdtqdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(0);
            }

            if (datum.containsKey("zddmccjcds")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("zddmccjcds").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(0);
            }

            if (datum.containsKey("zddmcchgds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("zddmcchgds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }

            if (datum.containsKey("zddmcchgl")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("zddmcchgl").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(0);
            }

            index ++;
        }


    }

    /**
     *
     * @param wb
     */
    private void DBExcelhdData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.1-4");
        XSSFSheet sheet = wb.getSheet("表4.1.1-4");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("hdtqdjcds")){
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("hdtqdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }

            if (datum.containsKey("hdtqdhgds")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("hdtqdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(0);
            }

            if (datum.containsKey("hdtqdhgl")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("hdtqdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(0);
            }

            if (datum.containsKey("hdjgccjcds")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("hdjgccjcds").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(0);
            }

            if (datum.containsKey("hdjgcchgds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("hdjgcchgds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }
            if (datum.containsKey("hdjgcchgl")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("hdjgcchgl").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(0);
            }

            index ++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelxqData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.1-3");
        XSSFSheet sheet = wb.getSheet("表4.1.1-3");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("xqtqdjcds")){
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("xqtqdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }
            if (datum.containsKey("xqtqdhgds")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("xqtqdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(0);
            }

            if (datum.containsKey("xqtqdhgl")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("xqtqdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(0);
            }

            if (datum.containsKey("xqjgccjcds")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("xqjgccjcds").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(0);
            }

            if (datum.containsKey("xqjgcchgds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("xqjgcchgds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }
            if (datum.containsKey("xqjgcchgl")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("xqjgcchgl").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(0);
            }
            index ++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelpsData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.1-2");
        XSSFSheet sheet = wb.getSheet("表4.1.1-2");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("psdmccjcds")){
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("psdmccjcds").toString()));
            }else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }
            if (datum.containsKey("psdmcchgds")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("psdmcchgds").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(0);
            }
            if (datum.containsKey("psdmcchgl")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("psdmcchgl").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(0);
            }

            if (datum.containsKey("pspqhdjcds")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("pspqhdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(4).setCellValue(0);
            }

            if (datum.containsKey("pspqhdhgds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("pspqhdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }

            if (datum.containsKey("pspqhdhgl")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("pspqhdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(0);
            }

            index ++;
        }


    }

    /**
     *
     * @param wb
     */
    private void DBExceltsfData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表4.1.1-1");
        XSSFSheet sheet = wb.getSheet("表4.1.1-1");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("ysdjcds")){
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("ysdjcds").toString()));
            }else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }

            if (datum.containsKey("ysdhgds")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("ysdhgds").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(0);
            }

            if (datum.containsKey("ysdhgl")){
                sheet.getRow(index).getCell(3).setCellValue(Double.valueOf(datum.get("ysdhgl").toString()));
            }else {
                sheet.getRow(index).getCell(3).setCellValue(0);
            }

            if (datum.containsKey("cjjcds")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("cjjcds").toString()));

            }else {
                sheet.getRow(index).getCell(4).setCellValue(0);
            }
            if (datum.containsKey("cjhgds")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("cjhgds").toString()));

            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }

            if (datum.containsKey("cjhgl")){
                sheet.getRow(index).getCell(6).setCellValue(Double.valueOf(datum.get("cjhgl").toString()));
            }else {
                sheet.getRow(index).getCell(6).setCellValue(0);
            }

            if (datum.containsKey("wcjcds")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("wcjcds").toString()));
            }else {
                sheet.getRow(index).getCell(7).setCellValue(0);
            }

            if (datum.containsKey("wchgds")){
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("wchgds").toString()));
            }else {
                sheet.getRow(index).getCell(8).setCellValue(0);
            }

            if (datum.containsKey("wchgl")){
                sheet.getRow(index).getCell(9).setCellValue(Double.valueOf(datum.get("wchgl").toString()));
            }else {
                sheet.getRow(index).getCell(9).setCellValue(0);

            }

            if (datum.containsKey("bpjcds")){
                sheet.getRow(index).getCell(10).setCellValue(Double.valueOf(datum.get("bpjcds").toString()));
            }else {
                sheet.getRow(index).getCell(10).setCellValue(0);
            }
            if (datum.containsKey("bphgds")){
                sheet.getRow(index).getCell(11).setCellValue(Double.valueOf(datum.get("bphgds").toString()));

            }else {
                sheet.getRow(index).getCell(11).setCellValue(0);

            }
            if (datum.containsKey("bphgl")){
                sheet.getRow(index).getCell(12).setCellValue(Double.valueOf(datum.get("bphgl").toString()));
            }else {
                sheet.getRow(index).getCell(12).setCellValue(0);
            }
            index ++;
        }

    }

    /**
     *
     * @param wb
     * @param data
     */
    private void DBExcelSdcjData(XSSFWorkbook wb, List<Map<String, Object>> data) {
        createRow(wb,data.size(),"表3.4.4-1");
        XSSFSheet sheet = wb.getSheet("表3.4.4-1");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());

            double tcccs = 0;
            double tccjs = 0;
            if (datum.containsKey("tcccs")){
                sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("tcccs").toString()));
                tcccs = Double.valueOf(datum.get("tcccs").toString());

            }else {
                sheet.getRow(index).getCell(1).setCellValue(0);
            }
            if (datum.containsKey("tccjs")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("tccjs").toString()));
                tccjs = Double.valueOf(datum.get("tccjs").toString());

            }else {
                sheet.getRow(index).getCell(2).setCellValue(0);
            }
            sheet.getRow(index).getCell(3).setCellValue(tccjs != 0 ? tcccs*100/tccjs : 0);

            double cccs = 0;
            double ccjs = 0;
            if (datum.containsKey("cccs")){
                sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("cccs").toString()));
                cccs = Double.valueOf(datum.get("cccs").toString());
            }else {
                sheet.getRow(index).getCell(4).setCellValue(0);
            }
            if (datum.containsKey("ccjs")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("ccjs").toString()));
                ccjs = Double.valueOf(datum.get("ccjs").toString());
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }
            sheet.getRow(index).getCell(6).setCellValue(ccjs != 0 ? cccs*100/ccjs : 0);

            double zccs = 0;
            double zcjs = 0;
            if (datum.containsKey("zccs")){
                sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("zccs").toString()));
                zccs = Double.valueOf(datum.get("zccs").toString());
            }else {
                sheet.getRow(index).getCell(7).setCellValue(0);
            }
            if (datum.containsKey("zcjs")){
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("zcjs").toString()));
                zcjs = Double.valueOf(datum.get("zcjs").toString());
            }else {
                sheet.getRow(index).getCell(8).setCellValue(0);
            }
            sheet.getRow(index).getCell(9).setCellValue(zcjs != 0 ? zccs*100/zcjs : 0);

            double dccs = 0;
            double dcjs = 0;
            if (datum.containsKey("dccs")){
                sheet.getRow(index).getCell(10).setCellValue(Double.valueOf(datum.get("dccs").toString()));
                dccs = Double.valueOf(datum.get("dccs").toString());
            }else {
                sheet.getRow(index).getCell(10).setCellValue(0);
            }
            if (datum.containsKey("dcjs")){
                sheet.getRow(index).getCell(11).setCellValue(Double.valueOf(datum.get("dcjs").toString()));
                dcjs = Double.valueOf(datum.get("dcjs").toString());
            }else {
                sheet.getRow(index).getCell(11).setCellValue(0);
            }
            sheet.getRow(index).getCell(12).setCellValue(dcjs != 0 ? dccs*100/dcjs : 0);

            index ++;
        }

    }

    /**
     *
     * @param wb
     * @param data
     * @param proname
     */
    private void DBExcelJaData(XSSFWorkbook wb, List<Map<String, Object>> data, String proname) {
        createRow(wb,data.size(),"表3.4.3-1");
        XSSFSheet sheet = wb.getSheet("表3.4.3-1");
        int index = 3;
        for (Map<String, Object> datum : data) {
            sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
            if (datum.containsKey("bzccs")){
                sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("bzccs").toString()));
            }else {
                sheet.getRow(index).getCell(2).setCellValue(0);
            }

            if (datum.containsKey("bxccs")){
                sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("bxccs").toString()));
            }else {
                sheet.getRow(index).getCell(5).setCellValue(0);
            }

            if (datum.containsKey("fhlccs")){
                sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("fhlccs").toString()));
            }else {
                sheet.getRow(index).getCell(8).setCellValue(0);
            }

            index ++;
        }
    }

    /**
     *
     * @param wb
     * @param data
     * @param proname
     */
    private void DBExcelQlData(XSSFWorkbook wb, List<Map<String, Object>> data, String proname) throws IOException {
        if (data.size()>0){
            createRow(wb,data.size(),"表3.4.2-1");

            XSSFSheet sheet = wb.getSheet("表3.4.2-1");

            int index = 3;
            for (Map<String, Object> datum : data) {
                sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());

                double tdccs = 0;
                double tdcjs = 0;
                if (datum.containsKey("tdcjs")){
                    sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(datum.get("tdcjs").toString()));
                    tdccs = Double.valueOf(datum.get("tdccs").toString());
                }else {
                    sheet.getRow(index).getCell(1).setCellValue(Double.valueOf(0));
                }
                if (datum.containsKey("tdccs")){
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("tdccs").toString()));
                    tdcjs = Double.valueOf(datum.get("tdcjs").toString());
                }else {
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
                }
                sheet.getRow(index).getCell(3).setCellValue(tdcjs != 0 ? tdccs*100/tdcjs : 0);


                double dccs = 0;
                double dcjs = 0;
                if (datum.containsKey("dcjs")){
                    sheet.getRow(index).getCell(4).setCellValue(Double.valueOf(datum.get("dcjs").toString()));
                    dcjs = Double.valueOf(datum.get("dcjs").toString());
                }else {
                    sheet.getRow(index).getCell(4).setCellValue(0);
                }
                if (datum.containsKey("dccs")){
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("dccs").toString()));
                    dccs = Double.valueOf(datum.get("dccs").toString());
                }else {
                    sheet.getRow(index).getCell(5).setCellValue(0);
                }
                sheet.getRow(index).getCell(6).setCellValue(dcjs != 0 ? dccs*100/dcjs : 0);

                double zccs = 0;
                double zcjs = 0;
                if (datum.containsKey("zcjs")){
                    sheet.getRow(index).getCell(7).setCellValue(Double.valueOf(datum.get("zcjs").toString()));
                    zcjs = Double.valueOf(datum.get("zcjs").toString());
                }else {
                    sheet.getRow(index).getCell(7).setCellValue(0);
                }
                if (datum.containsKey("zccs")){
                    sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("zccs").toString()));
                    zccs = Double.valueOf(datum.get("zccs").toString());
                }else {
                    sheet.getRow(index).getCell(8).setCellValue(0);
                }
                sheet.getRow(index).getCell(9).setCellValue(zcjs != 0 ? zccs*100/zcjs : 0);

                double xccs = 0;
                double xcjs = 0;
                if (datum.containsKey("xcjs")){
                    sheet.getRow(index).getCell(10).setCellValue(Double.valueOf(datum.get("xcjs").toString()));
                    xcjs = Double.valueOf(datum.get("xcjs").toString());
                }else {
                    sheet.getRow(index).getCell(10).setCellValue(0);
                }

                if (datum.containsKey("xccs")){
                    sheet.getRow(index).getCell(11).setCellValue(Double.valueOf(datum.get("xccs").toString()));
                    xccs = Double.valueOf(datum.get("xccs").toString());
                }else {
                    sheet.getRow(index).getCell(11).setCellValue(0);
                }
                sheet.getRow(index).getCell(12).setCellValue(xcjs != 0 ? xccs*100/xcjs : 0);
                index ++;
            }


        }

    }

    /**
     *
     *
     * @param wb
     * @param data
     * @param proname
     * @throws IOException
     */
    private void DBExcelData(XSSFWorkbook wb, List<Map<String, Object>> data, String proname) throws IOException {
        if (data.size()>0){
            createRow(wb,data.size(),"表3.4.1-1");
            XSSFSheet sheet = wb.getSheet("表3.4.1-1");
            int index = 3;
            for (Map<String, Object> datum : data) {
                sheet.getRow(index).getCell(0).setCellValue(datum.get("htd").toString());
                if (datum.containsKey("hdccs")){
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(datum.get("hdccs").toString()));
                }else {
                    sheet.getRow(index).getCell(2).setCellValue(Double.valueOf(0));
                }
                if (datum.containsKey("zdccs")){
                    sheet.getRow(index).getCell(5).setCellValue(Double.valueOf(datum.get("zdccs").toString()));
                }else {
                    sheet.getRow(index).getCell(5).setCellValue(0);
                }
                if (datum.containsKey("xqccs")){
                    sheet.getRow(index).getCell(8).setCellValue(Double.valueOf(datum.get("xqccs").toString()));
                }else {
                    sheet.getRow(index).getCell(8).setCellValue(0);
                }
                index ++;
            }
        }
    }

    /**
     *
     * @param wb
     * @param gettableNum
     * @param sheetname
     */
    private void createRow(XSSFWorkbook wb,int gettableNum,String sheetname) {
        for(int i = 1; i < gettableNum; i++){
            RowCopy.copyRows(wb, sheetname, sheetname, 3, 4, i*1);
        }
    }


    /**
     * 表5.3-1  建设项目工程质量评定汇总表
     * @param proname
     * @return
     */
    private List<Map<String, Object>> getxsxmpdData(String proname) throws IOException {
        String path = filespath+ File.separator+proname+File.separator;
        List<String> filteredFiles = filterFiles(path);

        List<Map<String,Object>> mapList = new ArrayList<>();

        for (String filteredFile : filteredFiles) {
            String pdnpath = path+filteredFile+ File.separator;
            XSSFWorkbook wb = null;
            File f = new File(pdnpath+"00评定表.xlsx");
            if (!f.exists()){
                break;
            }else {
                FileInputStream in = new FileInputStream(f);
                wb = new XSSFWorkbook(in);
                XSSFSheet sheet = wb.getSheet("合同段");
                Map map = new HashMap();
                for (int i = sheet.getFirstRowNum(); i < sheet.getPhysicalNumberOfRows(); i++) {
                    XSSFRow row = sheet.getRow(i);

                    if (row.getCell(0).getStringCellValue().equals("合同段质量等级")){
                        String djValue = row.getCell(1).getStringCellValue();
                        map.put("htdValue",sheet.getRow(1).getCell(1).getStringCellValue());
                        map.put("djValue",djValue);
                        mapList.add(map);
                    }

                }
            }
        }
        return mapList;
    }

    /**
     * 表5.2-1  合同段工程质量评定汇总表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> gethtdpdData(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        XSSFWorkbook wb = null;
        List<Map<String, Object>> resultlist = new ArrayList<>();
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
        if (f.exists()) {
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            List<Map<String, Object>> list = processhtdSheet(wb);
            resultlist.addAll(list);
        }
        return resultlist;
    }

    /**
     * 表5.1.3-1  隧道单位工程质量评定汇总表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getsdpdData(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        XSSFWorkbook wb = null;
        List<Map<String, Object>> dwgclist = new ArrayList<>();
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
        if (f.exists()){
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);

            for (Sheet sheet : wb) {
                String sheetName = sheet.getSheetName();
                // 检查工作表名是否以"分部-"开头
                if (sheetName.contains("隧道")) {
                    // 处理工作表数据
                    List<Map<String, Object>> list = processSheet(sheet);
                    dwgclist.addAll(list);
                }
            }
        }
        return dwgclist;

    }

    /**
     * 表5.1.2-1  桥梁单位工程质量评定汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getqlpdData(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        XSSFWorkbook wb = null;
        List<Map<String, Object>> dwgclist = new ArrayList<>();
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
        if (f.exists()){
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);

            for (Sheet sheet : wb) {
                String sheetName = sheet.getSheetName();
                // 检查工作表名是否以"分部-"开头
                if (sheetName.contains("桥")) {
                    // 处理工作表数据
                    List<Map<String, Object>> list = processSheet(sheet);
                    dwgclist.addAll(list);
                }
            }
        }
        return dwgclist;

    }

    /**
     * 表5.1.1-1  路基单位工程质量评定汇总表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getljhzbData(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        XSSFWorkbook wb = null;
        List<Map<String, Object>> dwgclist = new ArrayList<>();
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "00评定表.xlsx");
        if (f.exists()){
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);


            for (Sheet sheet : wb) {
                String sheetName = sheet.getSheetName();
                // 检查工作表名是否以"分部-"开头
                if (sheetName.equals("分部-路基")) {
                    // 处理工作表数据
                    List<Map<String, Object>> list = processSheet(sheet);
                    dwgclist.addAll(list);
                }
            }
        }
        return dwgclist;
    }

    /**
     * 表4.1.6-3  隧道工程衬砌厚度检测数据分析表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getsdcqfxbData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> cqhd = jjgFbgcSdgcCqhdService.lookJdbjg(commonInfoVo);//[{合格点数=96, 检测总点数=96, 检测项目=李家湾隧道左线, 合格率=100.00}, {合格点数=96, 检测总点数=96, 检测项目=李家湾隧道右线, 合格率=100.00}]
        if (cqhd.size()>0){
            for (Map<String, Object> map : cqhd) {
                map.put("htd",commonInfoVo.getHtd());
                String sdmc = map.get("检测项目").toString();
                int num = jjgFbgcSdgcCqhdService.getds(commonInfoVo,sdmc);
                map.put("ds",num);
                resultList.add(map);
            }
        }
        return resultList;

    }

    /**
     * 表4.1.6-2  桥梁下部墩台竖直度检测数据分析表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getqlszdfxbData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> szd = jjgFbgcQlgcXbSzdService.lookJdbjg(commonInfoVo);
        if (szd.size()>0){
            String[] valus = szd.get(0).get("yxpc").toString().replace("[", "").replace("]", "").split(",");
            for (int i =0; i< valus.length;i++){
                Map map = new HashMap<>();
                map.put("htd",commonInfoVo.getHtd());
                map.put("gdz",valus[i]);
                map.put("max",getmaxvalue(szd.get(0).get("scz").toString()));
                map.put("min",getminvalue(szd.get(0).get("scz").toString()));
                map.put("jcds",szd.get(0).get("总点数"));
                map.put("hgds",szd.get(0).get("合格点数"));
                map.put("hgl",szd.get(0).get("合格率"));
                resultList.add(map);
            }

        }
        return resultList;

    }

    /**
     * 表4.1.6-1  路基工程压实度检测数据分析表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getysdfxbData(CommonInfoVo commonInfoVo) throws IOException {

        List<Map<String, Object>> resultList = new ArrayList<>();

        List<Map<String, Object>> list1 = jjgFbgcLjgcLjtsfysdHtService.lookJdbjg(commonInfoVo);
        //[{合格点数=7, 压实度值=[96.5, 98.2, 97.0, 95.5, 98.3, 99.1, 97.0], 检测点数=7, 平均值=97.37, 规定值=96, 结果=合格, 压实度项目=灰土, 合格率=100.00, 标准差=1.230, 代表值=96.70}]
        if (list1.size()>0){
            Map map = new HashMap<>();
            map.put("htd",commonInfoVo.getHtd());
            map.put("bzz",list1.get(0).get("规定值"));
            map.put("jz",Double.valueOf(list1.get(0).get("规定值").toString())-5);
            map.put("max",getmaxvalue(list1.get(0).get("压实度值").toString()));
            map.put("min",getminvalue(list1.get(0).get("压实度值").toString()));
            map.put("dbz",list1.get(0).get("代表值"));
            map.put("jcds",list1.get(0).get("检测点数"));
            map.put("hgds",list1.get(0).get("合格点数"));
            map.put("hgl",list1.get(0).get("合格率"));
            resultList.add(map);
        }
        return resultList;
    }


    /**
     * 表4.1.5-5  交通安全设施检测结果汇总表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getjabzData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        Map resultmap = new HashMap<>();
        resultmap.put("htd",commonInfoVo.getHtd());
        List<Map<String, Object>> list1 = getbzData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list1)){
            for (Map<String, Object> map : list1) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list2 = getbxData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list2)){
            for (Map<String, Object> map : list2) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list3 = getfhlData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list3)){
            for (Map<String, Object> map : list3) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list4 = getqdccData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list4)){
            for (Map<String, Object> map : list4) {
                resultmap.putAll(map);
            }
        }
        resultList.add(resultmap);
        return resultList;
    }


    /**
     *表4.1.5-4  防护栏（砼防护栏）检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getqdccData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list1 = jjgFbgcJtaqssJathlqdService.lookJdbjg(commonInfoVo);
        Map map = new LinkedHashMap();
        map.put("htd",commonInfoVo.getHtd());
        if (list1.size()>0){
            map.put("qdzds",list1.get(0).get("总点数"));
            map.put("qdhgs",list1.get(0).get("合格点数"));
            map.put("qdhgl",list1.get(0).get("合格率"));

        }

        List<Map<String, Object>> list2 = jjgFbgcJtaqssJathldmccService.lookJdbjg(commonInfoVo);
        if (list2.size()>0){
            map.put("cczds",list1.get(0).get("总点数"));
            map.put("cchgs",list1.get(0).get("合格点数"));
            map.put("cchgl",list1.get(0).get("合格率"));

        }
        resultList.add(map);
        return resultList;

    }

    /**
     *表4.1.5-3  防护栏（波形梁）检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getfhlData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgFbgcJtaqssJabxfhlService.lookJdbjg(commonInfoVo);
        //[{不合格点数=2, 合格点数=2, 总点数=2, 检测项目=波形梁板基底金属厚度, 合格率=100.00, 规定值或允许偏差=≥4},
        // {不合格点数=2, 合格点数=2, 总点数=2, 检测项目=波形梁钢护栏立柱壁厚, 合格率=100.00, 规定值或允许偏差=方柱6.0},
        // {不合格点数=0, 合格点数=0, 总点数=6, 检测项目=波形梁钢护栏横梁中心高度, 合格率=0.00, 规定值或允许偏差=三波板600±20},
        // {不合格点数=2, 合格点数=2, 总点数=2, 检测项目=波形梁钢护栏立柱埋入深度, 合格率=100.00, 规定值或允许偏差=方柱≥1650.0}]
        if (list.size()>0){
            Map resultmap = new LinkedHashMap<>();
            resultmap.put("htd",commonInfoVo.getHtd());
            for (Map<String, Object> map : list) {
                String jcxm = map.get("检测项目").toString();
                if (jcxm.equals("波形梁板基底金属厚度")){
                    resultmap.put("jzds",map.get("总点数"));
                    resultmap.put("jhgds",map.get("合格点数"));
                    resultmap.put("jhgl",map.get("合格率"));

                }else if (jcxm.equals("波形梁钢护栏立柱壁厚")){
                    resultmap.put("lzds",map.get("总点数"));
                    resultmap.put("lhgds",map.get("合格点数"));
                    resultmap.put("lhgl",map.get("合格率"));

                }else if (jcxm.equals("波形梁钢护栏立柱埋入深度")){
                    resultmap.put("szds",map.get("总点数"));
                    resultmap.put("shgds",map.get("合格点数"));
                    resultmap.put("shgl",map.get("合格率"));

                }else if (jcxm.equals("波形梁钢护栏横梁中心高度")){
                    resultmap.put("gzds",map.get("总点数"));
                    resultmap.put("ghgds",map.get("合格点数"));
                    resultmap.put("ghgl",map.get("合格率"));
                }
            }
            resultList.add(resultmap);
        }
        return resultList;

    }

    /**
     * 表4.1.5-2  标线检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getbxData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> list = jjgFbgcJtaqssJabxService.lookJdbjg(commonInfoVo);
        //[{不合格点数=33, 合格点数=897, 总点数=930, 检测项目=交安标线厚度, 合格率=96.45, 规定值或允许偏差=白、黄线2+0.5,-0.1白、黄线7+10},
        // {不合格点数=11, 合格点数=919, 总点数=930, 检测项目=交安标线白线逆反射系数, 合格率=98.82, 规定值或允许偏差=震动≥150;双组分≥350;热熔≥150}]
        if (list.size()>0){
            Map resultmap = new LinkedHashMap<>();
            resultmap.put("htd",commonInfoVo.getHtd());
            double zds = 0.0;
            double hgds = 0.0;
            for (Map<String, Object> map : list) {
                String jcxm = map.get("检测项目").toString();
                if (jcxm.equals("交安标线厚度")){
                    resultmap.put("bxhdzds",map.get("总点数"));
                    resultmap.put("bxhdhgds",map.get("合格点数"));
                    resultmap.put("bxhdhgl",map.get("合格率"));
                }else {
                    zds += Double.valueOf(map.get("总点数").toString());
                    hgds += Double.valueOf(map.get("合格点数").toString());
                }
            }
            resultmap.put("xzds",decf.format(zds));
            resultmap.put("xhgds",decf.format(hgds));
            resultmap.put("xhgl",zds!=0 ? df.format(hgds/zds*100) : 0);
            resultList.add(resultmap);
        }
        return resultList;

    }

    /**
     * 表4.1.5-1  标志检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getbzData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> bz = jjgFbgcJtaqssJabzService.lookJdbjg(commonInfoVo);
        if (bz.size()>0){
            Map resultmap = new LinkedHashMap<>();
            resultmap.put("htd",commonInfoVo.getHtd());
            double zds = 0.0;
            double hgds = 0.0;
            for (Map<String, Object> map : bz) {
                String xm = map.get("项目").toString();
                if (xm.contains("立柱竖直度")){
                    resultmap.put("szdzds",map.get("总点数"));
                    resultmap.put("szdhgds",map.get("合格点数"));
                    resultmap.put("szdhgl",map.get("合格率"));

                }
                if (xm.contains("标志板净空")){
                    resultmap.put("jkzds",map.get("总点数"));
                    resultmap.put("jkhgds",map.get("合格点数"));
                    resultmap.put("jkhgl",map.get("合格率"));

                }
                if (xm.contains("标志板厚度")){
                    resultmap.put("bzbhdzds",map.get("总点数"));
                    resultmap.put("bzbhdhgds",map.get("合格点数"));
                    resultmap.put("bzbhdhgl",map.get("合格率"));

                }
                if (xm.equals("标志面反光膜逆反射系数")){
                    zds += Double.valueOf(map.get("总点数").toString());
                    hgds += Double.valueOf(map.get("合格点数").toString());
                }
            }
            resultmap.put("xszds",decf.format(zds));
            resultmap.put("xshgds",decf.format(hgds));
            resultmap.put("xshgl",zds!=0 ? df.format(hgds/zds*100) : 0);
            resultList.add(resultmap);

        }
        return resultList;

    }


    /**
     *表4.1.4-3  隧道工程检测结果汇总表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getsdgcData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        Map resultmap = new HashMap<>();
        List<Map<String, Object>> list1 = getcqData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list1)){
            for (Map<String, Object> map : list1) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list2 = getztData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list2)){
            for (Map<String, Object> map : list2) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list3 = getsdlmData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list3)){
            for (Map<String, Object> map : list3) {
                resultmap.putAll(map);
            }
        }
        resultList.add(resultmap);
        return resultList;
    }


    /**
     * 厚度待确认，自动化指标未编写
     * 表4.1.4-3隧道路面检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdlmData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        Map map = new LinkedHashMap();
        map.put("htd",commonInfoVo.getHtd());
        //沥青路面压实度
        List<Map<String, Object>> list1 = jjgFbgcSdgcSdlqlmysdService.lookJdbjg(commonInfoVo);
        if (list1.size()>0){
            double zds = 0;
            double hgds = 0;
            for (Map<String, Object> objectMap : list1) {
                zds += Double.valueOf(objectMap.get("检测点数").toString());
                hgds += Double.valueOf(objectMap.get("合格点数").toString());
            }
            map.put("ysdjcds",decf.format(zds));
            map.put("ysdhgds",decf.format(hgds));
            map.put("ysdhgl",zds!=0 ? df.format(hgds/zds*100) : 0);
        }else {
            map.put("ysdjcds",0);
            map.put("ysdhgds",0);
            map.put("ysdhgl",0);
        }

        //沥青路面车辙

        //沥青路面渗水系数
        List<Map<String, Object>> list2 = jjgFbgcSdgcLmssxsService.lookJdbjg(commonInfoVo);
        if (list2.size()>0){
            double zds = 0;
            double hgds = 0;
            for (Map<String, Object> objectMap : list2) {
                zds += Double.valueOf(objectMap.get("检测点数").toString());
                hgds += Double.valueOf(objectMap.get("合格点数").toString());
            }
            map.put("xsjcds",decf.format(zds));
            map.put("xshgds",decf.format(hgds));
            map.put("xshgl",zds!=0 ? df.format(hgds/zds*100) : 0);
        }else {
            map.put("xsjcds",0);
            map.put("xshgds",0);
            map.put("xshgl",0);
        }

        //砼路面面强度
        List<Map<String, Object>> list3 = jjgFbgcSdgcHntlmqdService.lookJdbjg(commonInfoVo);
        if (list3.size()>0){
            double zds = 0;
            double hgds = 0;
            for (Map<String, Object> objectMap : list3) {
                zds += Double.valueOf(objectMap.get("检测点数").toString());
                hgds += Double.valueOf(objectMap.get("合格点数").toString());
            }
            map.put("qdjcds",decf.format(zds));
            map.put("qdhgds",decf.format(hgds));
            map.put("qdhgl",zds!=0 ? df.format(hgds/zds*100) : 0);
        }else {
            map.put("qdjcds",0);
            map.put("qdhgds",0);
            map.put("qdhgl",0);

        }

        //砼路面相邻板高差
        List<Map<String, Object>> list4 = jjgFbgcSdgcTlmxlbgcService.lookJdbjg(commonInfoVo);
        if (list4.size()>0){
            double zds = 0;
            double hgds = 0;
            for (Map<String, Object> objectMap : list4) {
                zds += Double.valueOf(objectMap.get("检测点数").toString());
                hgds += Double.valueOf(objectMap.get("合格点数").toString());
            }
            map.put("gcjcds",decf.format(zds));
            map.put("gchgds",decf.format(hgds));
            map.put("gchgl",zds!=0 ? df.format(hgds/zds*100) : 0);
        }else {
            map.put("gcjcds",0);
            map.put("gchgds",0);
            map.put("gchgl",0);
        }

        //平整度

        //抗滑

        //厚度
        List<Map<String, Object>> list5 = jjgFbgcSdgcSdhntlmhdzxfService.lookJdbjg(commonInfoVo);
        double hzds = 0;
        double hhgds = 0;
        if (list5.size()>0){

            for (Map<String, Object> objectMap : list5) {
                hzds += Double.valueOf(objectMap.get("检测点数").toString());
                hhgds += Double.valueOf(objectMap.get("合格点数").toString());
            }
            map.put("hdjcds",decf.format(hzds));
            map.put("hdhgds",decf.format(hhgds));
            map.put("hdhgl",hzds!=0 ? df.format(hhgds/hzds*100) : 0);
        }
        List<Map<String, Object>> list6 = jjgFbgcSdgcGssdlqlmhdzxfService.lookJdbjg(commonInfoVo);
        if (list6.size()>0){
            for (Map<String, Object> objectMap : list6) {
                hzds += Double.valueOf(objectMap.get("上面层厚度检测点数").toString());
                hhgds += Double.valueOf(objectMap.get("上面层厚度合格点数").toString());
            }
        }


        //横坡
        List<Map<String, Object>> list7 = jjgFbgcSdgcSdhpService.lookJdbjg(commonInfoVo);
        if (list7.size()>0){
            double zds = 0;
            double hgds = 0;
            for (Map<String, Object> objectMap : list7) {
                zds += Double.valueOf(objectMap.get("总点数").toString());
                hgds += Double.valueOf(objectMap.get("合格点数").toString());
            }
            map.put("hpjcds",decf.format(zds));
            map.put("hphgds",decf.format(hgds));
            map.put("hphgl",zds!=0 ? df.format(hgds/zds*100) : 0);
        }else {
            map.put("hpjcds",0);
            map.put("hphgds",0);
            map.put("hphgl",0);
        }
        resultList.add(map);
        return resultList;

    }


    /**
     * 表4.1.4-2总体检测结果汇总表
     * 还有个净空
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getztData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgFbgcSdgcZtkdService.lookJdbjg(commonInfoVo);
        Map map = new LinkedHashMap();
        map.put("htd",commonInfoVo.getHtd());
        if (list.size()>0){
            map.put("kdjcds",list.get(0).get("总点数"));
            map.put("kdhgds",list.get(0).get("合格点数"));
            map.put("kdhgl",list.get(0).get("合格率"));
        }else {
            map.put("kdjcds",0);
            map.put("kdhgds",0);
            map.put("kdhgl",0);
        }
        resultList.add(map);
        return resultList;

    }

    /**
     * 表4.1.4-1 衬砌检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getcqData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        //衬砌强度
        List<Map<String, Object>> cqqd = jjgFbgcSdgcCqtqdService.lookJdbjg(commonInfoVo);

        //衬砌厚度  多个
        List<Map<String, Object>> cqhd = jjgFbgcSdgcCqhdService.lookJdbjg(commonInfoVo);

        List<Map<String, Object>> dmpzd = jjgFbgcSdgcDmpzdService.lookJdbjg(commonInfoVo);

        Map map = new LinkedHashMap();
        map.put("htd",commonInfoVo.getHtd());
        if (cqqd.size()>0){
            map.put("cqqdjcds",cqqd.get(0).get("总点数"));
            map.put("cqqdhgds",cqqd.get(0).get("合格点数"));
            map.put("cqqdhgl",cqqd.get(0).get("合格率"));
        }else {
            map.put("cqqdjcds",0);
            map.put("cqqdhgds",0);
            map.put("cqqdhgl",0);
        }


        if (cqhd.size()>0){
            double zds = 0;
            double hgds = 0;
            for (Map<String, Object> stringObjectMap : cqhd) {
                zds += Double.valueOf(stringObjectMap.get("检测总点数").toString());
                hgds += Double.valueOf(stringObjectMap.get("合格点数").toString());
            }
            map.put("cqhdjcds",decf.format(zds));
            map.put("cqhdhgds",decf.format(hgds));
            map.put("cqhdhgl",zds!=0 ? df.format(hgds/zds*100) : 0);
        }else {
            map.put("cqhdjcds",0);
            map.put("cqhdhgds",0);
            map.put("cqhdhgl",0);
        }
        if (dmpzd.size()>0){
            map.put("dmpzdjcds",dmpzd.get(0).get("总点数"));
            map.put("dmpzdhgds",dmpzd.get(0).get("合格点数"));
            map.put("dmpzdhgl",dmpzd.get(0).get("合格率"));
        }else {
            map.put("dmpzdjcds",0);
            map.put("dmpzdhgds",0);
            map.put("dmpzdhgl",0);
        }
        resultList.add(map);
        return resultList;

    }

    /**
     * 表4.1.3-3  桥梁工程检测结果汇总表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getqlgcData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        Map resultmap = new HashMap<>();
        List<Map<String, Object>> list1 = getqlxbData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list1)){
            for (Map<String, Object> map : list1) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list2 = getqlsbData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list2)){
            for (Map<String, Object> map : list2) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list3 = getqmxData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list3)){
            for (Map<String, Object> map : list3) {
                resultmap.putAll(map);
            }
        }
        resultList.add(resultmap);
        return resultList;
    }


    /**
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getqmxData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        Map resultmap = new LinkedHashMap();
        resultmap.put("htd",commonInfoVo.getHtd());
        //平整度

        //横坡
        List<Map<String, Object>> list2 = jjgFbgcQlgcQmhpService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list2)){
            double zds = 0;
            double hgds = 0;

            for (Map<String, Object> map : list2) {
                zds += Double.valueOf(map.get("总点数").toString());
                hgds += Double.valueOf(map.get("合格点数").toString());
            }
            resultmap.put("qmhpzds",decf.format(zds));
            resultmap.put("qmhphgds",decf.format(hgds));
            resultmap.put("qmhphgl",zds!=0 ? df.format(hgds/zds*100) : 0);
        }

        //抗滑
        List<Map<String, Object>> list3 = jjgFbgcQlgcQmgzsdService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list3)){
            double zds = 0;
            double hgds = 0;

            for (Map<String, Object> map : list3) {
                zds += Double.valueOf(map.get("检测点数").toString());
                hgds += Double.valueOf(map.get("合格点数").toString());
            }
            resultmap.put("qmkhgzsdzds",decf.format(zds));
            resultmap.put("qmkhgzsdhgds",decf.format(hgds));
            resultmap.put("qmkhgzsdhgl",zds!=0 ? df.format(hgds/zds*100) : 0);
        }
        resultList.add(resultmap);
        return resultList;

    }

    /**
     * 表4.1.3-2  桥梁上部检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getqlsbData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();

        //墩台混凝土强度
        List<Map<String, Object>> list1 = jjgFbgcQlgcSbTqdService.lookJdbjg(commonInfoVo);

        //主要结构尺寸
        List<Map<String, Object>> list2 = jjgFbgcQlgcSbJgccService.lookJdbjg(commonInfoVo);

        //钢筋保护层厚度
        List<Map<String, Object>> list3 = jjgFbgcQlgcSbBhchdService.lookJdbjg(commonInfoVo);

        Map map = new LinkedHashMap();
        map.put("htd",commonInfoVo.getHtd());
        if (list1.size()>0){
            map.put("qlsbqdjcds",list1.get(0).get("总点数"));
            map.put("qlsbqdhgds",list1.get(0).get("合格点数"));
            map.put("qlsbqdhgl",list1.get(0).get("合格率"));

        }else {
            map.put("qlsbqdjcds",0);
            map.put("qlsbqdhgds",0);
            map.put("qlsbqdhgl",0);
        }
        if (list2.size()>0){
            map.put("qlsbjgccjcds",list2.get(0).get("总点数"));
            map.put("qlsbjgcchgds",list2.get(0).get("合格点数"));
            map.put("qlsbjgcchgl",list2.get(0).get("合格率"));
        }else {
            map.put("qlsbjgccjcds",0);
            map.put("qlsbjgcchgds",0);
            map.put("qlsbjgcchgl",0);
        }

        if (list3.size()>0){
            map.put("qlsbbhchdjcds",list3.get(0).get("总点数"));
            map.put("qlsbbhchdhgds",list3.get(0).get("合格点数"));
            map.put("qlsbbhchdhgl",list3.get(0).get("合格率"));
        }else {
            map.put("qlsbbhchdjcds",0);
            map.put("qlsbbhchdhgds",0);
            map.put("qlsbbhchdhgl",0);
        }

        resultList.add(map);
        return resultList;
    }

    /**
     * 表4.1.3-1  桥梁下部检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getqlxbData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();

        //墩台混凝土强度
        List<Map<String, Object>> tqd = jjgFbgcQlgcXbTqdService.lookJdbjg(commonInfoVo);

        //主要结构尺寸
        List<Map<String, Object>> jgcc = jjgFbgcQlgcXbJgccService.lookJdbjg(commonInfoVo);

        //钢筋保护层厚度
        List<Map<String, Object>> bhchd = jjgFbgcQlgcXbBhchdService.lookJdbjg(commonInfoVo);

        //墩台垂直度
        List<Map<String, Object>> szd = jjgFbgcQlgcXbSzdService.lookJdbjg(commonInfoVo);
        Map map = new LinkedHashMap();
        map.put("htd",commonInfoVo.getHtd());
        if (tqd.size()>0){
            map.put("tqdjcds",tqd.get(0).get("总点数"));
            map.put("tqdhgds",tqd.get(0).get("合格点数"));
            map.put("tqdhgl",tqd.get(0).get("合格率"));
        }else {
            map.put("tqdjcds",0);
            map.put("tqdhgds",0);
            map.put("tqdhgl",0);
        }

        if (jgcc.size()>0){
            map.put("jgccjcds",jgcc.get(0).get("总点数"));
            map.put("jgcchgds",jgcc.get(0).get("合格点数"));
            map.put("jgcchgl",jgcc.get(0).get("合格率"));
        }else {
            map.put("jgccjcds",0);
            map.put("jgcchgds",0);
            map.put("jgcchgl",0);
        }

        if (bhchd.size()>0){
            map.put("bhchdjcds",bhchd.get(0).get("总点数"));
            map.put("bhchdhgds",bhchd.get(0).get("合格点数"));
            map.put("bhchdhgl",bhchd.get(0).get("合格率"));
        }else {
            map.put("bhchdjcds",0);
            map.put("bhchdhgds",0);
            map.put("bhchdhgl",0);
        }
        if (szd.size()>0){
            map.put("szdjcds",szd.get(0).get("总点数"));
            map.put("szdhgds",szd.get(0).get("合格点数"));
            map.put("szdhgl",szd.get(0).get("合格率"));
        }else {
            map.put("szdjcds",0);
            map.put("szdhgds",0);
            map.put("szdhgl",0);
        }

        resultList.add(map);

        return resultList;

    }


    /**
     * 表4.1.2-24  路面面层、桥面系、隧道路面抽查项目汇总表
     * 路面工程下所有桥梁和隧道 路面的数据汇总
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getxmccData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        Map resultmap = new HashMap<>();
        resultmap.put("htd",commonInfoVo.getHtd());
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        //压实度
        List<Map<String, Object>> list1 = jjgFbgcLmgcLqlmysdService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list1)){
            double ysdzds = 0;
            double ysdhgds = 0;
            for (Map<String, Object> map : list1) {
                ysdzds += Double.valueOf(map.get("检测点数").toString());
                ysdhgds += Double.valueOf(map.get("合格点数").toString());
            }
            resultmap.put("ysdzds",decf.format(ysdzds));
            resultmap.put("ysdhgds",decf.format(ysdhgds));
            resultmap.put("ysdhgl",ysdzds!=0 ? df.format(ysdhgds/ysdzds*100) : 0);
        }
        //路面弯沉
        List<Map<String, Object>> list2 = jjgFbgcLmgcLmwcService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list2)){
            double wczds = 0;
            double wchgds = 0;
            for (Map<String, Object> map : list2) {
                wczds += Double.valueOf(map.get("检测单元数").toString());
                wchgds += Double.valueOf(map.get("合格单元数").toString());
            }
            resultmap.put("wczds",decf.format(wczds));
            resultmap.put("wchgds",decf.format(wchgds));
            resultmap.put("wchgl",wczds!=0 ? df.format(wchgds/wczds*100) : 0);
        }
        //车辙
        List<Map<String, Object>> list3 = jjgZdhCzService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list3)){
            double czzds = 0;
            double czhgds = 0;
            for (Map<String, Object> map : list3) {
                czzds += Double.valueOf(map.get("总点数").toString());
                czhgds += Double.valueOf(map.get("合格点数").toString());
            }
            resultmap.put("czzds",decf.format(czzds));
            resultmap.put("czhgds",decf.format(czhgds));
            resultmap.put("czhgl",czzds!=0 ? df.format(czhgds/czzds*100) : 0);
        }
        List<Map<String, Object>> list4 = jjgFbgcLmgcLmssxsService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list4)){
            double ssxszds = 0;
            double ssxshgds = 0;
            for (Map<String, Object> map : list4) {
                ssxszds += Double.valueOf(map.get("检测点数").toString());
                ssxshgds += Double.valueOf(map.get("合格点数").toString());
            }
            resultmap.put("ssxszds",decf.format(ssxszds));
            resultmap.put("ssxshgds",decf.format(ssxshgds));
            resultmap.put("ssxshgl",ssxszds!=0 ? df.format(ssxshgds/ssxszds*100) : 0);
        }
        List<Map<String, Object>> list5 = jjgZdhPzdService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list5)){
            double lqpzdzds = 0;
            double lqpzdhgds = 0;
            for (Map<String, Object> map : list5) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.equals("沥青路面")){
                    lqpzdzds += Double.valueOf(map.get("总点数").toString());
                    lqpzdhgds += Double.valueOf(map.get("合格点数").toString());
                }
            }
            resultmap.put("lqpzdzds",decf.format(lqpzdzds));
            resultmap.put("lqpzdhgds",decf.format(lqpzdhgds));
            resultmap.put("lqpzdhgl",lqpzdzds!=0 ? df.format(lqpzdhgds/lqpzdzds*100) : 0);
        }


        List<Map<String, Object>> list6 = jjgZdhMcxsService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list6)){
            double mcxszds = 0;
            double mcxshgds = 0;
            for (Map<String, Object> map : list6) {
                mcxszds += Double.valueOf(map.get("总点数").toString());
                mcxshgds += Double.valueOf(map.get("合格点数").toString());
            }
            resultmap.put("mcxszds",decf.format(mcxszds));
            resultmap.put("mcxshgds",decf.format(mcxshgds));
            resultmap.put("mcxshgl",mcxszds!=0 ? df.format(mcxshgds/mcxszds*100) : 0);
        }

        List<Map<String, Object>> list7 = jjgZdhGzsdService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list7)){
            double gzsdzds = 0;
            double gzsdhgds = 0;
            for (Map<String, Object> map : list7) {
                gzsdzds += Double.valueOf(map.get("总点数").toString());
                gzsdhgds += Double.valueOf(map.get("合格点数").toString());
            }
            resultmap.put("gzsdzds",decf.format(gzsdzds));
            resultmap.put("gzsdhgds",decf.format(gzsdhgds));
            resultmap.put("gzsdhgl",gzsdzds!=0 ? df.format(gzsdhgds/gzsdzds*100) : 0);
        }

        List<Map<String, Object>> list8 = jjgFbgcLmgcGslqlmhdzxfService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list8)){
            double lqhdzds = 0;
            double lqhdhgds = 0;
            for (Map<String, Object> map : list8) {
                lqhdzds += Double.valueOf(map.get("上面层厚度检测点数").toString());
                lqhdzds += Double.valueOf(map.get("总厚度检测点数").toString());
                lqhdhgds += Double.valueOf(map.get("上面层厚度合格点数").toString());
                lqhdhgds += Double.valueOf(map.get("总厚度合格点数").toString());
            }
            resultmap.put("lqhdzds",decf.format(lqhdzds));
            resultmap.put("lqhdhgds",decf.format(lqhdhgds));
            resultmap.put("lqhdhgl",lqhdzds!=0 ? df.format(lqhdhgds/lqhdzds*100) : 0);
        }

        List<Map<String, String>> list9 = jjgFbgcLmgcLmhpService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list9)){
            double lqhpzds = 0;
            double lqhphgds = 0;
            for (Map<String, String> map : list9) {
                String lmlx = map.get("路面类型");
                if (lmlx.contains("沥青路面")){
                    lqhpzds += Double.valueOf(map.get("检测点数"));
                    lqhphgds += Double.valueOf(map.get("合格点数"));
                }
            }
            resultmap.put("lqhpzds",decf.format(lqhpzds));
            resultmap.put("lqhphgds",decf.format(lqhphgds));
            resultmap.put("lqhphgl",lqhpzds!=0 ? df.format(lqhphgds/lqhpzds*100) : 0);
        }

        List<Map<String, Object>> list10 = jjgFbgcLmgcHntlmqdService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list10)){
            double hntqdzds = 0;
            double hntqdhgds = 0;
            for (Map<String, Object> map : list10) {
                hntqdzds += Double.valueOf(map.get("总点数").toString());
                hntqdhgds += Double.valueOf(map.get("合格点数").toString());
            }
            resultmap.put("hntqdzds",decf.format(hntqdzds));
            resultmap.put("hntqdhgds",decf.format(hntqdhgds));
            resultmap.put("hntqdhgl",hntqdzds!=0 ? df.format(hntqdhgds/hntqdzds*100) : 0);
        }
        List<Map<String, Object>> list11 = jjgFbgcLmgcTlmxlbgcService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list11)){
            double tlmxlbgczds = 0;
            double tlmxlbgchgds = 0;
            for (Map<String, Object> map : list11) {
                tlmxlbgczds += Double.valueOf(map.get("总点数").toString());
                tlmxlbgchgds += Double.valueOf(map.get("合格点数").toString());
            }
            resultmap.put("tlmxlbgczds",decf.format(tlmxlbgczds));
            resultmap.put("tlmxlbgchgds",decf.format(tlmxlbgchgds));
            resultmap.put("tlmxlbgchgl",tlmxlbgczds!=0 ? df.format(tlmxlbgchgds/tlmxlbgczds*100) : 0);
        }
        //混凝土路面平整度
        if (CollectionUtils.isNotEmpty(list5)){
            double hntpzdzds = 0;
            double hntpzdhgds = 0;
            for (Map<String, Object> map : list5) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.equals("混凝土路面")){
                    hntpzdzds += Double.valueOf(map.get("总点数").toString());
                    hntpzdhgds += Double.valueOf(map.get("合格点数").toString());
                }
            }
            resultmap.put("hntpzdzds",decf.format(hntpzdzds));
            resultmap.put("hntpzdhgds",decf.format(hntpzdhgds));
            resultmap.put("hntpzdhgl",hntpzdzds!=0 ? df.format(hntpzdhgds/hntpzdzds*100) : 0);
        }

        //抗滑
        List<Map<String, Object>> list12 = jjgFbgcLmgcLmgzsdsgpsfService.lookJdbjg(commonInfoVo);//手工铺沙法
        if (CollectionUtils.isNotEmpty(list12)){
            double khzds = 0;
            double khhgds = 0;
            for (Map<String, Object> map : list12) {
                khzds += Double.valueOf(map.get("检测点数").toString());
                khhgds += Double.valueOf(map.get("合格点数").toString());
            }
            resultmap.put("khzds",decf.format(khzds));
            resultmap.put("khhgds",decf.format(khhgds));
            resultmap.put("khhgl",khzds!=0 ? df.format(khhgds/khzds*100) : 0);
        }

        List<Map<String, Object>> list13 = jjgFbgcLmgcHntlmhdzxfService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list13)){
            double hnthdzds = 0;
            double hnthgds = 0;
            for (Map<String, Object> map : list13) {
                hnthdzds += Double.valueOf(map.get("检测点数").toString());
                hnthgds += Double.valueOf(map.get("合格点数").toString());
            }
            resultmap.put("hnthdzds",decf.format(hnthdzds));
            resultmap.put("hnthgds",decf.format(hnthgds));
            resultmap.put("hnthgl",hnthdzds!=0 ? df.format(hnthgds/hnthdzds*100) : 0);
        }

        if (CollectionUtils.isNotEmpty(list9)){
            double hnthpzds = 0;
            double hnthphgds = 0;
            for (Map<String, String> map : list9) {
                String lmlx = map.get("路面类型");
                if (lmlx.contains("混凝土路面")){
                    hnthpzds += Double.valueOf(map.get("检测点数"));
                    hnthphgds += Double.valueOf(map.get("合格点数"));
                }
            }
            resultmap.put("hnthpzds",decf.format(hnthpzds));
            resultmap.put("hnthphgds",decf.format(hnthphgds));
            resultmap.put("hnthphdl",hnthpzds!=0 ? df.format(hnthphgds/hnthpzds*100) : 0);
        }

        resultList.add(resultmap);
        return resultList;
    }

    /**
     *
     * 表4.1.2-23 隧道路面工程检测结果汇总表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getsdlmhzData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        Map resultmap = new HashMap();
        List<Map<String, Object>> list1 = getsdysdData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list1)){
            for (Map<String, Object> map : list1) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list2 = getsdczData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list2)){
            for (Map<String, Object> map : list2) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list3 = getsdssxsData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list3)){
            for (Map<String, Object> map : list3) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list4 = getsdhntqdData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list4)){
            for (Map<String, Object> map : list4) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list5 = getsdhntxlbgcData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list5)){
            for (Map<String, Object> map : list5) {
                resultmap.putAll(map);
            }
        }
        List<Map<String, Object>> list6 = getsdpzdData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list6)){
            for (Map<String, Object> map : list6) {
                String sdpzdlmlx = map.get("sdpzdlmlx").toString();
                if (sdpzdlmlx.equals("沥青路面")){
                    resultmap.put("sdpzdlqjcds",map.get("sdpzdjcds"));
                    resultmap.put("sdpzdlqhgds",map.get("sdpzdhgds"));
                    resultmap.put("sdpzdlqhgl",map.get("sdpzdhgl"));
                }else if (sdpzdlmlx.equals("混凝土路面")){
                    resultmap.put("sdpzdhntjcds",map.get("sdpzdjcds"));
                    resultmap.put("sdpzdhnthgds",map.get("sdpzdhgds"));
                    resultmap.put("sdpzdhnthgl",map.get("sdpzdhgl"));
                }
            }
        }

        List<Map<String, Object>> list7 = getsdkhData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list7)){
            for (Map<String, Object> map : list7) {
                String hplmlx = map.get("khlmlx").toString();
                if (hplmlx.equals("隧道路面")){
                    String khkpzb = map.get("sdkhzb").toString();
                    if (khkpzb.equals("摩擦系数")){
                        resultmap.put("khlqmcxsjcds",map.get("sdkhjcds"));
                        resultmap.put("khlqmcxshgds",map.get("sdkhhgds"));
                        resultmap.put("khlqmcxshgl",map.get("sdkhhgl"));
                    }else if (khkpzb.equals("构造深度")){
                        resultmap.put("khlqgzsdjcds",map.get("sdkhjcds"));
                        resultmap.put("khlqgzsdhgds",map.get("sdkhhgds"));
                        resultmap.put("khlqgzsdhgl",map.get("sdkhhgl"));
                    }
                }else if (hplmlx.equals("混凝土隧道")){
                    String khkpzb = map.get("sdkhzb").toString();
                    if (khkpzb.equals("摩擦系数")){
                        resultmap.put("khhntmcxsjcds",map.get("sdkhjcds"));
                        resultmap.put("khhntmcxshgds",map.get("sdkhhgds"));
                        resultmap.put("khhntmcxshgl",map.get("sdkhhgl"));
                    }else if (khkpzb.equals("构造深度")){
                        resultmap.put("khhntgzsdjcds",map.get("sdkhjcds"));
                        resultmap.put("khhntgzsdhgds",map.get("sdkhhgds"));
                        resultmap.put("khhntgzsdhgl",map.get("sdkhhgl"));
                    }

                }
            }

        }

        List<Map<String, Object>> list10 = getsdhdData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list10)){
            for (Map<String, Object> map : list10) {
                String hplmlx = map.get("sdlmlx").toString();
                if (hplmlx.equals("沥青路面")){
                    resultmap.put("hdlqzds",map.get("sdjcds"));
                    resultmap.put("hdlqhgds",map.get("sdhgs"));
                    resultmap.put("hdlqhgl",map.get("sdhgl"));
                }else if (hplmlx.equals("混凝土路面")){
                    resultmap.put("hdhntzds",map.get("sdjcds"));
                    resultmap.put("hdhnthgds",map.get("sdhgs"));
                    resultmap.put("hdhnthgl",map.get("sdhgl"));
                }
            }
        }
        List<Map<String, Object>> list11 = getsdhpData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list11)){
            for (Map<String, Object> map : list11) {
                String hplmlx = map.get("sdhplmlx").toString();
                if (hplmlx.equals("沥青路面")){
                    resultmap.put("hplqzds",map.get("sdhpzds"));
                    resultmap.put("hplqhgds",map.get("sdhphgds"));
                    resultmap.put("hplqhgl",map.get("sdhphgl"));
                }else if (hplmlx.equals("混凝土路面")){
                    resultmap.put("hphntzds",map.get("sdhpzds"));
                    resultmap.put("hphnthgds",map.get("sdhphgds"));
                    resultmap.put("hphnthgl",map.get("sdhphgl"));
                }
            }

        }
        resultList.add(resultmap);
        return resultList;

    }



    /**
     * 表4.1.2-22 隧道路面横坡检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdhpData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, String>> list = jjgFbgcLmgcLmhpService.lookJdbjg(commonInfoVo);
        if (list.size()>0){
            double lmzds = 0.0;
            double lmhgds = 0.0;

            double lmzds1 = 0.0;
            double lmhgds1 = 0.0;
            boolean a = false;
            boolean b = false;
            for (Map<String, String> map : list) {
                String lmlx = map.get("路面类型");
                if (lmlx.contains("沥青隧道")){
                    a = true;
                    lmzds +=Double.valueOf(map.get("检测点数"));
                    lmhgds +=Double.valueOf(map.get("合格点数"));

                }
                if (lmlx.contains("混凝土隧道")){
                    b = true;
                    lmzds1 +=Double.valueOf(map.get("检测点数"));
                    lmhgds1 +=Double.valueOf(map.get("合格点数"));
                }
            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("sdhpzds",decf.format(lmzds));
                map1.put("sdhphgds",decf.format(lmhgds));
                map1.put("sdhphgl",lmzds!=0 ? df.format(lmhgds/lmzds*100) : 0);
                map1.put("sdhplmlx","沥青隧道");
                resultList.add(map1);
            }
            if (b){
                Map map2 = new HashMap();
                map2.put("htd",commonInfoVo.getHtd());
                map2.put("sdhpzds",decf.format(lmzds1));
                map2.put("sdhphgds",decf.format(lmhgds1));
                map2.put("sdhphgl",lmzds1!=0 ? df.format(lmhgds1/lmzds1*100) : 0);
                map2.put("sdhplmlx","混凝土隧道");
                resultList.add(map2);
            }
        }
        return resultList;
    }

    /**
     * 表4.1.2-21（3）  隧道路面厚度检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdhdData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        //沥青
        List<Map<String, Object>> list = jjgFbgcLmgcGslqlmhdzxfService.lookJdbjg(commonInfoVo);
        if (list.size()>0){
            double jcds = 0;
            double hgds = 0;
            boolean a = false;
            for (Map<String, Object> map : list) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.contains("隧道") ){
                    a = true;
                    jcds += Double.valueOf(map.get("上面层厚度检测点数").toString());
                    jcds += Double.valueOf(map.get("总厚度检测点数").toString());
                    hgds += Double.valueOf(map.get("上面层厚度合格点数").toString());
                    hgds += Double.valueOf(map.get("总厚度合格点数").toString());
                }
            }
            if (a){
                Map map2 = new HashMap();
                map2.put("htd",commonInfoVo.getHtd());
                map2.put("sdlmlx","沥青路面");
                map2.put("sdjcds",decf.format(jcds));
                map2.put("sdhgs",decf.format(hgds));
                map2.put("sdhgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
                resultList.add(map2);
            }
        }
        return resultList;

    }

    /**
     * 表4.1.2-21（2）  隧道路面雷达厚度检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdldhdData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> list = jjgZdhLdhdService.lookJdbjg(commonInfoVo);
        List<Map<String, Object>> resultList = new ArrayList<>();
        Double sdMax = 0.0;
        Double sdMin = Double.MAX_VALUE;
        double jcds = 0;
        double hgs = 0;
        if (list.size()>0){
            boolean a = false;
            for (Map<String, Object> map : list) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.contains("隧道")){
                    a = true;
                    double max = Double.valueOf(map.get("Max").toString());
                    sdMax = (max > sdMax) ? max : sdMax;
                    double min = Double.valueOf(map.get("Min").toString());
                    sdMin = (min < sdMin) ? min : sdMin;

                    jcds += Double.valueOf(map.get("总点数").toString());
                    hgs += Double.valueOf(map.get("合格点数").toString());

                }
            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("sdldhdlb","总厚度");
                map1.put("sdldhdlmlx","沥青路面");
                map1.put("sdldhdsjz",list.get(0).get("设计值"));
                map1.put("sdldhdjcds",decf.format(jcds));
                map1.put("sdldhdhgds",decf.format(hgs));
                map1.put("sdldhdhgl",jcds!=0 ? df.format(hgs/jcds*100) : 0);
                map1.put("sdldhdbhfw",sdMin+"~"+sdMax);
                resultList.add(map1);
            }
        }

        return resultList;

    }

    /**
     * 表4.1.2-21（1）  隧道路面钻芯法厚度检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdzxfhdData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        //沥青
        List<Map<String, Object>> list = jjgFbgcLmgcGslqlmhdzxfService.lookJdbjg(commonInfoVo);
        if (list.size()>0){
            Double sMax = 0.0;
            Double sMin = Double.MAX_VALUE;
            Double zMax = 0.0;
            Double zMin = Double.MAX_VALUE;
            String zdbz = "";
            String ydbz = "";
            String zzdbz = "";
            String yzdbz = "";
            String smcsjz = "";
            String zhdsjz = "";
            double smcjcds = 0;
            double smchgs = 0;
            double zhdjcds = 0;
            double zhdhgs = 0;
            boolean a = false;
            for (Map<String, Object> map : list) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.equals("隧道左幅") || lmlx.equals("隧道右幅")){
                    a = true;
                    double max = Double.valueOf(map.get("上面层平均值最大值").toString());
                    sMax = (max > sMax) ? max : sMax;

                    double min = Double.valueOf(map.get("上面层平均值最小值").toString());
                    sMin = (min < sMin) ? min : sMin;

                    double max1 = Double.valueOf(map.get("总厚度平均值最大值").toString());
                    zMax = (max1 > zMax) ? max1 : zMax;

                    double min1 = Double.valueOf(map.get("总厚度平均值最小值").toString());
                    zMin = (min1 < zMin) ? min1 : zMin;

                    if (lmlx.equals("隧道左幅")){
                        zdbz = map.get("上面层代表值").toString();
                        zzdbz = map.get("总厚度代表值").toString();
                    }
                    if (lmlx.equals("隧道右幅")){
                        ydbz = map.get("上面层代表值").toString();
                        yzdbz = map.get("总厚度代表值").toString();
                    }
                    smcsjz = map.get("上面层设计值").toString();
                    zhdsjz = map.get("总厚度设计值").toString();
                    smcjcds += Double.valueOf(map.get("上面层厚度检测点数").toString());
                    smchgs += Double.valueOf(map.get("上面层厚度合格点数").toString());
                    zhdjcds += Double.valueOf(map.get("总厚度检测点数").toString());
                    zhdhgs += Double.valueOf(map.get("总厚度合格点数").toString());

                }
            }
            if (a){
                Map map = new HashMap();
                map.put("htd",commonInfoVo.getHtd());
                map.put("sdzxfhdlmlx","沥青路面");
                map.put("sdzxfhdlb","上面层厚度");
                map.put("sdzxfhdpjzfw",sMin+"~"+sMax);
                map.put("sdzxfhdzdbz",zdbz);
                map.put("sdzxfhdydbz",ydbz);
                map.put("sdzxfhdsjz",smcsjz);
                map.put("sdzxfhdjcds",decf.format(smcjcds));
                map.put("sdzxfhdhgs",decf.format(smchgs));
                map.put("sdzxfhdhgl",smcjcds!=0 ? df.format(smchgs/smcjcds*100) : 0);
                resultList.add(map);
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("sdzxfhdlmlx","沥青路面");
                map1.put("sdzxfhdlb","总厚度");
                map1.put("sdzxfhdpjzfw",zMin+"~"+zMax);
                map1.put("sdzxfhdzdbz",zzdbz);
                map1.put("sdzxfhdydbz",yzdbz);
                map1.put("sdzxfhdsjz",zhdsjz);
                map1.put("sdzxfhdjcds",decf.format(zhdjcds));
                map1.put("sdzxfhdhgs",decf.format(zhdhgs));
                map1.put("sdzxfhdhgl",zhdjcds!=0 ? df.format(zhdhgs/zhdjcds*100) : 0);
                resultList.add(map1);
            }



        }
        //混凝土
        List<Map<String, Object>> list1 = jjgFbgcSdgcSdhntlmhdzxfService.lookJdbjg(commonInfoVo);
        if (list1.size()>0){
            double jcds = 0;
            double hgds = 0;
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            boolean a = false;
            for (Map<String, Object> map : list1) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.contains("隧道")){
                    a = true;
                    jcds += Double.valueOf(map.get("检测点数").toString());
                    hgds += Double.valueOf(map.get("合格点数").toString());

                    double max = Double.valueOf(map.get("最大值").toString());
                    lmMax = (max > lmMax) ? max : lmMax;
                    double min = Double.valueOf(map.get("最小值").toString());
                    lmMin = (min < lmMin) ? min : lmMin;
                }
            }
            if (a){
                Map map2 = new HashMap();
                map2.put("htd",commonInfoVo.getHtd());
                map2.put("sdzxfhdlmlx","混凝土路面");
                map2.put("sdzxfhdlb","总厚度");
                map2.put("sdzxfhdpjzfw",lmMin+"~"+lmMax);
                map2.put("sdzxfhdzdbz",list1.get(0).get("代表值"));
                map2.put("sdzxfhdydbz",list1.get(0).get("代表值"));
                map2.put("sdzxfhdsjz",list1.get(0).get("设计值"));
                map2.put("sdzxfhdjcds",decf.format(jcds));
                map2.put("sdzxfhdhgs",decf.format(hgds));
                map2.put("sdzxfhdhgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
                resultList.add(map2);
            }
        }

        return resultList;


    }

    /**
     * 表4.1.2-20  隧道路面抗滑检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdkhData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgZdhGzsdService.lookJdbjg(commonInfoVo);//自动化
        String sjz = "";
        double jcds = 0.0;
        double hgds = 0.0;

        if (list.size()>0){
            boolean a = false;
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            for (Map<String, Object> map : list) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.contains("隧道")){
                    a = true;
                    sjz = map.get("设计值").toString();
                    jcds += Double.valueOf(map.get("总点数").toString());
                    hgds += Double.valueOf(map.get("合格点数").toString());
                    double max = Double.valueOf(map.get("最大值").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("最小值").toString());
                    lmMin = (min < lmMin) ? min : lmMin;
                }
            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("sdkhlmlx","隧道路面");
                map1.put("sdkhzb","构造深度");
                map1.put("sdkhsjz",sjz);
                map1.put("sdkhbhfw",lmMin+"~"+lmMax);
                map1.put("sdkhjcds",decf.format(jcds));
                map1.put("sdkhhgds",decf.format(hgds));
                map1.put("sdkhhgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
                resultList.add(map1);
            }

        }

        List<Map<String, Object>> list1 = jjgFbgcLmgcLmgzsdsgpsfService.lookJdbjg(commonInfoVo);//手工铺沙法
        if (list1.size()>0){
            String gdz = "";
            double szds = 0;
            double shgs = 0;
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            boolean a = false;
            for (Map<String, Object> map : list1) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.equals("混凝土隧道")){
                    a = true;
                    szds += Double.valueOf(map.get("检测点数").toString());
                    shgs += Double.valueOf(map.get("合格点数").toString());
                    double max = Double.valueOf(map.get("Max").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("Min").toString());
                    lmMin = (min < lmMin) ? min : lmMin;

                    gdz = map.get("规定值").toString();
                }else {
                    continue;
                }
            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("sdkhlmlx","混凝土隧道");
                map1.put("sdkhkpzb","构造深度");
                map1.put("sdkhsjz",gdz);
                map1.put("sdkhbhfw",lmMin+"~"+lmMax);
                map1.put("sdkhjcds",decf.format(szds));
                map1.put("sdkhhgds",decf.format(shgs));
                map1.put("sdkhhgl",szds!=0 ? df.format(shgs/szds*100) : 0);
                resultList.add(map1);
            }

        }
        List<Map<String, Object>> list2 = jjgZdhMcxsService.lookJdbjg(commonInfoVo);
        if (list2.size()>0){
            boolean a = false;
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            double zds = 0;
            double hgs = 0;
            String sjz2 = "";
            for (Map<String, Object> map : list2) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.contains("隧道")){
                    a = true;
                    zds += Double.valueOf(map.get("总点数").toString());
                    hgs += Double.valueOf(map.get("合格点数").toString());
                    sjz2 = map.get("设计值").toString();

                    double max = Double.valueOf(map.get("Max").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("Min").toString());
                    lmMin = (min < lmMin) ? min : lmMin;
                }

            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("sdkhlmlx","隧道路面");
                map1.put("sdkhzb","摩擦系数");
                map1.put("sdkhsjz",sjz2);
                map1.put("sdkhbhfw",lmMin+"~"+lmMax);
                map1.put("sdkhjcds",decf.format(zds));
                map1.put("sdkhhgds",decf.format(hgs));
                map1.put("sdkhhgl",zds!=0 ? df.format(hgs/zds*100) : 0);
                resultList.add(map1);
            }

        }
        return resultList;


    }

    /**
     * 表4.1.2-19  隧道路面平整度检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdpzdData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgZdhPzdService.lookJdbjg(commonInfoVo);
        if (list.size()>0){
            for (Map<String, Object> map : list) {

                String lmlx = map.get("路面类型").toString();
                if (lmlx.equals("沥青隧道")){
                    Map map1 = new HashMap();
                    map1.put("htd",commonInfoVo.getHtd());
                    map1.put("sdpzdzb","IRI");
                    map1.put("sdpzdlmlx","沥青路面");
                    map1.put("sdpzdgdz",map.get("设计值"));
                    map1.put("sdpzdjcds",map.get("总点数"));
                    map1.put("sdpzdhgds",map.get("合格点数"));
                    map1.put("sdpzdhgl",map.get("合格率"));
                    map1.put("sdpzdbhfw",map.get("Min")+"~"+map.get("Max"));
                    resultList.add(map1);
                }
                if (lmlx.equals("混凝土隧道")){
                    Map map1 = new HashMap();
                    map1.put("htd",commonInfoVo.getHtd());
                    map1.put("sdpzdzb","IRI");
                    map1.put("sdpzdlmlx","混凝土路面");
                    map1.put("sdpzdgdz",map.get("设计值"));
                    map1.put("sdpzdjcds",map.get("总点数"));
                    map1.put("sdpzdhgds",map.get("合格点数"));
                    map1.put("sdpzdhgl",map.get("合格率"));
                    map1.put("sdpzdbhfw",map.get("Min")+"~"+map.get("Max"));
                    resultList.add(map1);
                }
            }
        }
        return resultList;

    }

    /**
     * 和getsdhntqdData一样
     * 有待确认的事项
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdhntxlbgcData(CommonInfoVo commonInfoVo) throws IOException {
        /*List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> list = jjgFbgcSdgcTlmxlbgcService.lookJdbjg(commonInfoVo);
        double jcds = 0;
        double hgds = 0;
        Double lmMax = 0.0;
        Double lmMin = Double.MAX_VALUE;
        boolean a = false;
        if (list.size()>0){
            a = true;
            for (Map<String, Object> map : list) {
                jcds += Double.valueOf(map.get("检测点数").toString());
                hgds += Double.valueOf(map.get("合格点数").toString());

                double max = Double.valueOf(map.get("最大值").toString());
                lmMax = (max > lmMax) ? max : lmMax;
                double min = Double.valueOf(map.get("最小值").toString());
                lmMin = (min < lmMin) ? min : lmMin;

            }

        }
        if (a){
            Map map = new HashMap();
            map.put("jcds",decf.format(jcds));
            map.put("hgds",decf.format(hgds));
            map.put("hgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
            map.put("max",lmMax);
            map.put("min",lmMin);
            map.put("gdz",list.get(0).get("规定值"));
            map.put("pjz",list.get(0).get("平均值"));
            resultList.add(map);
        }
        return resultList;*/
        return null;
    }

    /**
     * 有待确认的事项
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdhntqdData(CommonInfoVo commonInfoVo) throws IOException {
        /*List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> list = jjgFbgcSdgcHntlmqdService.lookJdbjg(commonInfoVo);

        *//**
         * 这块会有多个隧道的数据，平均值如何取？
         * 最大值和最小值是取全部比较后的结构
         * 还有规定值
         *//*
        double jcds = 0;
        double hgds = 0;
        Double lmMax = 0.0;
        Double lmMin = Double.MAX_VALUE;
        if (list.size()>0){
            for (Map<String, Object> map : list) {
                jcds += Double.valueOf(map.get("检测点数").toString());
                hgds += Double.valueOf(map.get("合格点数").toString());

                double max = Double.valueOf(map.get("最大值").toString());
                lmMax = (max > lmMax) ? max : lmMax;
                double min = Double.valueOf(map.get("最小值").toString());
                lmMin = (min < lmMin) ? min : lmMin;

            }
            Map map = new HashMap();
            map.put("jcds",decf.format(jcds));
            map.put("hgds",decf.format(hgds));
            map.put("hgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
            map.put("max",lmMax);
            map.put("min",lmMin);
            map.put("gdz",list.get(0).get("规定值"));
            map.put("pjz",list.get(0).get("平均值"));
            resultList.add(map);
        }

        return resultList;*/
        return null;
    }

    /**
     * 表4.1.2-16  隧道路面渗水系数检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdssxsData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.00");
        List<Map<String, Object>> list = jjgFbgcLmgcLmssxsService.lookJdbjg(commonInfoVo);
        double jcds =0.0;
        double hgds =0.0;
        Double lmMax = 0.0;
        Double lmMin = Double.MAX_VALUE;
        boolean a = false;
        if (list.size()>0){
            for (Map<String, Object> map : list) {
                String lmlx = map.get("检测项目").toString();
                if (lmlx.contains("隧道路面")){
                    a = true;
                    jcds += Double.valueOf(map.get("检测点数").toString());
                    hgds += Double.valueOf(map.get("合格点数").toString());

                    double max = Double.valueOf(map.get("最大值").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("最小值").toString());
                    lmMin = (min < lmMin) ? min : lmMin;
                }

            }
            if (a){
                Map<String,Object> map = new HashMap<>();
                map.put("htd",commonInfoVo.getHtd());
                map.put("sdssxsjcds",jcds);
                map.put("sdssxsmax",lmMax);
                map.put("sdssxsmin",lmMin);
                map.put("sdssxshgds",hgds);
                map.put("sdssxshgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
                map.put("sdssxsgdz",list.get(0).get("规定值"));
                resultList.add(map);
            }
        }
        return resultList;
    }

    /**
     * 表4.1.2-19  隧道路面车辙检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdczData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgZdhCzService.lookJdbjg(commonInfoVo);
        if (list.size()>0){
            for (Map<String, Object> map : list) {
                Map map1 = new HashMap();
                String lmlx = map.get("路面类型").toString();
                if (lmlx.contains("隧道")){
                    map1.put("htd",commonInfoVo.getHtd());
                    map1.put("sdczzb","MTD");
                    map1.put("sdczlmlx","隧道路面");
                    map1.put("sdczgdz",map.get("设计值"));
                    map1.put("sdczjcds",map.get("总点数"));
                    map1.put("sdczhgds",map.get("合格点数"));
                    map1.put("sdczhgl",map.get("合格率"));
                    map1.put("sdczbhfw",map.get("Min")+"~"+map.get("Max"));
                    resultList.add(map1);
                }
            }
        }
        return resultList;

    }

    /**
     * 表4.1.2-15  隧道路面面层压实度检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getsdysdData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> list = jjgFbgcLmgcLqlmysdService.lookJdbjg(commonInfoVo);
        Double lmMax = 0.0;
        Double lmMin = Double.MAX_VALUE;
        String gdz = "";
        String zdbz = "";
        String ydbz = "";
        Double jcds = 0.0;
        Double hgds = 0.0;
        boolean a= false;
        //还有连接线隧道
        if (list.size()>0){
            for (Map<String, Object> map : list) {
                String lx = map.get("路面类型").toString();
                if (lx.contains("隧道")){
                    a = true;
                    double max = Double.valueOf(map.get("最大值").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("最小值").toString());
                    lmMin = (min < lmMin) ? min : lmMin;

                    gdz = map.get("规定值").toString();
                    jcds += Double.valueOf(map.get("检测点数").toString());
                    hgds += Double.valueOf(map.get("合格点数").toString());
                    if (lx.equals("隧道左幅")){
                        zdbz = map.get("代表值").toString();

                    }else if (lx.equals("隧道右幅")){
                        ydbz = map.get("代表值").toString();
                    }

                }

            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("ysdlx","主线");
                map1.put("ysdsczbh",lmMin+"~"+lmMax);
                map1.put("ysdzdbz",zdbz);
                map1.put("ysdydbz",ydbz);
                map1.put("ysdgdz",gdz);
                map1.put("ysdzs",decf.format(jcds));
                map1.put("ysdhgs",decf.format(hgds));
                map1.put("ysdhgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
                resultList.add(map1);
            }
        }
        return resultList;

    }

    /**
     * 表4.1.2-14  桥面系检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getlmqlhzData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        Map resultmap = new HashMap();
        resultmap.put("htd",commonInfoVo.getHtd());
        List<Map<String, Object>> list1 = getqmpzdData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list1)){
            double qmpzdjcds = 0;
            double qmpzdhgds = 0;
            for (Map<String, Object> map : list1) {
                qmpzdjcds += Double.valueOf(map.get("qmpzdjcds").toString());
                qmpzdhgds += Double.valueOf(map.get("qmpzdhgds").toString());
            }
            resultmap.put("qmpzdjcds",decf.format(qmpzdjcds));
            resultmap.put("qmpzdhgds",decf.format(qmpzdhgds));
            resultmap.put("qmpzdhgl",qmpzdjcds!=0 ? df.format(qmpzdhgds/qmpzdjcds*100) : 0);
        }
        List<Map<String, Object>> list2 = getqmhpData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list2)){
            double zds = 0;
            double hgds = 0;
            for (Map<String, Object> map : list2) {
                zds += Double.valueOf(map.get("qmhpzds").toString());
                hgds += Double.valueOf(map.get("qmhphgds").toString());
            }
            resultmap.put("qmhpzds",decf.format(zds));
            resultmap.put("qmhphgds", decf.format(hgds));
            resultmap.put("qmhphgl",zds!=0 ? df.format(hgds/zds*100) : 0);
        }
        List<Map<String, Object>> list3 = getqmkhData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list3)){
            double mcxszds = 0;
            double mcxshgds = 0;
            double gzsdzds = 0;
            double gzsdhgds = 0;
            for (Map<String, Object> map : list3) {
                String qmkhzb = map.get("qmkhzb").toString();
                if (qmkhzb.equals("摩擦系数")){
                    mcxszds += Double.valueOf(map.get("qmkhjcds").toString());
                    mcxshgds += Double.valueOf(map.get("qmkhhgds").toString());
                }else if (qmkhzb.equals("构造深度")){
                    gzsdzds += Double.valueOf(map.get("qmkhjcds").toString());
                    gzsdhgds += Double.valueOf(map.get("qmkhhgds").toString());
                }
            }
            resultmap.put("mcxszds",decf.format(mcxszds));
            resultmap.put("mcxshgds", decf.format(mcxshgds));
            resultmap.put("mcxshgl",mcxszds!=0 ? df.format(mcxshgds/mcxszds*100) : 0);

            resultmap.put("gzsdzds",decf.format(gzsdzds));
            resultmap.put("gzsdhgds", decf.format(gzsdhgds));
            resultmap.put("gzsdhgl",gzsdzds!=0 ? df.format(gzsdhgds/gzsdzds*100) : 0);
        }
        resultList.add(resultmap);
        return resultList;
    }

    /**
     * 表4.1.2-13  桥面抗滑检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getqmkhData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgZdhGzsdService.lookJdbjg(commonInfoVo);//自动化
        String sjz = "";
        double jcds = 0.0;
        double hgds = 0.0;

        if (list.size()>0){
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            boolean a = false;
            for (Map<String, Object> map : list) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.contains("桥")){
                    a = true;
                    sjz = map.get("设计值").toString();
                    jcds += Double.valueOf(map.get("总点数").toString());
                    hgds += Double.valueOf(map.get("合格点数").toString());
                    double max = Double.valueOf(map.get("最大值").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("最小值").toString());
                    lmMin = (min < lmMin) ? min : lmMin;
                }
            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("qmkhlmlx","沥青桥");
                map1.put("qmkhzb","构造深度");
                map1.put("qmkhsjz",sjz);
                map1.put("qmkhbhfw",lmMin+"~"+lmMax);
                map1.put("qmkhjcds",decf.format(jcds));
                map1.put("qmkhhgds",decf.format(hgds));
                map1.put("qmkhhgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
                resultList.add(map1);
            }

        }

        List<Map<String, Object>> list1 = jjgFbgcLmgcLmgzsdsgpsfService.lookJdbjg(commonInfoVo);//手工铺沙法
        if (list1.size()>0){
            String gdz = "";
            double szds = 0;
            double shgs = 0;
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            boolean a = false;
            for (Map<String, Object> map : list) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.contains("桥")){
                    a = true;
                    szds += Double.valueOf(map.get("检测点数").toString());
                    shgs += Double.valueOf(map.get("合格点数").toString());
                    double max = Double.valueOf(map.get("Max").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("Min").toString());
                    lmMin = (min < lmMin) ? min : lmMin;

                    gdz = map.get("规定值").toString();
                }else {
                    continue;
                }
            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("qmkhlmlx","混凝土桥");
                map1.put("qmkhzb","构造深度");
                map1.put("qmkhsjz",gdz);
                map1.put("qmkhbhfw",lmMin+"~"+lmMax);
                map1.put("qmkhjcds",decf.format(szds));
                map1.put("qmkhhgds",decf.format(shgs));
                map1.put("qmkhhgl",szds!=0 ? df.format(shgs/szds*100) : 0);
                resultList.add(map1);
            }


        }
        List<Map<String, Object>> list2 = jjgZdhMcxsService.lookJdbjg(commonInfoVo);
        if (list2.size()>0){
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            double zds = 0;
            double hgs = 0;
            boolean a = false;
            String sjz2 = "";
            for (Map<String, Object> map : list2) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.contains("桥")){
                    a = true;
                    zds += Double.valueOf(map.get("总点数").toString());
                    hgs += Double.valueOf(map.get("合格点数").toString());
                    sjz2 = map.get("设计值").toString();

                    double max = Double.valueOf(map.get("Max").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("Min").toString());
                    lmMin = (min < lmMin) ? min : lmMin;
                }

            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("qmkhlmlx","沥青桥");
                map1.put("qmkhzb","摩擦系数");
                map1.put("qmkhsjz",sjz2);
                map1.put("qmkhbhfw",lmMin+"~"+lmMax);
                map1.put("qmkhjcds",decf.format(zds));
                map1.put("qmkhhgds",decf.format(hgs));
                map1.put("qmkhhgl",zds!=0 ? df.format(hgs/zds*100) : 0);
                resultList.add(map1);
            }

        }
        return resultList;
    }

    /**
     *表4.1.2-12  桥面系横坡检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getqmhpData(CommonInfoVo commonInfoVo) throws IOException {

        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, String>> list = jjgFbgcLmgcLmhpService.lookJdbjg(commonInfoVo);

        if (list.size()>0){
            double lmzds = 0.0;
            double lmhgds = 0.0;

            double lmzds1 = 0.0;
            double lmhgds1 = 0.0;
            boolean a = false;
            boolean b = false;
            for (Map<String, String> map : list) {
                String lmlx = map.get("路面类型");
                if (lmlx.contains("沥青桥面")){
                    a = true;
                    lmzds +=Double.valueOf(map.get("检测点数"));
                    lmhgds +=Double.valueOf(map.get("合格点数"));

                }
                if (lmlx.contains("混凝土桥面")){
                    b = true;
                    lmzds1 +=Double.valueOf(map.get("检测点数"));
                    lmhgds1 +=Double.valueOf(map.get("合格点数"));
                }
            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("qmhpzds",decf.format(lmzds));
                map1.put("qmhphgds",decf.format(lmhgds));
                map1.put("qmhphgl",lmzds!=0 ? df.format(lmhgds/lmzds*100) : 0);
                map1.put("qmhplmlx","沥青桥面");
                resultList.add(map1);
            }
            if (b){
                Map map2 = new HashMap();
                map2.put("qmhphtd",commonInfoVo.getHtd());
                map2.put("qmhpzds",decf.format(lmzds1));
                map2.put("qmhphgds",decf.format(lmhgds1));
                map2.put("qmhphgl",lmzds1!=0 ? df.format(lmhgds1/lmzds1*100) : 0);
                map2.put("qmhplmlx","混凝土桥面");
                resultList.add(map2);
            }
        }
        return resultList;
    }

    /**
     * 表4.1.2-11  桥面铺装平整度检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getqmpzdData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgZdhPzdService.lookJdbjg(commonInfoVo);
        for (Map<String, Object> map : list) {

            String lmlx = map.get("路面类型").toString();
            if (lmlx.equals("沥青桥")){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("qmpzdzb","IRI");
                map1.put("qmpzdlmlx","沥青路面");
                map1.put("qmpzdgdz",map.get("设计值"));
                map1.put("qmpzdjcds",map.get("总点数"));
                map1.put("qmpzdhgds",map.get("合格点数"));
                map1.put("qmpzdhgl",map.get("合格率"));
                map1.put("qmpzdbhfw",map.get("Min")+"~"+map.get("Max"));
                resultList.add(map1);
            }
            if (lmlx.equals("混凝土桥")){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("qmpzdzb","IRI");
                map1.put("qmpzdlmlx","混凝土路面");
                map1.put("qmpzdgdz",map.get("设计值"));
                map1.put("qmpzdjcds",map.get("总点数"));
                map1.put("qmpzdhgds",map.get("合格点数"));
                map1.put("qmpzdhgl",map.get("合格率"));
                map1.put("qmpzdbhfw",map.get("Min")+"~"+map.get("Max"));
                resultList.add(map1);
            }
        }
        return resultList;

    }

    /**
     * 表4.1.2-10  路面工程检测结果汇总表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getlmhzData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        Map resultmap = new HashMap<>();
        //处理一下list1，将所有的检测点数和合格点数相加，算合格率
        List<Map<String, Object>> list1 = getlmysdData(commonInfoVo);
        //List<Map<String, Object>> ysd = new ArrayList<>();
        if (CollectionUtils.isNotEmpty(list1)){
            Map mapysd = new HashMap<>();
            double zds = 0;
            double hgds = 0;
            for (Map<String, Object> map : list1) {
                zds += Double.valueOf(map.get("ysdzs").toString());
                hgds += Double.valueOf(map.get("ysdhgs").toString());

            }
            resultmap.put("ysdzds",decf.format(zds));
            resultmap.put("ysdhgds",decf.format(hgds));
            resultmap.put("ysdhgl",zds!=0 ? df.format(hgds/zds*100) : 0);
            resultmap.put("htd",list1.get(0).get("htd"));
            //ysd.add(mapysd);
            //resultList.add(mapysd);
        }
        //弯沉的话，待确认
        List<Map<String, Object>> list2 = getlmwcData(commonInfoVo);
        List<Map<String, Object>> list3 = getlmwclcfData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list2)){
            for (Map<String, Object> map : list2) {
                resultmap.putAll(map);
            }
            //resultList.addAll(list2);
        }else {
            for (Map<String, Object> map : list3) {
                resultmap.putAll(map);
            }
            //resultList.addAll(list3);
        }

        //车辙是分桥 隧道 路面  这块只拿了路面的数据
        List<Map<String, Object>> list4 = getczData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list4)){
            for (Map<String, Object> map : list4) {
                resultmap.putAll(map);
            }
            //resultList.addAll(list4);
        }

        List<Map<String, Object>> list5 = getlmssxsData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list5)){
            for (Map<String, Object> map : list5) {
                resultmap.putAll(map);
            }
            //resultList.addAll(list5);
        }
        List<Map<String, Object>> list6 = gethntqdData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list6)){
            for (Map<String, Object> map : list6) {
                resultmap.putAll(map);
            }
            //resultList.addAll(list6);
        }
        List<Map<String, Object>> list7 = gethntxlbgcData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list7)){
            for (Map<String, Object> map : list7) {
                resultmap.putAll(map);
            }
            //resultList.addAll(list7);
        }
        //平整度需要分沥青和水泥 list8分了 根据lmlx判断
        List<Map<String, Object>> list8 = getpzdData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list8)){
            for (Map<String, Object> map : list8) {
                String hplmlx = map.get("pzdlmlx").toString();
                if (hplmlx.equals("沥青路面")){
                    resultmap.put("pzdlqjcds",map.get("pzdjcds"));
                    resultmap.put("pzdlqhgds",map.get("pzdhgds"));
                    resultmap.put("pzdlqhgl",map.get("pzdhgl"));
                }else if (hplmlx.equals("混凝土路面")){
                    resultmap.put("pzdhntjcds",map.get("pzdjcds"));
                    resultmap.put("pzdhnthgds",map.get("pzdhgds"));
                    resultmap.put("pzdhnthgl",map.get("pzdhgl"));
                }
            }
            //resultList.addAll(list8);
        }
        List<Map<String, Object>> list9 = getkhData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list9)){
            for (Map<String, Object> map : list9) {
                String hplmlx = map.get("khlmlx").toString();
                if (hplmlx.equals("沥青路面")){
                    String khkpzb = map.get("khkpzb").toString();
                    if (khkpzb.equals("摩擦系数")){
                        resultmap.put("khlqmcxsjcds",map.get("khjcds"));
                        resultmap.put("khlqmcxshgds",map.get("khhgds"));
                        resultmap.put("khlqmcxshgl",map.get("khhgl"));
                    }else if (khkpzb.equals("构造深度")){
                        resultmap.put("khlqgzsdjcds",map.get("khjcds"));
                        resultmap.put("khlqgzsdhgds",map.get("khhgds"));
                        resultmap.put("khlqgzsdhgl",map.get("khhgl"));
                    }
                }else if (hplmlx.equals("混凝土路面")){
                    String khkpzb = map.get("khkpzb").toString();
                    if (khkpzb.equals("摩擦系数")){
                        resultmap.put("khhntmcxsjcds",map.get("khjcds"));
                        resultmap.put("khhntmcxshgds",map.get("khhgds"));
                        resultmap.put("khhntmcxshgl",map.get("khhgl"));
                    }else if (khkpzb.equals("构造深度")){
                        resultmap.put("khhntgzsdjcds",map.get("khjcds"));
                        resultmap.put("khhntgzsdhgds",map.get("khhgds"));
                        resultmap.put("khhntgzsdhgl",map.get("khhgl"));
                    }

                }
            }
        }

        List<Map<String, Object>> list10 = gethpData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list10)){
            for (Map<String, Object> map : list10) {
                String hplmlx = map.get("hplmlx").toString();
                if (hplmlx.equals("沥青路面")){
                    resultmap.put("hplqzds",map.get("hpzds"));
                    resultmap.put("hplqhgds",map.get("hphgds"));
                    resultmap.put("hplqhgl",map.get("hphgl"));
                }else if (hplmlx.equals("混凝土路面")){
                    resultmap.put("hphntzds",map.get("hpzds"));
                    resultmap.put("hphnthgds",map.get("hphgds"));
                    resultmap.put("hphnthgl",map.get("hphgl"));
                }
            }
        }
        List<Map<String, Object>> list13 = getlmhdData(commonInfoVo);

        if (CollectionUtils.isNotEmpty(list13)){
            for (Map<String, Object> map : list13) {
                String hplmlx = map.get("lmhdlmlx").toString();
                if (hplmlx.equals("沥青路面")){
                    resultmap.put("lmhdlqjcds",map.get("lmhdjcds"));
                    resultmap.put("lmhdlqhgs",map.get("lmhdhgs"));
                    resultmap.put("lmhdlqhgl",map.get("lmhdhgl"));
                }else if (hplmlx.equals("混凝土路面")){
                    resultmap.put("lmhdhntjcds",map.get("lmhdjcds"));
                    resultmap.put("lmhdhnthgs",map.get("lmhdhgs"));
                    resultmap.put("lmhdhnthgl",map.get("lmhdhgl"));
                }
            }
        }
        resultList.add(resultmap);
        return resultList;
    }


    /**
     * 表4.1.2-8(3) 厚度
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getlmhdData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        //沥青
        List<Map<String, Object>> list = jjgFbgcLmgcGslqlmhdzxfService.lookJdbjg(commonInfoVo);
        if (list.size()>0){
            double jcds = 0;
            double hgds = 0;
            for (Map<String, Object> map : list) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.equals("路面左幅") || lmlx.equals("路面右幅")){
                    jcds += Double.valueOf(map.get("上面层厚度检测点数").toString());
                    jcds += Double.valueOf(map.get("总厚度检测点数").toString());
                    hgds += Double.valueOf(map.get("上面层厚度合格点数").toString());
                    hgds += Double.valueOf(map.get("总厚度合格点数").toString());
                }
            }
            Map map2 = new HashMap();
            map2.put("htd",commonInfoVo.getHtd());
            map2.put("lmhdlmlx","沥青路面");
            map2.put("lmhdjcds",decf.format(jcds));
            map2.put("lmhdhgs",decf.format(hgds));
            map2.put("lmhdhgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
            resultList.add(map2);

        }
        //混凝土
        List<Map<String, Object>> list1 = jjgFbgcLmgcHntlmhdzxfService.lookJdbjg(commonInfoVo);
        if (list1.size()>0){
            Map map2 = new HashMap();
            map2.put("htd",commonInfoVo.getHtd());
            map2.put("lmhdlmlx","混凝土路面");
            map2.put("lmhdjcds",list1.get(0).get("检测点数"));
            map2.put("lmhdhgs",list1.get(0).get("合格点数"));
            map2.put("lmhdhgl",list1.get(0).get("合格率"));
            resultList.add(map2);

        }
        return resultList;

    }

    /**
     * 表4.1.2-8（1）  钻芯法厚度检测结果汇总表
     * 待解决  根据设计值分开，鉴定表生成的时候，要根据设计值分工作簿
     * 钻芯法厚度（总厚度是两行），是匝道的？
     *
     * 关于沥青上面层，需要将桥面左右幅，路面左右幅，和路面匝道的
     * 且设计值相同的检测点数相加，
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getzxfhdData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        //沥青
        List<Map<String, Object>> list = jjgFbgcLmgcGslqlmhdzxfService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list)) {
            List<Map<String, Object>> smc = new ArrayList<>();
            List<Map<String, Object>> zhd = new ArrayList<>();
            for (Map<String, Object> map : list) {
                String lmlx = map.get("路面类型").toString();
                if (!lmlx.contains("隧道")) {
                    Map map1 = new HashMap();
                    Map map2 = new HashMap();
                    map1.put("上面层厚度合格点数", map.get("上面层厚度合格点数").toString());
                    map1.put("上面层平均值最小值", map.get("上面层平均值最小值").toString());
                    map1.put("上面层设计值", map.get("上面层设计值").toString());
                    map1.put("上面层代表值", map.get("上面层代表值").toString());
                    map1.put("上面层厚度合格率", map.get("上面层厚度合格率").toString());
                    map1.put("上面层平均值最大值", map.get("上面层平均值最大值").toString());
                    map1.put("上面层厚度检测点数", map.get("上面层厚度检测点数").toString());
                    map1.put("路面类型", map.get("路面类型").toString());
                    smc.add(map1);

                    map2.put("总厚度合格点数", map.get("总厚度合格点数").toString());
                    map2.put("总厚度平均值最小值", map.get("总厚度平均值最小值").toString());
                    map2.put("总厚度检测点数", map.get("总厚度检测点数").toString());
                    map2.put("总厚度代表值", map.get("总厚度代表值").toString());
                    map2.put("总厚度平均值最大值", map.get("总厚度平均值最大值").toString());
                    map2.put("总厚度设计值", map.get("总厚度设计值").toString());
                    map2.put("总厚度合格率", map.get("总厚度合格率").toString());
                    map2.put("路面类型", map.get("路面类型").toString());
                    zhd.add(map2);
                }
            }
            Map<String, List<Map<String, Object>>> smcgrouped =
                    smc.stream()
                            .collect(Collectors.groupingBy(
                                    item -> item.get("上面层设计值").toString()
                            ));
            Map<String, List<Map<String, Object>>> zhdgrouped =
                    zhd.stream()
                            .collect(Collectors.groupingBy(
                                    item -> item.get("总厚度设计值").toString()
                            ));
            for (List<Map<String, Object>> slist : smcgrouped.values()) {
                double zds = 0;
                double hgds = 0;
                Double sMax = 0.0;
                Double sMin = Double.MAX_VALUE;
                String zzhdbz = "0";
                String yzhdbz = "0";
                Map mapzhd = new HashMap();
                for (Map<String, Object> map : slist) {
                    zds += Double.valueOf(map.get("上面层厚度检测点数").toString());
                    hgds += Double.valueOf(map.get("上面层厚度合格点数").toString());

                    double max = Double.valueOf(map.get("上面层平均值最大值").toString());
                    sMax = (max > sMax) ? max : sMax;

                    double min = Double.valueOf(map.get("上面层平均值最小值").toString());
                    sMin = (min < sMin) ? min : sMin;
                    if (map.get("路面类型").equals("路面左幅")) {
                        zzhdbz = map.get("上面层代表值").toString();
                    }
                    if (map.get("路面类型").equals("路面右幅")) {
                        yzhdbz = map.get("上面层代表值").toString();
                    }
                }
                mapzhd.put("htd",commonInfoVo.getHtd());
                mapzhd.put("zxfhdlmlx","沥青路面");
                mapzhd.put("zxfhdlb","上面层厚度");
                mapzhd.put("zxfhdjcds", zds);
                mapzhd.put("zxfhdhgs", hgds);
                mapzhd.put("zxfhdhgl", zds != 0 ? df.format(hgds / zds * 100) : 0);
                mapzhd.put("zxfhdpjzfw", sMin + "~" + sMax);
                mapzhd.put("zxfhdsjz", slist.get(0).get("上面层设计值"));
                mapzhd.put("zxfhdzdbz", zzhdbz);
                mapzhd.put("zxfhdydbz", yzhdbz);
                resultList.add(mapzhd);
            }

            for (List<Map<String, Object>> zlist : zhdgrouped.values()) {
                double zds = 0;
                double hgds = 0;
                Double sMax = 0.0;
                Double sMin = Double.MAX_VALUE;
                String zzhdbz = "0";
                String yzhdbz = "0";
                Map mapzhd = new HashMap();
                for (Map<String, Object> map : zlist) {
                    zds += Double.valueOf(map.get("总厚度检测点数").toString());
                    hgds += Double.valueOf(map.get("总厚度合格点数").toString());

                    double max = Double.valueOf(map.get("总厚度平均值最大值").toString());
                    sMax = (max > sMax) ? max : sMax;

                    double min = Double.valueOf(map.get("总厚度平均值最小值").toString());
                    sMin = (min < sMin) ? min : sMin;
                    if (map.get("路面类型").equals("路面左幅")) {
                        zzhdbz = map.get("总厚度代表值").toString();
                    }
                    if (map.get("路面类型").equals("路面右幅")) {
                        yzhdbz = map.get("总厚度代表值").toString();
                    }
                }
                mapzhd.put("htd",commonInfoVo.getHtd());
                mapzhd.put("zxfhdlmlx","沥青路面");
                mapzhd.put("zxfhdlb","总厚度");
                mapzhd.put("zxfhdjcds", zds);
                mapzhd.put("zxfhdhgs", hgds);
                mapzhd.put("zxfhdhgl", zds != 0 ? df.format(hgds / zds * 100) : 0);
                mapzhd.put("zxfhdpjzfw", sMin + "~" + sMax);
                mapzhd.put("zxfhdsjz", zlist.get(0).get("总厚度设计值"));
                mapzhd.put("zxfhdzdbz", zzhdbz);
                mapzhd.put("zxfhdydbz", yzhdbz);
                resultList.add(mapzhd);
            }
        }

        //混凝土
        List<Map<String, Object>> list1 = jjgFbgcLmgcHntlmhdzxfService.lookJdbjg(commonInfoVo);
        if (list1.size()>0){
            Map map2 = new HashMap();
            map2.put("htd",commonInfoVo.getHtd());
            map2.put("zxfhdlmlx","混凝土路面");
            map2.put("zxfhdlb","总厚度");
            map2.put("zxfhdpjzfw",list1.get(0).get("最大值")+"~"+list1.get(0).get("最小值"));
            map2.put("zxfhdzdbz",list1.get(0).get("代表值"));
            map2.put("zxfhdydbz",list1.get(0).get("代表值"));
            map2.put("zxfhdsjz",list1.get(0).get("设计值"));
            map2.put("zxfhdjcds",list1.get(0).get("检测点数"));
            map2.put("zxfhdhgs",list1.get(0).get("合格点数"));
            map2.put("zxfhdhgl",list1.get(0).get("合格率"));
            resultList.add(map2);
        }
        return resultList;
    }


    /**
     * 表4.1.2-8（2）  路面雷达厚度检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getldhdData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> list = jjgZdhLdhdService.lookJdbjg(commonInfoVo);
        List<Map<String, Object>> resultList = new ArrayList<>();
        Double sdMax = 0.0;
        Double sdMin = Double.MAX_VALUE;
        double jcds = 0;
        double hgs = 0;
        if (list.size()>0){
            for (Map<String, Object> map : list) {

                String lmlx = map.get("路面类型").toString();
                if (lmlx.contains("左幅") || lmlx.contains("右幅")){
                    double max = Double.valueOf(map.get("Max").toString());
                    sdMax = (max > sdMax) ? max : sdMax;
                    double min = Double.valueOf(map.get("Min").toString());
                    sdMin = (min < sdMin) ? min : sdMin;

                    jcds += Double.valueOf(map.get("总点数").toString());
                    hgs += Double.valueOf(map.get("合格点数").toString());

                }
            }

            Map map1 = new HashMap();
            map1.put("htd",commonInfoVo.getHtd());
            map1.put("ldhdlb","总厚度");
            map1.put("ldhdlmlx","沥青路面");
            map1.put("ldhdsjz",list.get(0).get("设计值"));
            map1.put("ldhdjcds",decf.format(jcds));
            map1.put("ldhdhgds",decf.format(hgs));
            map1.put("ldhdhgl",jcds!=0 ? df.format(hgs/jcds*100) : 0);
            map1.put("ldhdbhfw",sdMin+"~"+sdMax);
            resultList.add(map1);
        }


        return resultList;

    }

    /**
     * 表4.1.2-9 横坡检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> gethpData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, String>> list = jjgFbgcLmgcLmhpService.lookJdbjg(commonInfoVo);
        double lmzds = 0.0;
        double lmhgds = 0.0;

        double lmzds1 = 0.0;
        double lmhgds1 = 0.0;
        if (list.size()>0){
            boolean a = false;
            boolean b = false;
            for (Map<String, String> map : list) {
                String lmlx = map.get("路面类型");
                if (lmlx.contains("沥青路面")){
                    a = true;
                    lmzds +=Double.valueOf(map.get("检测点数"));
                    lmhgds +=Double.valueOf(map.get("合格点数"));

                }
                if (lmlx.contains("混凝土路面")){
                    b = true;
                    lmzds1 +=Double.valueOf(map.get("检测点数"));
                    lmhgds1 +=Double.valueOf(map.get("合格点数"));
                }
            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("hpzds",decf.format(lmzds));
                map1.put("hphgds",decf.format(lmhgds));
                map1.put("hphgl",lmzds!=0 ? df.format(lmhgds/lmzds*100) : 0);
                map1.put("hplmlx","沥青路面");
                resultList.add(map1);
            }
            if (b){
                Map map2 = new HashMap();
                map2.put("htd",commonInfoVo.getHtd());
                map2.put("hpzds",decf.format(lmzds1));
                map2.put("hphgds",decf.format(lmhgds1));
                map2.put("hphgl",lmzds1!=0 ? df.format(lmhgds1/lmzds1*100) : 0);
                map2.put("hplmlx","混凝土路面");
                resultList.add(map2);
            }
        }
        return resultList;

    }


    /**
     * 表4.1.2-7  抗滑检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getkhData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgZdhGzsdService.lookJdbjg(commonInfoVo);//自动化
        String sjz = "";
        double jcds = 0.0;
        double hgds = 0.0;

        if (list.size()>0){
            boolean a = false;
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            for (Map<String, Object> map : list) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.contains("路面")){
                    a = true;
                    sjz = map.get("设计值").toString();
                    jcds += Double.valueOf(map.get("总点数").toString());
                    hgds += Double.valueOf(map.get("合格点数").toString());
                    double max = Double.valueOf(map.get("最大值").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("最小值").toString());
                    lmMin = (min < lmMin) ? min : lmMin;
                }
            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("khlmlx","沥青路面");
                map1.put("khkpzb","构造深度");
                map1.put("khsjz",sjz);
                map1.put("khbhfw",lmMin+"~"+lmMax);
                map1.put("khjcds",decf.format(jcds));
                map1.put("khhgds",decf.format(hgds));
                map1.put("khhgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
                resultList.add(map1);
            }

        }

        List<Map<String, Object>> list1 = jjgFbgcLmgcLmgzsdsgpsfService.lookJdbjg(commonInfoVo);//手工铺沙法
        if (list1.size()>0){
            String gdz = "";
            double szds = 0;
            double shgs = 0;
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            boolean a = false;
            for (Map<String, Object> map : list1) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.equals("混凝土路面")){
                    a = true;
                    szds += Double.valueOf(map.get("检测点数").toString());
                    shgs += Double.valueOf(map.get("合格点数").toString());
                    double max = Double.valueOf(map.get("Max").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("Min").toString());
                    lmMin = (min < lmMin) ? min : lmMin;

                    gdz = map.get("规定值").toString();
                }else {
                    continue;
                }
            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("khlmlx","混凝土路面");
                map1.put("khkpzb","构造深度");
                map1.put("khsjz",gdz);
                map1.put("khbhfw",lmMin+"~"+lmMax);
                map1.put("khjcds",decf.format(szds));
                map1.put("khhgds",decf.format(shgs));
                map1.put("khhgl",szds!=0 ? df.format(shgs/szds*100) : 0);
                resultList.add(map1);
            }

        }
        List<Map<String, Object>> list2 = jjgZdhMcxsService.lookJdbjg(commonInfoVo);
        if (list2.size()>0){
            boolean a = false;
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            double zds = 0;
            double hgs = 0;
            String sjz2 = "";
            for (Map<String, Object> map : list2) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.contains("路面")){
                    a = true;
                    zds += Double.valueOf(map.get("总点数").toString());
                    hgs += Double.valueOf(map.get("合格点数").toString());
                    sjz2 = map.get("设计值").toString();

                    double max = Double.valueOf(map.get("Max").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("Min").toString());
                    lmMin = (min < lmMin) ? min : lmMin;
                }

            }
            if (a){
                Map map1 = new HashMap();
                map1.put("htd",commonInfoVo.getHtd());
                map1.put("khlmlx","沥青路面");
                map1.put("khkpzb","摩擦系数");
                map1.put("khsjz",sjz2);
                map1.put("khbhfw",lmMin+"~"+lmMax);
                map1.put("khjcds",decf.format(zds));
                map1.put("khhgds",decf.format(hgs));
                map1.put("khhgl",zds!=0 ? df.format(hgs/zds*100) : 0);
                resultList.add(map1);
            }

        }
        return resultList;
    }


    /**
     * 表4.1.2-6  平整度检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getpzdData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgZdhPzdService.lookJdbjg(commonInfoVo);
        if (list.size()>0){
            for (Map<String, Object> map : list) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.equals("沥青路面")){
                    Map map1 = new HashMap();
                    map1.put("htd",commonInfoVo.getHtd());
                    map1.put("pzdzb","IRI");
                    map1.put("pzdlmlx","沥青路面");
                    map1.put("pzdgdz",map.get("设计值"));
                    map1.put("pzdjcds",map.get("总点数"));
                    map1.put("pzdhgds",map.get("合格点数"));
                    map1.put("pzdhgl",map.get("合格率"));
                    map1.put("pzdbhfw",map.get("Min")+"~"+map.get("Max"));
                    resultList.add(map1);
                }
                if (lmlx.equals("混凝土路面")){
                    Map map1 = new HashMap();
                    map1.put("htd",commonInfoVo.getHtd());
                    map1.put("pzdzb","IRI");
                    map1.put("pzdlmlx","混凝土路面");
                    map1.put("pzdgdz",map.get("设计值"));
                    map1.put("pzdjcds",map.get("总点数"));
                    map1.put("pzdhgds",map.get("合格点数"));
                    map1.put("pzdhgl",map.get("合格率"));
                    map1.put("pzdbhfw",map.get("Min")+"~"+map.get("Max"));
                    resultList.add(map1);
                }
            }
        }else {
            Map map1 = new HashMap();
            map1.put("htd",commonInfoVo.getHtd());
            map1.put("pzdzb","IRI");
            map1.put("pzdlmlx","混凝土路面");
            map1.put("pzdgdz",0);
            map1.put("pzdjcds",0);
            map1.put("pzdhgds",0);
            map1.put("pzdhgl",0);
            map1.put("pzdbhfw",0);
            resultList.add(map1);
        }
        return resultList;
    }

    /**
     * 表4.1.2-5  混凝土路面相邻板高差检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> gethntxlbgcData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgFbgcLmgcTlmxlbgcService.lookJdbjg(commonInfoVo);
        if (list.size()>0){
            for (Map<String, Object> map : list) {
                map.put("htd",commonInfoVo.getHtd());
                map.put("hntxlbgczds",map.get("总点数"));
                map.put("hntxlbgchgds",map.get("合格点数"));
                map.put("hntxlbgchgl",map.get("合格率"));
                map.put("hntxlbgcgdz",map.get("规定值"));
                map.put("hntxlbgcmax",map.get("max"));
                map.put("hntxlbgcmin",map.get("min"));
                resultList.add(map);
            }
        }
        return resultList;
    }

    /**
     * 表4.1.2-4  混凝土路面强度检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> gethntqdData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgFbgcLmgcHntlmqdService.lookJdbjg(commonInfoVo);
        if (list.size()>0){
            for (Map<String, Object> map : list) {
                map.put("htd",commonInfoVo.getHtd());
                map.put("hntqdzds",map.get("总点数"));
                map.put("hntqdhgds",map.get("合格点数"));
                map.put("hntqdhgl",map.get("合格率"));
                map.put("hntqdgdz",map.get("规定值"));
                map.put("hntqdmax",map.get("最大值"));
                map.put("hntqdmin",map.get("最小值"));
                map.put("hntqdpjz",map.get("平均值"));

                resultList.add(map);
            }
        }
        return resultList;
    }


    /**
     * 表4.1.2-3  渗水系数检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getlmssxsData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.00");
        List<Map<String, Object>> list = jjgFbgcLmgcLmssxsService.lookJdbjg(commonInfoVo);
        double jcds =0.0;
        double hgds =0.0;
        Double lmMax = 0.0;
        Double lmMin = Double.MAX_VALUE;
        boolean a= false;
        if (list.size()>0){
            for (Map<String, Object> map : list) {
                String lmlx = map.get("检测项目").toString();
                if (lmlx.contains("沥青路面")){
                    a = true;
                    jcds += Double.valueOf(map.get("检测点数").toString());
                    hgds += Double.valueOf(map.get("合格点数").toString());

                    double max = Double.valueOf(map.get("最大值").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("最小值").toString());
                    lmMin = (min < lmMin) ? min : lmMin;
                }
            }
            if (a){
                Map<String,Object> map = new HashMap<>();
                map.put("htd",commonInfoVo.getHtd());
                map.put("lmssxsssjcds",jcds);
                map.put("lmssxsmax",lmMax);
                map.put("lmssxsmin",lmMin);
                map.put("lmssxssshgds",hgds);
                map.put("lmssxssshgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
                map.put("lmssxsgdz",list.get(0).get("规定值"));
                resultList.add(map);
            }
        }
        return resultList;

    }


    /**
     * 表4.1.2-6  车辙检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getczData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgZdhCzService.lookJdbjg(commonInfoVo);
        if (list.size()>0){
            double zds =0.0;
            double hgds =0.0;
            Double lmMax = 0.0;
            Double lmMin = Double.MAX_VALUE;
            Map map1 = new HashMap();
            map1.put("htd",commonInfoVo.getHtd());
            map1.put("czzb","MTD");
            map1.put("czlmlx","沥青路面");
            String sjz = "";
            for (Map<String, Object> map : list) {
                String lmlx = map.get("路面类型").toString();
                if (lmlx.contains("路面")){
                    sjz = map.get("设计值").toString();
                    zds += Double.valueOf(map.get("总点数").toString());
                    hgds += Double.valueOf(map.get("合格点数").toString());

                    double max = Double.valueOf(map.get("Max").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("Min").toString());
                    lmMin = (min < lmMin) ? min : lmMin;
                }
            }
            map1.put("czgdz",sjz);
            map1.put("czbhfw",lmMin+"~"+lmMax);
            map1.put("czzds",decf.format(zds));
            map1.put("czhgds",decf.format(hgds));
            map1.put("czhgl",zds!=0 ? df.format(hgds/zds*100) : 0);
            resultList.add(map1);
        }else {
            Map map1 = new HashMap();
            map1.put("htd",commonInfoVo.getHtd());
            map1.put("czzb","MTD");
            map1.put("czlmlx","沥青路面");
            map1.put("czgdz",0);
            map1.put("czbhfw",0);
            map1.put("czzds",0);
            map1.put("czhgds",0);
            map1.put("czhgl",0);
            resultList.add(map1);
        }
        return resultList;
    }


    /**
     * 表4.1.2-2 弯沉(落锤法)
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getlmwclcfData(CommonInfoVo commonInfoVo) throws IOException {

        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgFbgcLmgcLmwcLcfService.lookJdbjg(commonInfoVo);
        if (list.size()>0){
            Map map = new HashMap();
            map.put("htd",commonInfoVo.getHtd());
            map.put("lmwclcfdbz",list.get(0).get("规定值"));
            map.put("lmwclcfmax",getmaxvalue(list.get(0).get("代表值").toString()));
            map.put("lmwclcfmin",getminvalue(list.get(0).get("代表值").toString()));
            map.put("lmwclcfzds",list.get(0).get("检测单元数"));
            map.put("lmwclcfhgds",list.get(0).get("合格单元数"));
            map.put("lmwclcfhgl",list.get(0).get("合格率"));
            resultList.add(map);
        }
        return resultList;

    }


    /**
     * 表4.1.2-2 弯沉(贝尔曼梁法)
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getlmwcData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        List<Map<String, Object>> list = jjgFbgcLmgcLmwcService.lookJdbjg(commonInfoVo);
        if (list.size()>0){
            Map map = new HashMap();
            map.put("htd",commonInfoVo.getHtd());
            map.put("lmwcdbz",list.get(0).get("规定值"));
            map.put("lmwcmax",getmaxvalue(list.get(0).get("代表值").toString()));
            map.put("lmwcmin",getminvalue(list.get(0).get("代表值").toString()));
            map.put("lmwczds",list.get(0).get("检测单元数"));
            map.put("lmwchgds",list.get(0).get("合格单元数"));
            map.put("lmwchgl",list.get(0).get("合格率"));
            resultList.add(map);
        }
        return resultList;
    }


    private String getminvalue(String myList) {
        DecimalFormat df = new DecimalFormat("0.00");
        myList = myList.replace("[", "").replace("]", ""); // 去除方括号
        String[] valus = myList.split(",");
        double min = Integer.MAX_VALUE; // 初始化最大值为最小整数
        for (String num : valus) {
            Double n = Double.valueOf(num);
            if (n < min) {
                min = n;
            }
        }
        return String.valueOf(df.format(min));
    }


    private String getmaxvalue(String myList) {
        DecimalFormat df = new DecimalFormat("0.00");
        myList = myList.replace("[", "").replace("]", ""); // 去除方括号
        String[] valus = myList.split(",");
        double max = Integer.MIN_VALUE; // 初始化最大值为最小整数
        for (String num : valus) {
            Double n = Double.valueOf(num);
            if (n > max) {
                max = n;
            }
        }
        return String.valueOf(df.format(max));
    }

    /**
     * 表4.1.2-1  沥青路面面层压实度检测结果
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getlmysdData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultList = new ArrayList<>();
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String, Object>> list = jjgFbgcLmgcLqlmysdService.lookJdbjg(commonInfoVo);
        Double lmMax = 0.0;
        Double lmMin = Double.MAX_VALUE;
        String gdz = "";
        String zdbz = "";
        String ydbz = "";
        Double jcds = 0.0;
        Double hgds = 0.0;

        Double ljxMax = 0.0;
        Double ljxMin = Double.MAX_VALUE;
        String ljxgdz = "";
        String ljxdbz = "";
        Double ljxjcds = 0.0;
        Double ljxhgds = 0.0;
        boolean a = false;
        boolean b = false;
        if (list.size()>0){
            for (Map<String, Object> map : list) {
                String lx = map.get("路面类型").toString();
                if (lx.equals("沥青路面压实度左幅") || lx.equals("沥青路面压实度右幅")){
                    a = true;
                    double max = Double.valueOf(map.get("最大值").toString());
                    lmMax = (max > lmMax) ? max : lmMax;

                    double min = Double.valueOf(map.get("最小值").toString());
                    lmMin = (min < lmMin) ? min : lmMin;

                    gdz = map.get("规定值").toString();
                    if (lx.equals("沥青路面压实度左幅")){
                        zdbz = map.get("代表值").toString();
                        jcds += Double.valueOf(map.get("检测点数").toString());
                        hgds += Double.valueOf(map.get("合格点数").toString());
                    }else if (lx.equals("沥青路面压实度右幅")){
                        ydbz = map.get("代表值").toString();
                        jcds += Double.valueOf(map.get("检测点数").toString());
                        hgds += Double.valueOf(map.get("合格点数").toString());
                    }

                }else if (lx.equals("沥青路面压实度匝道")){
                    Map map1 = new HashMap<>();
                    map1.put("htd",commonInfoVo.getHtd());
                    map1.put("ysdlx","匝道");
                    map1.put("ysdsczbh",map.get("最小值").toString()+"~"+map.get("最大值").toString());
                    map1.put("ysdzdbz",map.get("代表值").toString());
                    map1.put("ysdydbz",map.get("代表值").toString());
                    map1.put("ysdgdz",map.get("规定值").toString());
                    map1.put("ysdzs",map.get("检测点数").toString());
                    map1.put("ysdhgs",map.get("合格点数").toString());
                    map1.put("ysdhgl",map.get("合格率").toString());
                    resultList.add(map1);

                }else if (lx.contains("连接线")){
                    b = true;
                    double max = Double.valueOf(map.get("最大值").toString());
                    ljxMax = (max > ljxMax) ? max : ljxMax;

                    double min = Double.valueOf(map.get("最小值").toString());
                    ljxMin = (min < ljxMin) ? min : ljxMin;
                    ljxgdz = map.get("规定值").toString();
                    ljxdbz = map.get("代表值").toString();

                    ljxjcds += Double.valueOf(map.get("检测点数").toString());
                    ljxhgds += Double.valueOf(map.get("合格点数").toString());

                }

            }
        }
        if (a){
            Map map1 = new HashMap();
            map1.put("htd",commonInfoVo.getHtd());
            map1.put("ysdlx","主线");
            map1.put("ysdsczbh",lmMin+"~"+lmMax);
            map1.put("ysdzdbz",zdbz);
            map1.put("ysdydbz",ydbz);
            map1.put("ysdgdz",gdz);
            map1.put("ysdzs",decf.format(jcds));
            map1.put("ysdhgs",decf.format(hgds));
            map1.put("ysdhgl",jcds!=0 ? df.format(hgds/jcds*100) : 0);
            resultList.add(map1);
        }
        if (b){
            Map map2 = new HashMap();
            map2.put("htd",commonInfoVo.getHtd());
            map2.put("ysdlx","连接线");
            map2.put("ysdsczbh",ljxMin+"~"+ljxMax);
            map2.put("ysdzdbz",ljxdbz);
            map2.put("ysdydbz",ljxdbz);
            map2.put("ysdgdz",ljxgdz);
            map2.put("ysdzs",decf.format(ljxjcds));
            map2.put("ysdhgs",decf.format(ljxhgds));
            map2.put("ysdhgl",ljxjcds!=0 ? df.format(ljxhgds/ljxjcds*100) : 0);
            resultList.add(map2);
        }
        return resultList;
    }



    /**
     * 表4.1.1-6 路基工程检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getljhzData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> resultlist = new ArrayList<>();
        Map resultmap = new LinkedHashMap();
        List<Map<String, Object>> list1 = gettsfData(commonInfoVo);
        List<Map<String, Object>> list2 = getpsData(commonInfoVo);
        List<Map<String, Object>> list3 = getxqData(commonInfoVo);
        List<Map<String, Object>> list4 = gethdData(commonInfoVo);
        List<Map<String, Object>> list5 = getzdData(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list1)){
            resultmap.putAll(list1.get(0));
        }
        if (CollectionUtils.isNotEmpty(list2)){
            resultmap.putAll(list2.get(0));
        }
        if (CollectionUtils.isNotEmpty(list3)){
            resultmap.putAll(list3.get(0));
        }
        if (CollectionUtils.isNotEmpty(list4)){
            resultmap.putAll(list4.get(0));
        }
        if (CollectionUtils.isNotEmpty(list5)){
            resultmap.putAll(list5.get(0));
        }
        resultlist.add(resultmap);
        return resultlist;
    }


    /**
     * 表4.1.1-5  支挡工程检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> getzdData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> zdlist = new ArrayList<>();
        Map<String, Object> zdmap = new HashMap<>();
        List<Map<String, Object>> list12 = jjgFbgcLjgcZddmccService.lookJdbjg(commonInfoVo);
        if (list12.size()>0){
            zdmap.put("zddmccjcds",list12.get(0).get("检测总点数"));
            zdmap.put("zddmcchgds",list12.get(0).get("合格点数"));
            zdmap.put("zddmcchgl",list12.get(0).get("合格点数"));
        }

        List<Map<String, Object>> list13 = jjgFbgcLjgcZdgqdService.lookJdbjg(commonInfoVo);
        if (list13.size()>0){
            zdmap.put("zdtqdjcds",list13.get(0).get("总点数"));
            zdmap.put("zdtqdhgds",list13.get(0).get("合格点数"));
            zdmap.put("zdtqdhgl",list13.get(0).get("合格点数"));
        }
        zdmap.put("htd", commonInfoVo.getHtd());
        zdmap.put("sheetname","表4.1.1-5");
        zdlist.add(zdmap);
        return zdlist;
    }

    /**
     * 表4.1.1-4  涵洞检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> gethdData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> hdlist = new ArrayList<>();
        Map<String, Object> hdmap = new HashMap<>();
        List<Map<String, Object>> list10 = jjgFbgcLjgcHdgqdService.lookJdbjg(commonInfoVo);
        System.out.println(list10);
        List<Map<String, Object>> list11 = jjgFbgcLjgcHdjgccService.lookJdbjg(commonInfoVo);
        if (list10.size()>0){
            hdmap.put("hdtqdjcds",list10.get(0).get("总点数"));
            hdmap.put("hdtqdhgds",list10.get(0).get("合格点数"));
            hdmap.put("hdtqdhgl",list10.get(0).get("合格率"));
        }
        if (list11.size()>0){
            hdmap.put("hdjgccjcds",list11.get(0).get("总点数"));
            hdmap.put("hdjgcchgds",list11.get(0).get("合格点数"));
            hdmap.put("hdjgcchgl",list11.get(0).get("合格率"));
        }
        hdmap.put("htd", commonInfoVo.getHtd());
        hdmap.put("sheetname","表4.1.1-4");
        hdlist.add(hdmap);
        return hdlist;
    }

    /**
     * 表4.1.1-2  小桥检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>>  getxqData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> xqlist = new ArrayList<>();
        Map<String, Object> xqmap = new HashMap<>();
        List<Map<String, Object>> list8 = jjgFbgcLjgcXqgqdService.lookJdbjg(commonInfoVo);
        if (CollectionUtils.isNotEmpty(list8)){
            xqmap.put("xqtqdjcds",list8.get(0).get("检测总点数"));
            xqmap.put("xqtqdhgds",list8.get(0).get("合格点数"));
            xqmap.put("xqtqdhgl",list8.get(0).get("合格率"));
        }
        List<Map<String, Object>> list9 = jjgFbgcLjgcXqjgccService.lookJdbjg(commonInfoVo);
        if (list9.size()>0){
            xqmap.put("xqjgccjcds",list9.get(0).get("检测总点数"));
            xqmap.put("xqjgcchgds",list9.get(0).get("合格点数"));
            xqmap.put("xqjgcchgl",list9.get(0).get("合格率"));
        }

        xqmap.put("htd", commonInfoVo.getHtd());
        xqmap.put("sheetname","表4.1.1-3");
        xqlist.add(xqmap);
        return xqlist;
    }

    /**
     * 表4.1.1-2  排水工程检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>>  getpsData(CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String, Object>> pslist = new ArrayList<>();
        Map<String, Object> psmap = new LinkedHashMap<>();
        List<Map<String, Object>> list6 = jjgFbgcLjgcPsdmccService.lookJdbjg(commonInfoVo);
        if (list6.size()>0){
            psmap.put("psdmccjcds",list6.get(0).get("检测总点数"));
            psmap.put("psdmcchgds",list6.get(0).get("合格点数"));
            psmap.put("psdmcchgl",list6.get(0).get("合格率"));
        }
        List<Map<String, Object>> list7 = jjgFbgcLjgcPspqhdService.lookJdbjg(commonInfoVo);
        if (list7.size()>0){
            psmap.put("pspqhdjcds",list7.get(0).get("检测总点数"));
            psmap.put("pspqhdhgds",list7.get(0).get("合格点数"));
            psmap.put("pspqhdhgl",list7.get(0).get("合格率"));
        }
        psmap.put("htd", commonInfoVo.getHtd());
        psmap.put("sheetname","表4.1.1-2");
        pslist.add(psmap);
        return pslist;
    }

    /**
     * 表4.1.1-1  路基土石方检测结果汇总表
     * @param commonInfoVo
     * @return
     * @throws IOException
     */
    private List<Map<String, Object>> gettsfData(CommonInfoVo commonInfoVo) throws IOException {
        DecimalFormat df = new DecimalFormat("0.00");
        DecimalFormat decf = new DecimalFormat("0.##");
        List<Map<String,Object>> tsflist = new ArrayList<>();
        Map<String, Object> tsdmap = new HashMap<>();
        List<Map<String, Object>> list1 = jjgFbgcLjgcLjtsfysdHtService.lookJdbjg(commonInfoVo);
        if (list1.size()>0){
            double j = 0;
            double h = 0;
            for (Map<String, Object> map : list1) {
                j += Double.valueOf(map.get("检测点数").toString());
                h += Double.valueOf(map.get("合格点数").toString());

            }
            tsdmap.put("ysdjcds", decf.format(j));
            tsdmap.put("ysdhgds", decf.format(h));
            tsdmap.put("ysdhgl",j!=0 ? df.format(h/j*100) : 0);
        }
        //沉降
        List<Map<String, Object>> list2 = jjgFbgcLjgcLjcjService.lookJdbjg(commonInfoVo);
        if (list2.size()>0){
            for (Map<String, Object> map : list2) {
                tsdmap.put("cjjcds",map.get("总点数"));
                tsdmap.put("cjhgds",map.get("合格点数"));
                tsdmap.put("cjhgl",map.get("合格率"));
            }
        }

        //弯沉
        List<Map<String, Object>> list3 = jjgFbgcLmgcLmwcService.lookJdbjg(commonInfoVo);
        List<Map<String, Object>> list4 = jjgFbgcLjgcLjwcLcfService.lookJdbjg(commonInfoVo);
        double d = Double.valueOf(list3.get(0).get("检测单元数").toString()) + Double.valueOf(list4.get(0).get("检测单元数").toString());
        double dd = Double.valueOf(list3.get(0).get("合格单元数").toString()) + Double.valueOf(list4.get(0).get("合格单元数").toString());
        tsdmap.put("wcjcds", decf.format(d));
        tsdmap.put("wchgds", decf.format(dd));
        tsdmap.put("wchgl",d!=0 ? df.format(dd/d*100) : 0);

        //边坡
        List<Map<String, Object>> list5 = jjgFbgcLjgcLjbpService.lookJdbjg(commonInfoVo);
        if (list5.size()>0){
            tsdmap.put("bpjcds",list5.get(0).get("总点数"));
            tsdmap.put("bphgds",list5.get(0).get("合格点数"));
            tsdmap.put("bphgl",list5.get(0).get("合格率"));
        }

        tsdmap.put("htd", commonInfoVo.getHtd());
        tsdmap.put("sheetname","表4.1.1-1");
        tsflist.add(tsdmap);
        return tsflist;
    }

    /**
     * 表3.4.3-1 交通安全设施抽检统计表
     * @param proname
     * @param htd
     * @return
     */
    private List<Map<String,Object>> getLjjaData(String proname, String htd) {
        List<Map<String,Object>> fhllist = new ArrayList<>();
        Map<String,Object> map = new HashMap<>();
        map.put("htd", htd);
        map.put("sheetname","表3.4.3-1");
        Map<String,Object> map1 = jjgFbgcJtaqssJabzService.selectchs(proname, htd);
        if (map1.size()>0){
            map.put("bzccs",map1.get("wz"));
        }

        Map<String,Object> map2 = jjgFbgcJtaqssJabxService.selectchs(proname, htd);
        if (map2.size()>0){
            map.put("bxccs",map2.get("wz"));
        }

        Map<String,Object> map3 = jjgFbgcJtaqssJathlqdService.selectchs(proname, htd);
        if (map3.size()>0){
            map.put("fhlccs",map3.get("zh"));
        }
        fhllist.add(map);
        return fhllist;
    }

    /**
     * 表3.4.4-1  隧道工程抽检统计表
     * @param commonInfoVo
     * @return
     */
    private List<Map<String, Object>> getsdcjData(CommonInfoVo commonInfoVo) {
        List<Map<String,Object>> resultlist = new ArrayList<>();
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        List<JjgLqsSd> list = jjgLqsSdService.getsdName(proname,htd);
        List<Map<String,Object>> sdcdlist = new ArrayList<>();//[{sdmc=qqq, length=特长隧道}, {sdmc=李家湾隧道左线, length=长隧道}, {sdmc=www, length=特长隧道}, {sdmc=李家湾隧道右线, length=长隧道}]
        if (list.size()>0){
            /**
             * 获取到路桥隧文件中当前项目合同段下的所有隧道
             * 然后判断每个隧道的length,结果存入sdcdlist
             */
            for (JjgLqsSd jjgLqsSd : list) {
                Map map = new HashMap();
                String sdqc = jjgLqsSd.getSdqc();
                String sdname = jjgLqsSd.getSdname();
                Double sdc = Double.valueOf(sdqc);
                if (sdc >= 3000){
                    map.put("length","特长隧道");

                }else if (sdc < 3000 && sdc >= 1000){
                    map.put("length","长隧道");

                }else if (sdc < 1000 && sdc >= 500){
                    map.put("length","中隧道");

                }else if (sdc < 500){
                    map.put("length","短隧道");
                }
                map.put("sdmc",sdname);
                sdcdlist.add(map);
            }
            /**
             * 根据length属性的值分组，然后得到分组后每个map的value长度，也就是某个length的数量
             */
            Map<String, List<Map<String, Object>>> groupedData =
                    sdcdlist.stream()
                            .collect(Collectors.groupingBy(
                                    item -> item.get("length").toString()
                            ));
            Map<String, Object> keyLengths = new HashMap<>();//{长隧道=2, 特长隧道=2}  实际的 路桥遂文件中的
            groupedData.forEach((key, value) -> keyLengths.put(key, value.size()));

            /**
             * 获取实测数据中的隧道
             */
            List<Map<String,Object>> sdnumList = jjgFbgcSdgcCqtqdService.getsdnum(proname,htd);//[{sdmc=李家湾隧道左线}, {sdmc=李家湾隧道右线}]

            Map resultMap = new HashMap<>();
            resultMap.put("htd",htd);
            resultMap.put("sheetname","表3.4.4-1");

            if (sdnumList.size()>0){
                /**
                 * 先判断实测数据中每个隧道的length,将结果存入到cjnum
                 */
                List<Map<String,Object>> cjnum = new ArrayList<>();
                for (Map<String, Object> cjmap : sdnumList) {//实测数据的
                    String cjsdmc = cjmap.get("sdmc").toString();
                    for (Map<String, Object> symap : sdcdlist) {//路桥隧的
                        String sysdmc = symap.get("sdmc").toString();
                        if (cjsdmc.equals(sysdmc)){
                            Map map = new HashMap();
                            map.put("sdmc",sysdmc);
                            map.put("length",symap.get("length"));
                            cjnum.add(map);
                        }
                    }
                }
                if (cjnum.size()>0){
                    Map<String, List<Map<String, Object>>> cjgroupedData =
                            cjnum.stream()
                                    .collect(Collectors.groupingBy(
                                            item -> item.get("length").toString()
                                    ));
                    Map<String, Object> cjkeyLengths = new HashMap<>();//{长隧道=2} 抽检的 实测文件中的
                    /**
                     * 然后就得到了实测数据中隧道的length情况，cjgroupedData
                     */
                    cjgroupedData.forEach((key, value) -> cjkeyLengths.put(key, value.size()));

                    for (String key : cjkeyLengths.keySet()) {
                        if (key.equals("特长隧道")){
                            resultMap.put("tccjs",cjkeyLengths.get(key));
                        }else {
                            resultMap.put("tccjs",0);
                        }

                        if (key.equals("长隧道")){
                            resultMap.put("ccjs",cjkeyLengths.get(key));
                        }else {
                            resultMap.put("ccjs",0);
                        }

                        if (key.equals("中隧道")){
                            resultMap.put("zcjs",cjkeyLengths.get(key));
                        }else {
                            resultMap.put("zcjs",0);
                        }
                        if (key.equals("短隧道")){
                            resultMap.put("dcjs",cjkeyLengths.get(key));
                        }else {
                            resultMap.put("dcjs",0);
                        }

                    }
                }
            }
            if (keyLengths.size()>0){
                for (String key : keyLengths.keySet()) {
                    if (key.equals("特长隧道")){
                        resultMap.put("tcccs",keyLengths.get(key));
                    }else {
                        resultMap.put("tcccs",0);
                    }

                    if (key.equals("长隧道")){
                        resultMap.put("cccs",keyLengths.get(key));
                    }else {
                        resultMap.put("cccs",0);
                    }

                    if (key.equals("中隧道")){
                        resultMap.put("zccs",keyLengths.get(key));
                    }else {
                        resultMap.put("zccs",0);
                    }
                    if (key.equals("短隧道")){
                        resultMap.put("dccs",keyLengths.get(key));
                    }else {
                        resultMap.put("dccs",0);
                    }

                }
            }
            resultlist.add(resultMap);
        }
        return resultlist;
    }

    /**
     * 表3.4.2-1  桥梁工程抽检统计表
     * @param proname
     * @param htd
     * @return
     */
    private List<Map<String,Object>>  getLjqlcjData(String proname, String htd) {
        List<Map<String,Object>> qllist = new ArrayList<>();
        List<JjgLqsQl> list = jjgLqsQlService.getqlName(proname, htd);
        List<Map<String,Object>> qlxh = new ArrayList<>();
        if (list.size()>0){
            for (JjgLqsQl jjgLqsQl : list) {
                Map<String,Object> map = new HashMap<>();
                String qlname = jjgLqsQl.getQlname();
                String dkkj = jjgLqsQl.getDkkj();
                if (dkkj == null || dkkj.equals("")){
                    dkkj = "5";
                }
                Double kj = Double.valueOf(dkkj);
                if (kj >= 5 && kj < 20){
                    map.put("length","小桥");
                }else if (kj >= 20 && kj < 40){
                    map.put("length","中桥");
                }else if (kj >= 40 && kj <= 150){
                    map.put("length","大桥");
                }else if (kj > 150){
                    map.put("length","特大桥");
                }
                map.put("qlname",qlname);
                qlxh.add(map);
            }

            Map<String, List<Map<String, Object>>> groupedData =
                    qlxh.stream()
                            .collect(Collectors.groupingBy(
                                    item -> item.get("length").toString()
                            ));
            Map<String, Object> keyLengths = new HashMap<>();//路桥遂文件中的
            groupedData.forEach((key, value) -> keyLengths.put(key, value.size()));
            List<Map<String,Object>> qlnumList = jjgFbgcQlgcSbTqdService.getqlname(proname,htd);
            List<Map<String,Object>> cjnum = new ArrayList<>();

            Map resultMap = new HashMap<>();
            resultMap.put("htd",htd);
            resultMap.put("sheetname","表3.4.2-1");
            if (qlnumList.size()>0){
                for (Map<String, Object> cjmap : qlnumList) {//实测数据的
                    String cjqlmc = cjmap.get("qlmc").toString();
                    for (Map<String, Object> symap : qlxh) {//路桥隧的
                        String syqlmc = symap.get("qlname").toString();
                        if (cjqlmc.equals(syqlmc)){
                            Map map = new HashMap();
                            map.put("qlmc",syqlmc);
                            map.put("length",symap.get("length"));
                            cjnum.add(map);
                        }
                    }
                }
                if (cjnum.size()>0){
                    Map<String, List<Map<String, Object>>> cjgroupedData =
                            cjnum.stream()
                                    .collect(Collectors.groupingBy(
                                            item -> item.get("length").toString()
                                    ));
                    Map<String, Object> cjkeyLengths = new HashMap<>();
                    cjgroupedData.forEach((key, value) -> cjkeyLengths.put(key, value.size()));
                    for (String key : cjkeyLengths.keySet()) {
                        if (key.equals("特大桥")){
                            resultMap.put("tdcjs",cjkeyLengths.get(key));
                        }else {
                            resultMap.put("tdcjs",0);
                        }

                        if (key.equals("大桥")){
                            resultMap.put("dcjs",cjkeyLengths.get(key));
                        }else {
                            resultMap.put("dcjs",0);
                        }

                        if (key.equals("中桥")){
                            resultMap.put("zcjs",cjkeyLengths.get(key));
                        }else {
                            resultMap.put("zcjs",0);
                        }
                        if (key.equals("小桥")){
                            resultMap.put("xcjs",cjkeyLengths.get(key));
                        }else {
                            resultMap.put("xcjs",0);
                        }

                    }
                }


            }
            if (keyLengths.size()>0){
                for (String key : keyLengths.keySet()) {
                    if (key.equals("特大桥")){
                        resultMap.put("tdccs",keyLengths.get(key));
                    }else {
                        resultMap.put("tdccs",0);
                    }

                    if (key.equals("大桥")){
                        resultMap.put("dccs",keyLengths.get(key));
                    }else {
                        resultMap.put("dccs",0);
                    }

                    if (key.equals("中桥")){
                        resultMap.put("zccs",keyLengths.get(key));
                    }else {
                        resultMap.put("zccs",0);
                    }
                    if (key.equals("小桥")){
                        resultMap.put("xccs",keyLengths.get(key));
                    }else {
                        resultMap.put("xccs",0);
                    }

                }
            }
            qllist.add(resultMap);
        }
        return qllist;
    }

    /**
     * 表3.4.1-1 涵洞、支挡工程抽检统计表
     * 实有数是用户自己输入，但是这块还要给抽检频率插入公式
     * @param proname
     * @param htd
     * @return
     */
    private List<Map<String,Object>> getLjcjtjbData(String proname, String htd) {
        List<Map<String,Object>> ljlist = new ArrayList<>();
        Map<String,Object> map = new LinkedHashMap<>();
        Map<String,Object> map1 = jjgFbgcLjgcHdgqdService.selectchs(proname, htd);
        map.put("htd", htd);
        map.put("hdccs",map1.get("ccs"));
        //支挡工程
        Map<String,Object> map2 = jjgFbgcLjgcZddmccService.selectchs(proname, htd);
        map.put("zdccs",map2.get("ccs"));
        //小桥
        Map<String,Object> map3 = jjgFbgcLjgcXqgqdService.selectchs(proname, htd);
        map.put("xqccs",map3.get("ccs"));
        map.put("sheetname","表3.4.1-1");
        ljlist.add(map);
        return ljlist;
    }

}
