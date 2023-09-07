package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.*;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.zdh.JjgZdhPzdVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.*;
import glgc.jjgys.system.service.JjgZdhPzdService;
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
 * @since 2023-06-17
 */
@Service
public class JjgZdhPzdServiceImpl extends ServiceImpl<JjgZdhPzdMapper, JjgZdhPzd> implements JjgZdhPzdService {

    @Autowired
    private JjgZdhPzdMapper jjgZdhPzdMapper;

    @Autowired
    private JjgHtdMapper jjgHtdMapper;

    @Autowired
    private JjgLqsSdMapper jjgLqsSdMapper;

    @Autowired
    private JjgLqsQlMapper jjgLqsQlMapper;

    @Autowired
    private JjgLqsHntlmzdMapper jjgLqsHntlmzdMapper;

    @Autowired
    private JjgLqsSfzMapper jjgLqsSfzMapper;

    @Autowired
    private JjgLqsLjxMapper jjgLqsLjxMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();

        List<Map<String,Object>> lxlist = jjgZdhPzdMapper.selectlx(proname,htd);
        for (Map<String, Object> map : lxlist) {
            String zx = map.get("lxbs").toString();
            int num = jjgZdhPzdMapper.selectcdnum(proname,htd,zx);
            int cds = 0;
            if (num == 1){
                cds = 2;
            }else {
                cds=num;
            }
            handlezxData(proname,htd,zx,cds,commonInfoVo.getSjz());
        }

    }

    /**
     *
     * @param proname
     * @param htd
     * @param zx
     * @param cdsl
     * @param sjz
     */
    private void handlezxData(String proname, String htd, String zx, int cdsl, String sjz) throws IOException, ParseException {
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


            List<Map<String,Object>> datazf = jjgZdhPzdMapper.selectzfList(proname,htd,zx,result);
            List<Map<String,Object>> datayf = jjgZdhPzdMapper.selectyfList(proname,htd,zx,result);

            /**
             * 查询合同的的起止桩号
             */
            QueryWrapper<JjgHtd> wrapperhtd = new QueryWrapper<>();
            wrapperhtd.like("proname",proname);
            wrapperhtd.like("name",htd);
            List<JjgHtd> htdList = jjgHtdMapper.selectList(wrapperhtd);

            String htdzhq = htdList.get(0).getZhq();
            String htdzhz = htdList.get(0).getZhz();

            //隧道
            List<JjgLqsSd> jjgLqsSdzf = jjgLqsSdMapper.selectsdzf(proname,htdzhq,htdzhz,"左幅");
            List<JjgLqsSd> jjgLqsSdyf = jjgLqsSdMapper.selectsdyf(proname,htdzhq,htdzhz,"右幅");

            //桥
            List<JjgLqsQl> jjgLqsQlzf = jjgLqsQlMapper.selectqlzf(proname,htdzhq,htdzhz,"左幅");
            List<JjgLqsQl> jjgLqsQlyf = jjgLqsQlMapper.selectqlyf(proname,htdzhq,htdzhz,"右幅");


            List<Map<String,Object>> sdzfdata = new ArrayList<>();
            if (jjgLqsSdzf.size()>0){
                for (JjgLqsSd jjgLqsSd : jjgLqsSdzf) {
                    String zhq = String.valueOf(jjgLqsSd.getZhq());
                    String zhz = String.valueOf(jjgLqsSd.getZhz());
                    sdzfdata.addAll(jjgZdhPzdMapper.selectSdZfData(proname,htd,zx,zhq,zhz,result));
                }
            }


            List<Map<String,Object>> sdyfdata = new ArrayList<>();
            if (jjgLqsSdyf.size()>0){
                for (JjgLqsSd jjgLqsSd : jjgLqsSdyf) {
                    String zhq = String.valueOf(jjgLqsSd.getZhq());
                    String zhz = String.valueOf(jjgLqsSd.getZhz());
                    sdyfdata.addAll(jjgZdhPzdMapper.selectSdyfData(proname,htd,zx,zhq,zhz,result));
                }
            }
            List<Map<String,Object>> qlzfdata = new ArrayList<>();
            if (jjgLqsQlzf.size()>0){
                for (JjgLqsQl jjgLqsQl : jjgLqsQlzf) {
                    String zhq = String.valueOf(jjgLqsQl.getZhq());
                    String zhz = String.valueOf(jjgLqsQl.getZhz());
                    qlzfdata.addAll(jjgZdhPzdMapper.selectQlZfData(proname,htd,zx,zhq,zhz,result));
                }
            }
            List<Map<String,Object>> qlyfdata = new ArrayList<>();
            if (jjgLqsQlyf.size()>0){
                for (JjgLqsQl jjgLqsQl : jjgLqsQlyf) {
                    String zhq = String.valueOf(jjgLqsQl.getZhq());
                    String zhz = String.valueOf(jjgLqsQl.getZhz());
                    qlyfdata.addAll(jjgZdhPzdMapper.selectQlYfData(proname,htd,zx,zhq,zhz,result));
                }
            }

            List<Map<String, Object>> sdzxList = montageIRI(sdzfdata);
            List<Map<String, Object>> sdyxList = montageIRI(sdyfdata);
            List<Map<String, Object>> qlzxList = montageIRI(qlzfdata);
            List<Map<String, Object>> qlyxList = montageIRI(qlyfdata);

            List<Map<String, Object>> lmzfList = montageIRI(datazf);
            List<Map<String, Object>> lmyfList = montageIRI(datayf);

            double zdzh = Double.parseDouble(lmzfList.get(0).get("qdzh").toString());
            double finzdzh = Double.parseDouble(lmzfList.get(lmzfList.size()-1).get("qdzh").toString());
            List<Map<String, Object>> lmzf = decrementNumberByStep(zdzh,finzdzh,lmzfList,cdsl);

            double yzdzh = Double.parseDouble(lmyfList.get(0).get("qdzh").toString());
            double yfinzdzh = Double.parseDouble(lmyfList.get(lmyfList.size()-1).get("qdzh").toString());
            List<Map<String, Object>> lmyf = decrementNumberByStep(yzdzh,yfinzdzh,lmyfList,cdsl);

            writeExcelData(proname,htd,lmzf,lmyf,sdzxList,sdyxList,qlzxList,qlyxList,cdsl,sjz,zx);
        }else if (zx.contains("连接线")){
            //查询的是摩擦系数表中的连接线
            List<Map<String,Object>> dataljxzf = jjgZdhPzdMapper.selectzfList(proname,htd,zx,result);
            List<Map<String,Object>> dataljxyf = jjgZdhPzdMapper.selectyfList(proname,htd,zx,result);
            //连接线
            QueryWrapper<JjgLjx> wrapperljx = new QueryWrapper<>();
            wrapperljx.like("proname",proname);
            wrapperljx.like("sshtd",htd);
            List<JjgLjx> jjgLjxList = jjgLqsLjxMapper.selectList(wrapperljx);

            List<Map<String,Object>> sdpzd = new ArrayList<>();
            List<Map<String,Object>> qlpzd = new ArrayList<>();

            for (JjgLjx jjgLjx : jjgLjxList) {
                String zhq = jjgLjx.getZhq();
                String zhz = jjgLjx.getZhz();
                String bz = jjgLjx.getBz();
                String ljxlf = jjgLjx.getLf();
                String wz = jjgLjx.getLjxname();
                List<JjgLqsSd> jjgLqssd = jjgLqsSdMapper.selectsdList(proname,zhq,zhz,bz,wz,ljxlf);
                for (JjgLqsSd jjgLqsSd : jjgLqssd) {
                    String lf = jjgLqsSd.getLf();

                    Double sdq = jjgLqsSd.getZhq()+10;
                    String sdzhq = String.valueOf(sdq);
                    String sdzhz = jjgLqsSd.getZhz().toString();

                    String zhq1 = String.valueOf((jjgLqsSd.getZhq()));
                    String zhz1 = String.valueOf((jjgLqsSd.getZhz()));
                    sdpzd.addAll(jjgZdhPzdMapper.selectsdpzd1(proname,bz,lf,zx,zhq1,zhz1,sdzhq,sdzhz));
                }

                List<JjgLqsQl> jjgLqsql = jjgLqsQlMapper.selectqlList(proname,zhq,zhz,bz,wz,ljxlf);
                for (JjgLqsQl jjgLqsQl : jjgLqsql) {
                    String lf = jjgLqsQl.getLf();
                    Double qlq = jjgLqsQl.getZhq()+10;
                    Double qlz = jjgLqsQl.getZhz();

                    String qlzhq = String.valueOf(qlq);
                    String qlzhz = String.valueOf(qlz);

                    String zhq1 = String.valueOf(jjgLqsQl.getZhq());
                    String zhz1 = String.valueOf(jjgLqsQl.getZhz());
                    qlpzd.addAll(jjgZdhPzdMapper.selectqlpzd1(proname,bz,lf, qlzhq, qlzhz, zx, zhq1, zhz1));
                }
            }
            List<Map<String,Object>> zfqlpzd = new ArrayList<>();
            List<Map<String,Object>> yfqlpzd = new ArrayList<>();
            if (qlpzd.size()>0){
                for (int i = 0; i < qlpzd.size(); i++) {
                    if (qlpzd.get(i).get("cd").toString().contains("左幅")){
                        zfqlpzd.add(qlpzd.get(i));
                    }
                    if (qlpzd.get(i).get("cd").toString().contains("右幅")){
                        yfqlpzd.add(qlpzd.get(i));
                    }
                }
            }
            List<Map<String,Object>> zfsdpzd = new ArrayList<>();
            List<Map<String,Object>> yfsdpzd = new ArrayList<>();
            if (sdpzd.size()>0){
                for (Map<String, Object> sdmcx : sdpzd) {
                    if (sdmcx.get("cd").toString().contains("左幅")){
                        zfsdpzd.add(sdmcx);
                    }
                    if (sdmcx.get("cd").toString().contains("右幅")){
                        yfsdpzd.add(sdmcx);
                    }
                }
            }

            //拼接IRI数据
            List<Map<String, Object>> sdzList = montageZDIRI(zfsdpzd,cdsl);
            List<Map<String, Object>> sdyList = montageZDIRI(yfsdpzd,cdsl);
            List<Map<String, Object>> qlzList = montageZDIRI(zfqlpzd,cdsl);

            List<Map<String, Object>> qlyList = montageZDIRI(yfqlpzd,cdsl);

            List<Map<String, Object>> zdzfList = montageZDIRI(dataljxzf,cdsl);
            List<Map<String, Object>> zdyfList = montageZDIRI(dataljxyf,cdsl);


            List<Map<String, Object>> zdzf = new ArrayList<>();
            double zdzh = Double.parseDouble(zdzfList.get(0).get("qdzh").toString());
            double finzdzh = Double.parseDouble(zdzfList.get(zdzfList.size()-1).get("qdzh").toString());;
            zdzf.addAll(decrementNumberByStep(zdzh,finzdzh,zdzfList,cdsl));

            Collections.sort(zdzf, new Comparator<Map<String, Object>>() {
                @Override
                public int compare(Map<String, Object> o1, Map<String, Object> o2) {
                    // 名字相同时按照 qdzh 排序
                    Double qdzh1 = Double.parseDouble(o1.get("qdzh").toString());
                    Double qdzh2 = Double.parseDouble(o2.get("qdzh").toString());
                    return qdzh1.compareTo(qdzh2);
                }
            });

            List<Map<String, Object>> zdyf = new ArrayList<>();
            double zdzh1 = Double.parseDouble(zdyfList.get(0).get("qdzh").toString());
            double finzdzh1 = Double.parseDouble(zdyfList.get(zdyfList.size()-1).get("qdzh").toString());;
            zdyf.addAll(decrementNumberByStep(zdzh1,finzdzh1,zdyfList,cdsl));
            Collections.sort(zdyf, new Comparator<Map<String, Object>>() {
                @Override
                public int compare(Map<String, Object> o1, Map<String, Object> o2) {
                    // 名字相同时按照 qdzh 排序
                    Double qdzh1 = Double.parseDouble(o1.get("qdzh").toString());
                    Double qdzh2 = Double.parseDouble(o2.get("qdzh").toString());
                    return qdzh1.compareTo(qdzh2);
                }
            });
            writeExcelData(proname,htd,zdzf,zdyf,sdzList,sdyList,qlzList,qlyList,cdsl,sjz,zx);

        }else {
            //匝道的所有数据
            List<Map<String,Object>> datazdzf = jjgZdhPzdMapper.selectzfList(proname,htd,zx,result);
            List<Map<String,Object>> datazdyf = jjgZdhPzdMapper.selectyfList(proname,htd,zx,result);

            //匝道表中的数据
            QueryWrapper<JjgLqsHntlmzd> wrapperzd = new QueryWrapper<>();
            wrapperzd.like("proname",proname);
            wrapperzd.like("wz",zx);
            List<JjgLqsHntlmzd> zdList = jjgLqsHntlmzdMapper.selectList(wrapperzd);

            List<Map<String,Object>> sdpzd = new ArrayList<>();//隧道的实测数据
            List<Map<String,Object>> qlpzd = new ArrayList<>();//桥梁的实测数据
            for (JjgLqsHntlmzd jjgLqsHntlmzd : zdList) {
                String zhq = jjgLqsHntlmzd.getZhq();
                String zhz = jjgLqsHntlmzd.getZhz();
                String bz = jjgLqsHntlmzd.getZdlx();
                String wz = jjgLqsHntlmzd.getWz();
                String zdlf = jjgLqsHntlmzd.getLf();

                List<JjgLqsSd> jjgLqssd = jjgLqsSdMapper.selectsdList(proname,zhq,zhz,bz,wz,zdlf);
                if (jjgLqssd.size()>0){
                    for (JjgLqsSd jjgLqsSd : jjgLqssd) {
                        String lf = jjgLqsSd.getLf();
                        Double sdz = jjgLqsSd.getZhz()-100;
                        String sdzhz = String.valueOf(sdz);

                        String zhq1 = String.valueOf((jjgLqsSd.getZhq()));
                        String zhz1 = String.valueOf((jjgLqsSd.getZhz()));
                        sdpzd.addAll(jjgZdhPzdMapper.selectsdpzd(proname,bz,lf,zx,zhq1,zhz1,sdzhz));
                    }

                }

                List<JjgLqsQl> jjgLqsql = jjgLqsQlMapper.selectqlList(proname,zhq,zhz,bz,wz,zdlf);

                if (jjgLqsql.size()>0){
                    for (JjgLqsQl jjgLqsQl : jjgLqsql) {
                        String lf = jjgLqsQl.getLf();
                        Double qlz = jjgLqsQl.getZhz()-100;
                        String qlzhzj = String.valueOf(qlz);
                        String qlzhq = String.valueOf(jjgLqsQl.getZhq());
                        String qlzhz = String.valueOf(jjgLqsQl.getZhz());
                        qlpzd.addAll(jjgZdhPzdMapper.selectqlpzd(proname,bz,lf, qlzhq, qlzhz, zx,qlzhzj));

                    }
                }
            }

            List<Map<String,Object>> zfqlpzd = new ArrayList<>();
            List<Map<String,Object>> yfqlpzd = new ArrayList<>();
            if (qlpzd.size()>0){
                for (int i = 0; i < qlpzd.size(); i++) {
                    if (qlpzd.get(i).get("cd").toString().contains("左幅")){
                        zfqlpzd.add(qlpzd.get(i));
                    }
                    if (qlpzd.get(i).get("cd").toString().contains("右幅")){
                        yfqlpzd.add(qlpzd.get(i));
                    }
                }
            }
            List<Map<String,Object>> zfsdpzd = new ArrayList<>();
            List<Map<String,Object>> yfsdpzd = new ArrayList<>();
            if (sdpzd.size()>0){
                for (Map<String, Object> sdmcx : sdpzd) {
                    if (sdmcx.get("cd").toString().contains("左幅")){
                        zfsdpzd.add(sdmcx);
                    }
                    if (sdmcx.get("cd").toString().contains("右幅")){
                        yfsdpzd.add(sdmcx);
                    }
                }
            }


            //拼接IRI数据
            List<Map<String, Object>> sdzList = montageZDIRI(zfsdpzd,cdsl);
            List<Map<String, Object>> sdyList = montageZDIRI(yfsdpzd,cdsl);
            List<Map<String, Object>> qlzList = montageZDIRI(zfqlpzd,cdsl);

            List<Map<String, Object>> qlyList = montageZDIRI(yfqlpzd,cdsl);

            List<Map<String, Object>> zdzfList = montageZDIRI(datazdzf,cdsl);
            List<Map<String, Object>> zdyfList = montageZDIRI(datazdyf,cdsl);


            List<Map<String, Object>> zdzf = new ArrayList<>();

            Map<String, List<Map<String, Object>>> zdzfgroup = zdzfList.stream()
                    .collect(Collectors.groupingBy(
                            m -> Objects.toString(m.get("zdbs"), ""),
                            Collectors.collectingAndThen(Collectors.toList(), list -> {
                                list.sort(Comparator.comparing(m ->Objects.toString(m.get("qdzh"), "")));
                                return list;
                            })));
            for (Map.Entry<String, List<Map<String, Object>>> entry : zdzfgroup.entrySet()) {
                List<Map<String, Object>> value = entry.getValue();

                double zdzh = Double.parseDouble(value.get(0).get("qdzh").toString());
                double finzdzh = Double.parseDouble(value.get(value.size()-1).get("qdzh").toString());
                zdzf.addAll(decrementNumberByStep(zdzh,finzdzh,value,cdsl));
            }

            List<Map<String, Object>> zdyf = new ArrayList<>();
            Map<String, List<Map<String, Object>>> zdyfgroup = zdyfList.stream()
                    .collect(Collectors.groupingBy(
                            m -> Objects.toString(m.get("zdbs"), ""),
                            Collectors.collectingAndThen(Collectors.toList(), list -> {
                                list.sort(Comparator.comparing(m ->Objects.toString(m.get("qdzh"), "")));
                                return list;
                            })));
            for (Map.Entry<String, List<Map<String, Object>>> entry : zdyfgroup.entrySet()) {
                List<Map<String, Object>> value = entry.getValue();

                double zdzh = Double.parseDouble(value.get(0).get("qdzh").toString());
                double finzdzh = Double.parseDouble(value.get(value.size()-1).get("qdzh").toString());
                zdyf.addAll(decrementNumberByStep(zdzh,finzdzh,value,cdsl));
            }

            Collections.sort(zdzf, new Comparator<Map<String, Object>>() {
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
                    Double qdzh1 = Double.parseDouble(o1.get("qdzh").toString());
                    Double qdzh2 = Double.parseDouble(o2.get("qdzh").toString());
                    return qdzh1.compareTo(qdzh2);
                }
            });

            Collections.sort(zdyf, new Comparator<Map<String, Object>>() {
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
                    Double qdzh1 = Double.parseDouble(o1.get("qdzh").toString());
                    Double qdzh2 = Double.parseDouble(o2.get("qdzh").toString());
                    return qdzh1.compareTo(qdzh2);
                }
            });

            writeExcelData(proname,htd,zdzf,zdyf,sdzList,sdyList,qlzList,qlyList,cdsl,sjz,zx);
        }

    }

    /**
     * 拼接匝道的IRI
     * @param list
     * @return
     */
    private static List<Map<String, Object>> montageZDIRI(List<Map<String, Object>> list,int cdsl) {
        if (list == null || list.isEmpty()){
            return new ArrayList<>();
        }else {
            Map<String, List<String>> resultMapz = new TreeMap<>();
            for (Map<String, Object> map : list) {
                String qdzh = map.get("qdzh").toString();
                String ziri = map.get("ziri").toString();
                String yiri = map.get("yiri").toString();
                String zdbs = map.get("zdbs").toString();
                String key = qdzh + "_" + zdbs; // 添加zdbs的唯一标识，防止重复覆盖
                if (resultMapz.containsKey(key) ) {
                    resultMapz.get(key).add(ziri);
                    resultMapz.get(key).add(yiri);
                } else {
                    List<String> sfcList = new ArrayList<>();
                    sfcList.add(ziri);
                    sfcList.add(yiri);
                    resultMapz.put(key, sfcList);
                    //resultMapz.put(qdzh, sfcList);
                }
            }

            List<Map<String, Object>> resultList = new LinkedList<>();
            for (Map.Entry<String, List<String>> entry : resultMapz.entrySet()) {
                Map<String, Object> map = new TreeMap<>();
                //map.put("qdzh", entry.getKey());
                map.put("iri", String.join(",", entry.getValue()));
                for (Map<String, Object> item : list) {
                    String itemkey = item.get("qdzh").toString() + "_" + item.get("zdbs").toString();
                    if (itemkey.equals(entry.getKey())) {
                        map.put("qdzh", item.get("qdzh"));
                        map.put("zdzh", item.get("zdzh"));
                        map.put("pzlx", item.get("pzlx"));
                        map.put("name", item.get("name"));
                        map.put("zdbs", item.get("zdbs"));
                        map.put("cd", item.get("cd").toString());
                        map.put("createTime", item.get("createTime"));
                        break;
                    }
                }
                resultList.add(map);
            }
            String iriss = "-";
            StringBuilder iriBuilder = new StringBuilder();
            for (int i = 0; i < cdsl; i++) {
                if (i == cdsl - 1) {
                    iriBuilder.append("-");
                } else {
                    iriBuilder.append("-,");
                }
            }
            iriss = iriBuilder.toString();
            for (Map<String, Object> map : resultList) {
                String[] iris = map.get("iri").toString().split(",");
                String cd = map.get("cd").toString();
                if (iris.length < cdsl*2){
                    if (cd.equals("左幅一车道") || cd.equals("右幅一车道")){
                        map.put("iri",map.get("iri").toString()+","+iriss);
                    }else if (cd.equals("左幅二车道") || cd.equals("右幅二车道")){
                        map.put("iri",iriss+","+map.get("iri").toString());
                    }
                    map.put("cd",map.get("cd").toString().substring(0,2));
                }else {
                    map.put("cd",map.get("cd").toString().substring(0,2));
                }
            }
            resultList.sort((o1, o2) -> {
                String zdbs1 = o1.get("zdbs").toString();
                String zdbs2 = o2.get("zdbs").toString();
                int compareZdbs = zdbs1.compareTo(zdbs2);
                if (compareZdbs == 0) {
                    String qdzh1 = o1.get("qdzh").toString();
                    String qdzh2 = o2.get("qdzh").toString();
                    return qdzh1.compareTo(qdzh2);
                } else {
                    return compareZdbs;
                }
            });
            return resultList;
        }
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
     * @param cdsl
     * @param sjz
     * @param zx
     */
    private void writeExcelData(String proname, String htd, List<Map<String, Object>> lmzfList, List<Map<String, Object>> lmyfList, List<Map<String, Object>> sdzxList, List<Map<String, Object>> sdyxList, List<Map<String, Object>> qlzxList, List<Map<String, Object>> qlyxList, int cdsl, String sjz, String zx) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        String fname="";
        if (zx.equals("主线")){
            fname = "18路面平整度.xlsx";
        }else {
            fname = "61互通平整度-"+zx+".xlsx";
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
            filename = "平整度-5车道.xlsx";
        }else if (cdsl == 4){
            filename = "平整度-4车道.xlsx";
        }else if (cdsl == 3){
            filename = "平整度-3车道.xlsx";
        }else if (cdsl == 2){
            filename = "平整度-2车道.xlsx";
        }

        String path = reportPath + File.separator + filename;
        Files.copy(Paths.get(path), new FileOutputStream(f));
        FileInputStream out = new FileInputStream(f);
        wb = new XSSFWorkbook(out);


        List<Map<String,Object>> sdqlData = new ArrayList<>();


        sdzxList.addAll(sdyxList);
        if (sdzxList.size() >0 && !sdzxList.isEmpty()){
            //把有隧道的数据存入到sdqlData
            for (Map<String, Object> map : sdzxList) {
                Map<String,Object> map1 = new HashMap<>();
                map1.put("qdzh",map.get("qdzh").toString());
                map1.put("zdzh",map.get("zdzh").toString());
                map1.put("pzlx",map.get("pzlx").toString());
                map1.put("name",map.get("name").toString());
                map1.put("createTime",map.get("createTime").toString());
                map1.put("cd",map.get("cd").toString());
                if (map.get("zdbs")!=null){
                    map1.put("zdbs",map.get("zdbs").toString());
                }
                sdqlData.add(map1);
            }
            List<Map<String,Object>> addList = addMissingData(sdzxList,cdsl);
            List<Map<String, Object>> mapslist = mergedList(addList,cdsl);
            //sdqlData.addAll(mapslist);
            String sheetmame = "";
            if (mapslist.get(0).get("pzlx").toString().contains("水泥") && zx.equals("主线")){
                sheetmame = "混凝土隧道";
            }else if(mapslist.get(0).get("pzlx").toString().contains("水泥")){
                sheetmame = "混凝土匝道隧道";
            }else if (mapslist.get(0).get("pzlx").toString().contains("沥青") && zx.equals("主线")){
                sheetmame = "沥青隧道";
            }else if (mapslist.get(0).get("pzlx").toString().contains("沥青")){
                sheetmame = "沥青匝道隧道";
            }
            DBtoExcel(proname,htd,mapslist,wb,sheetmame,cdsl,sjz,zx);
        }

        qlzxList.addAll(qlyxList);
        if (qlzxList.size()>0 && !qlzxList.isEmpty()){
            //把有桥梁的数据存入到sdqlData
            for (Map<String, Object> map : qlzxList) {
                Map<String,Object> map1 = new HashMap<>();
                map1.put("qdzh",map.get("qdzh").toString());
                map1.put("zdzh",map.get("zdzh").toString());
                map1.put("pzlx",map.get("pzlx").toString());
                map1.put("name",map.get("name").toString());
                map1.put("createTime",map.get("createTime").toString());
                map1.put("cd",map.get("cd").toString());
                if (map.get("zdbs")!=null){
                    map1.put("zdbs",map.get("zdbs").toString());
                }
                sdqlData.add(map1);
            }
            List<Map<String,Object>> addList = addMissingData(qlzxList,cdsl);
            List<Map<String, Object>> mapslist = mergedList(addList,cdsl);
            //sdqlData.addAll(mapslist);
            String sheetmame = "";

            if (mapslist.get(0).get("pzlx").toString().contains("水泥") && zx.equals("主线")){
                sheetmame = "混凝土桥";
            }else if(mapslist.get(0).get("pzlx").toString().contains("水泥")){
                sheetmame = "混凝土匝道桥";
            }else if (mapslist.get(0).get("pzlx").toString().contains("沥青") && zx.equals("主线")){
                sheetmame = "沥青桥";
            }else if (mapslist.get(0).get("pzlx").toString().contains("沥青")){
                sheetmame = "沥青匝道桥";
            }
            DBtoExcel(proname,htd,mapslist,wb,sheetmame,cdsl,sjz,zx);
        }

        lmzfList.addAll(lmyfList);
        if (lmzfList.size()>0 && !lmzfList.isEmpty()){
            List<Map<String,Object>> addList = addMissingData(lmzfList,cdsl);
            String sheetmame = "";
            if (zx.equals("主线")){
                sheetmame = "沥青路面";
                DBtoExcelLM(proname,htd,addList,sdqlData,wb,sheetmame,cdsl,sjz,zx);
            }else {
                sheetmame = "沥青匝道";
                DBtoExcelZD(proname,htd,addList,sdqlData,wb,sheetmame,cdsl,sjz,zx);
            }
            /*if (addList.get(0).get("pzlx").toString().contains("水泥")){
                sheetmame = "混凝土路面";
            }
            if (addList.get(0).get("pzlx").toString().contains("沥青")){
                sheetmame = "沥青路面";
            }*/

        }

        //String[] arr = {"混凝土收费站","混凝土匝道隧道","沥青匝道隧道","混凝土匝道桥","沥青匝道桥","沥青匝道","混凝土匝道","混凝土隧道","沥青隧道","混凝土桥","沥青桥","混凝土路面","沥青路面"};
        String[] arr = {"混凝土隧道","沥青隧道","混凝土桥","沥青桥","混凝土路面","沥青路面"};
        for (int i = 0; i < arr.length; i++) {
            if (shouldBeCalculate(wb.getSheet(arr[i]))) {
                calculateAsphaltPavementSheet(wb,wb.getSheet(arr[i]),cdsl);
            }else {
                wb.removeSheetAt(wb.getSheetIndex(arr[i]));
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
     * @param data
     * @param sdqlData
     * @param wb
     * @param sheetname
     * @param cdsl
     * @param sjz
     * @param zx
     */
    private void DBtoExcelZD(String proname, String htd, List<Map<String, Object>> data,List<Map<String, Object>> sdqlData, XSSFWorkbook wb, String sheetname, int cdsl, String sjz, String zx) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        SimpleDateFormat outputDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        if (data != null && !data.isEmpty()) {
            createTable(getZDNum(data, cdsl), wb, sheetname, cdsl);
            XSSFSheet sheet = wb.getSheet(sheetname);

            String time = String.valueOf(data.get(0).get("createTime"));
            Date parse = simpleDateFormat.parse(time);
            String sj = outputDateFormat.format(parse);

            String fbgcname = "";
            if (sheetname.contains("隧道")) {
                fbgcname = "匝道隧道";
            } else if (sheetname.contains("桥")) {
                fbgcname = "匝道桥";
            } else {
                fbgcname = "匝道路面";
            }

            String name = data.get(0).get("zdbs").toString()+"匝道";
            sheet.getRow(1).getCell(5).setCellValue(proname);
            sheet.getRow(1).getCell(cdsl * 4 + 4).setCellValue(htd);
            sheet.getRow(2).getCell(5).setCellValue(zx);
            sheet.getRow(3).getCell(cdsl * 4 + 4).setCellValue(sj);
            sheet.getRow(2).getCell(cdsl * 4 + 4).setCellValue(fbgcname + "(" + name + ")");
            sheet.getRow(3).getCell(5).setCellValue(Double.parseDouble(sjz));

            Map<String, List<Map<String, Object>>> mapslist = handlezdData(proname,data,sdqlData,zx);
            List<Map<String, Object>> datamaplist = mapslist.get("data");
            List<Map<String, Object>> sfzmaplist = mapslist.get("sfzlist");

            datamaplist.sort((o1, o2) -> {
                String zdbs1 = o1.get("zdbs").toString();
                String zdbs2 = o2.get("zdbs").toString();
                int compareZdbs = zdbs1.compareTo(zdbs2);
                if (compareZdbs == 0) {
                    String qdzh1 = o1.get("qdzh").toString();
                    String qdzh2 = o2.get("qdzh").toString();
                    return qdzh1.compareTo(qdzh2);
                } else {
                    return compareZdbs;
                }
            });

            List<Map<String, Object>> zf = new ArrayList<>();
            List<Map<String, Object>> yf = new ArrayList<>();
            for (Map<String, Object> stringObjectMap : datamaplist) {
                if (stringObjectMap.get("cd").toString().equals("左幅")){
                    zf.add(stringObjectMap);
                }
                if (stringObjectMap.get("cd").toString().equals("右幅")){
                    yf.add(stringObjectMap);
                }
            }

            writeZDzyf(sheet,zf,proname, htd, name,sj,sheetname,cdsl,sjz);
            writeZDzyf(sheet,yf,proname, htd, name,sj,sheetname,cdsl,sjz);
            if (sfzmaplist.size()>0){
                List<Map<String, Object>> sfz = mergedList(sfzmaplist,cdsl);
                double zdzh = Double.parseDouble(sfz.get(0).get("qdzh").toString());
                double finzdzh = Double.parseDouble(sfz.get(sfz.size()-1).get("qdzh").toString());
                List<Map<String, Object>> sfzaddlist = decrementNumberByStep(zdzh, finzdzh, sfz, cdsl);
                sfzaddlist.sort((o1, o2) -> {
                    String zdbs1 = o1.get("zdbs").toString();
                    String zdbs2 = o2.get("zdbs").toString();
                    int compareZdbs = zdbs1.compareTo(zdbs2);
                    if (compareZdbs == 0) {
                        String qdzh1 = o1.get("qdzh").toString();
                        String qdzh2 = o2.get("qdzh").toString();
                        return qdzh1.compareTo(qdzh2);
                    } else {
                        return compareZdbs;
                    }
                });
                /*List<Map<String, Object>> zfsfz = new ArrayList<>();
                List<Map<String, Object>> yfsfz = new ArrayList<>();
                for (Map<String, Object> stringObjectMap : sfzmaplist) {
                    if (stringObjectMap.get("cd").toString().equals("左幅")){
                        zfsfz.add(stringObjectMap);
                    }
                    if (stringObjectMap.get("cd").toString().equals("右幅")){
                        yfsfz.add(stringObjectMap);
                    }
                }*/
                String sheesfzname = "混凝土收费站";
                createTable(getZDNum(sfzmaplist, cdsl), wb, sheetname, cdsl);
                XSSFSheet sheetsfz = wb.getSheet(sheesfzname);

                /*String sfztime = String.valueOf(sfzmaplist.get(0).get("createTime"));
                Date parse1 = simpleDateFormat.parse(time);
                String sfzsj = outputDateFormat.format(parse1);

                sheetsfz.getRow(1).getCell(5).setCellValue(proname);
                sheetsfz.getRow(1).getCell(cdsl * 4 + 4).setCellValue(htd);
                sheetsfz.getRow(2).getCell(5).setCellValue(zx);
                sheetsfz.getRow(3).getCell(cdsl * 4 + 4).setCellValue(sj);
                sheetsfz.getRow(2).getCell(cdsl * 4 + 4).setCellValue("收费站" + "(" + name + ")");
                sheetsfz.getRow(3).getCell(5).setCellValue(Double.parseDouble(sjz));*/
                int b = 0;
                if (cdsl == 2){
                    b=40;
                }else if (cdsl == 3){
                    b=20;
                }else if (cdsl == 4 || cdsl ==5){
                    b=30;
                }

                writesdqlzyf(wb,sfz,proname,htd,sheesfzname,cdsl,sjz,b,zx);

            }
        }


    }

    /**
     *
     * @param sheet
     * @param mapslist
     * @param proname
     * @param htd
     * @param name
     * @param sj
     * @param sheetname
     * @param cdsl
     * @param sjz
     */
    private void writeZDzyf(XSSFSheet sheet, List<Map<String, Object>> mapslist, String proname, String htd, String name, String sj, String sheetname, int cdsl, String sjz) {
        if (mapslist.size()>0) {
            List<Map<String, Object>> rowAndcol = new ArrayList<>();
            int startRow = -1, endRow = -1, startCol = -1, endCol = -1;
            int a = 0;
            int b = 0;
            if (cdsl == 2){
                a = 47;
                b=40;
            }else if (cdsl == 3){
                a = 27;
                b=20;
            }else if (cdsl == 4 || cdsl ==5){
                a = 37;
                b=30;
            }
            int index = 0;
            int tableNum = 0;
            String zdbs = mapslist.get(0).get("zdbs").toString();
            for (Map<String, Object> zd : mapslist) {
                if (zd.get("zdbs").toString().equals(zdbs)) {
                    if(index > (b-1)){
                        tableNum ++;
                        fillTitleCellData(sheet, tableNum, proname, htd, name,sj,sheetname,cdsl,sjz);
                        index = 0;
                    }
                    if (!zd.get("iri").toString().equals("") && !zd.get("iri").toString().isEmpty()) {
                        double n = Double.valueOf(zd.get("zdzh").toString()) / 1000;
                        int m = (int) n;
                        //fillCommonCellData(sheet, tableNum, index, lm,cdsl,"zx");
                        sheet.getRow(tableNum * a + 7 + index % b).getCell(0).setCellValue(m);
                        sheet.getRow(tableNum * a + 7 + index % b).getCell(4).setCellValue(Double.valueOf(zd.get("zdzh").toString()));
                        String[] sfc = zd.get("iri").toString().split(",");
                        for (int i = 0 ; i < sfc.length ; i++) {
                            if (zd.get("cd").equals("左幅")){
                                if (!sfc[i].equals("-")) {
                                    sheet.getRow(tableNum * a + 7 + index % b).getCell(5 + i).setCellValue(Double.parseDouble(sfc[i]));
                                }
                            }else {
                                if (!sfc[i].equals("-")){
                                    sheet.getRow(tableNum * a + 7 + index % b).getCell((2*cdsl+5)+i).setCellValue(Double.parseDouble(sfc[i]));
                                }

                            }
                        }

                    }else {
                        for (int i = 0; i < cdsl; i++) {
                            double n = Double.valueOf(zd.get("zdzh").toString()) / 1000;
                            int m = (int) n;
                            sheet.getRow(tableNum * a + 7 + index % b).getCell(0).setCellValue(m);
                            sheet.getRow(tableNum * a + 7 + index % b).getCell(4).setCellValue(Double.valueOf(zd.get("zdzh").toString()));
                            if (zd.get("cd").equals("左幅")){
                                sheet.getRow(tableNum * a + 7 + index % b).getCell(5 + i).setCellValue(zd.get("name").toString());

                                startRow = tableNum * a + 7 + index % b ;
                                endRow = tableNum * a + 7 + index % b ;

                                startCol = 5;
                                endCol = 2*cdsl+4;

                            }else {
                                sheet.getRow(tableNum * a + 7 + index % b).getCell((2*cdsl+5) + i).setCellValue(zd.get("name").toString());
                                startRow = tableNum * a + 7 + index % b ;
                                endRow = tableNum * a + 7 + index % b ;

                                startCol = 2*cdsl+5;
                                endCol = 4*cdsl+4;
                            }
                        }
                        //可以在这块记录一个行和列
                        Map<String, Object> map = new HashMap<>();
                        map.put("startRow",startRow);
                        map.put("endRow",endRow);
                        map.put("startCol",startCol);
                        map.put("endCol",endCol);
                        map.put("name",zd.get("name"));
                        map.put("tableNum",tableNum);
                        rowAndcol.add(map);
                    }
                    index++;
                }else {
                    zdbs = zd.get("zdbs").toString();
                    tableNum++;
                    index = 0;
                    fillTitleCellData(sheet, tableNum, proname, htd, zdbs + "匝道", sj, sheetname, cdsl, sjz);
                    if(index > (b-1)){
                        tableNum ++;
                        fillTitleCellData(sheet, tableNum, proname, htd, zdbs + "匝道",sj,sheetname,cdsl,sjz);
                        index = 0;
                    }
                    if (!zd.get("iri").toString().equals("") && !zd.get("iri").toString().isEmpty()) {
                        double n = Double.valueOf(zd.get("zdzh").toString()) / 1000;
                        int m = (int) n;
                        sheet.getRow(tableNum * a + 7 + index % b).getCell(0).setCellValue(m);
                        sheet.getRow(tableNum * a + 7 + index % b).getCell(4).setCellValue(Double.valueOf(zd.get("zdzh").toString()));
                        String[] sfc = zd.get("iri").toString().split(",");
                        for (int i = 0 ; i < sfc.length ; i++) {
                            if (zd.get("cd").equals("左幅")){
                                if (!sfc[i].equals("-")) {
                                    sheet.getRow(tableNum * a + 7 + index % b).getCell(5 + i).setCellValue(Double.parseDouble(sfc[i]));
                                }
                            }else {
                                if (!sfc[i].equals("-")){
                                    sheet.getRow(tableNum * a + 7 + index % b).getCell((2*cdsl+5)+i).setCellValue(Double.parseDouble(sfc[i]));
                                }

                            }
                        }

                    }else {
                        for (int i = 0; i < cdsl; i++) {
                            double n = Double.valueOf(zd.get("zdzh").toString()) / 1000;
                            int m = (int) n;
                            sheet.getRow(tableNum * a + 7 + index % b).getCell(0).setCellValue(m);
                            sheet.getRow(tableNum * a + 7 + index % b).getCell(4).setCellValue(Double.valueOf(zd.get("zdzh").toString()));
                            if (zd.get("cd").equals("左幅")){
                                sheet.getRow(tableNum * a + 7 + index % b).getCell(5 + i).setCellValue(zd.get("name").toString());

                                startRow = tableNum * a + 7 + index % b ;
                                endRow = tableNum * a + 7 + index % b ;

                                startCol = 5;
                                endCol = 2*cdsl+4;

                            }else {
                                sheet.getRow(tableNum * a + 7 + index % b).getCell((2*cdsl+5) + i).setCellValue(zd.get("name").toString());
                                startRow = tableNum * a + 7 + index % b ;
                                endRow = tableNum * a + 7 + index % b ;

                                startCol = 2*cdsl+5;
                                endCol = 4*cdsl+4;
                            }
                        }
                        //可以在这块记录一个行和列
                        Map<String, Object> map = new HashMap<>();
                        map.put("startRow",startRow);
                        map.put("endRow",endRow);
                        map.put("startCol",startCol);
                        map.put("endCol",endCol);
                        map.put("name",zd.get("name"));
                        map.put("tableNum",tableNum);
                        rowAndcol.add(map);
                    }
                    index++;
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
     * @param data
     * @param sdqlData
     * @param zx
     * @return
     */
    private Map<String, List<Map<String, Object>>>handlezdData(String proname, List<Map<String, Object>> data, List<Map<String, Object>> sdqlData, String zx) {
        /**
         * data是全部的数据，zfsdqlData是包含桥梁和隧道的数据，还有可能为空
         */

        Map<String, List<Map<String, Object>>> resultMap = new HashMap<>();
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
        List<JjgSfz> jjgSfzs = jjgLqsSfzMapper.selectList(wrapper);

        List<Map<String, Object>> sfzlist = new ArrayList<>();

        for (int i = 0; i < jjgSfzs.size(); i++) {
            double zhq = Double.parseDouble(jjgSfzs.get(i).getZhq());
            double zhz = Double.parseDouble(jjgSfzs.get(i).getZhz());
            String sszd = jjgSfzs.get(i).getSszd();
            String sfzlf = jjgSfzs.get(i).getLf();

            /**
             * 按data中的zdzh
             */
            for (Map<String, Object> datum : data) {
                double datazhq = Double.parseDouble(datum.get("zdzh").toString());
                String zdbs = datum.get("zdbs").toString();
                String cd = datum.get("cd").toString();
                if (datazhq > zhq && datazhq < zhz && sszd.equals(zdbs) && sfzlf.equals(cd)){
                    Map<String, Object> backupData = new HashMap<>();
                    backupData.put("zdzh",datum.get("zdzh"));
                    backupData.put("qdzh",datum.get("qdzh"));
                    backupData.put("cd",datum.get("cd"));
                    backupData.put("iri",datum.get("iri"));
                    backupData.put("pzlx",datum.get("pzlx"));
                    backupData.put("zdbs",datum.get("zdbs"));
                    backupData.put("createTime",datum.get("createTime"));
                    backupData.put("name",jjgSfzs.get(i).getZdsfzname());
                    sfzlist.add(backupData);
                    datum.put("iri","");
                    datum.put("name",jjgSfzs.get(i).getZdsfzname());

                }
            }
        }
        resultMap.put("sfzlist",sfzlist);

        if (sdqlData.size() > 0) {
            for (Map<String, Object> datum : data) {
                for (Map<String, Object> zfsdqlDatum : sdqlData) {
                    if (datum.get("zdbs").toString().equals(zfsdqlDatum.get("zdbs")) && datum.get("qdzh").toString().equals(zfsdqlDatum.get("qdzh")) && datum.get("cd").toString().equals(zfsdqlDatum.get("cd"))) {
                        datum.put("iri", "");
                        datum.put("name", zfsdqlDatum.get("name").toString());
                    }
                }
            }
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
                Double qdzh1 = Double.parseDouble(o1.get("qdzh").toString());
                Double qdzh2 = Double.parseDouble(o2.get("qdzh").toString());
                return qdzh1.compareTo(qdzh2);
            }
        });
        resultMap.put("data",data);

        return resultMap;

    }


    /**
     * 沥青路面
     * @param wb
     * @param sheet
     */
    private void calculateAsphaltPavementSheet(XSSFWorkbook wb, XSSFSheet sheet,int cdsl) {
        XSSFRow row = null;
        boolean flag = false;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        XSSFRow rowtotal = null;
        String name = "";
        FormulaEvaluator e = new XSSFFormulaEvaluator(wb);
        sheet.setColumnHidden(4,true);
        for (int i = 0; i <= sheet.getPhysicalNumberOfRows()-7; i++) {
            row = sheet.getRow(i);
            //如果当前行 row 为空行或第一列为空，则跳过本次循环，处理下一行数据
            if (row == null) {
                continue;
            }
            if (!"".equals(row.getCell(0).toString()) && row.getCell(0).toString().contains("质量鉴定表") && flag) {
                flag = false;
            }
            //如果当前行 row 的第一列不为空且与上一次读取到的不同，则说明当前行 row 是一个新的路段数据开始行，记录下路段数据的起始行 rowstart
            if (!"".equals(row.getCell(0).toString()) && !name.equals(row.getCell(0).toString()) && flag) {
                rowstart = sheet.getRow(i+1);
                name = rowstart.getCell(0).toString();
                rowend = sheet.getRow(getCellEndRow(sheet, rowstart.getRowNum(), 0));
                //rowstart.getRowNum() + 3作为起始行，总点数
                //获取连续的数据行 rowstart 到路段数据结束行 rowend。
                calculateTotalData(sheet, rowstart.getRowNum() + 3, rowstart, rowend, 2,e,cdsl);
            }
            if ("桩号".equals(row.getCell(0).toString())) {
                rowtotal = sheet.getRow(i + 1);
                i += 2;
                rowstart = sheet.getRow(i + 1);
                name = rowstart.getCell(0).toString();
                rowend = sheet.getRow(getCellEndRow(sheet, rowstart.getRowNum(), 0));
                calculateTotalData(sheet, rowstart.getRowNum() + 3, rowstart, rowend, 2,e,cdsl);
                flag = true;
            }
        }
        setExtraWholeTotalData(sheet, sheet.getRow(5),e,cdsl);

    }

    /**
     *
     * @param sheet
     * @param rowtotal
     * @param e
     */
    private void setExtraWholeTotalData(XSSFSheet sheet, XSSFRow rowtotal, FormulaEvaluator e,int cdsl) {
        /**
         * 2c 19  cdsl*4+11
         * 3c 21
         * 4c 25
         * 5c 29
         */
        rowtotal.createCell(cdsl*4+9).setCellFormula(
                "SUM("
                        + sheet.getRow(rowtotal.getRowNum() + 1).createCell(cdsl*4+9)
                        .getReference()
                        + ":"
                        + sheet.getRow(sheet.getPhysicalNumberOfRows() - 1)
                        .createCell(cdsl*4+9).getReference() + ")");
        double value = e.evaluate(rowtotal.getCell(cdsl*4+9)).getNumberValue();
        rowtotal.getCell(cdsl*4+9).setCellFormula(null);
        rowtotal.getCell(cdsl*4+9).setCellValue(value);

        rowtotal.createCell(cdsl*4+10).setCellFormula(
                "SUM("
                        + sheet.getRow(rowtotal.getRowNum() + 1).createCell(cdsl*4+10)
                        .getReference()
                        + ":"
                        + sheet.getRow(sheet.getPhysicalNumberOfRows() - 1)
                        .createCell(cdsl*4+10).getReference() + ")");
        value = e.evaluate(rowtotal.getCell(cdsl*4+10)).getNumberValue();
        rowtotal.getCell(cdsl*4+10).setCellFormula(null);
        rowtotal.getCell(cdsl*4+10).setCellValue(value);

        rowtotal.createCell(cdsl*4+11).setCellFormula(
                rowtotal.getCell(cdsl*4+10).getReference() + "/"
                        + rowtotal.getCell(cdsl*4+9).getReference() + "*100");
        value = e.evaluate(rowtotal.getCell(cdsl*4+11)).getNumberValue();
        rowtotal.getCell(cdsl*4+11).setCellFormula(null);
        rowtotal.getCell(cdsl*4+11).setCellValue(value);

        rowtotal.createCell(cdsl*4+12).setCellFormula(
                "MIN("
                        + sheet.getRow(rowtotal.getRowNum() + 1).createCell(cdsl*4+12)
                        .getReference()
                        + ":"
                        + sheet.getRow(sheet.getPhysicalNumberOfRows() - 1)
                        .createCell(cdsl*4+12).getReference() + ")");
        value = e.evaluate(rowtotal.getCell(cdsl*4+12)).getNumberValue();
        rowtotal.getCell(cdsl*4+12).setCellFormula(null);
        rowtotal.getCell(cdsl*4+12).setCellValue(value);

        rowtotal.createCell(cdsl*4+11).setCellFormula(
                "MAX("
                        + sheet.getRow(rowtotal.getRowNum() + 1).createCell(cdsl*4+11)
                        .getReference()
                        + ":"
                        + sheet.getRow(sheet.getPhysicalNumberOfRows() - 1)
                        .createCell(cdsl*4+11).getReference() + ")");
        value = e.evaluate(rowtotal.getCell(cdsl*4+11)).getNumberValue();
        rowtotal.getCell(cdsl*4+11).setCellFormula(null);
        rowtotal.getCell(cdsl*4+11).setCellValue(value);

    }

    /**
     * 计算每一个桩号内的总结数据
     */
    private void calculateTotalData(XSSFSheet sheet, int rownum, XSSFRow rowstart, XSSFRow rowend, int num, FormulaEvaluator e,int cdsl) {

        if (num == 1) {
            if (rowend.getRowNum() <= getCellEndRow(sheet, rowstart.getRowNum(), 5)) {
                fillTotalData(sheet, rownum, rowstart, rowend, 0, cdsl*4+7, false,e,cdsl);
            } else {
                if (rowstart.getCell(5).getCellType() == Cell.CELL_TYPE_FORMULA || rowstart.getCell(5).getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    fillTotalData(sheet, rownum, rowstart, rowend, 0, cdsl*4+7, true,e,cdsl);
                } else {
                    int record = getCellEndRow(sheet, rowstart.getRowNum(), 5);
                    boolean flag = true;
                    while (rowend.getRowNum() > record) {
                        if (sheet.getRow(record + 1).getCell(5).getCellType() == Cell.CELL_TYPE_FORMULA || sheet.getRow(record + 1).getCell(5).getCellType() == Cell.CELL_TYPE_NUMERIC) {
                            fillTotalData(sheet, rownum, rowstart, rowend, 0, cdsl*4+7, true,e,cdsl);
                            flag = false;
                            break;
                        }
                        record = getCellEndRow(sheet, record + 1, 5);
                    }
                    if (flag) {
                        fillTotalData(sheet, rownum, rowstart, rowend, 0, cdsl*4+7, false,e,cdsl);
                    }
                }
            }
        } else if (num == 2) {
            if (rowend.getRowNum() <= getCellEndRow(sheet, rowstart.getRowNum(), 5)) {
                fillTotalData(sheet, rownum, rowstart, rowend, 0, cdsl*4+7, true,e,cdsl);
            } else {
                if (rowstart.getCell(5).getCellType() == Cell.CELL_TYPE_FORMULA || rowstart.getCell(5).getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    fillTotalData(sheet, rownum, rowstart, rowend, 0, cdsl*4+7, true,e,cdsl);
                } else {
                    int record = getCellEndRow(sheet, rowstart.getRowNum(), 5);
                    boolean flag = true;
                    while (rowend.getRowNum() > record) {
                        if (sheet.getRow(record + 1).getCell(5).getCellType() == Cell.CELL_TYPE_FORMULA || sheet.getRow(record + 1).getCell(5).getCellType() == Cell.CELL_TYPE_NUMERIC) {
                            fillTotalData(sheet, rownum, rowstart, rowend, 0, cdsl*4+7, true,e,cdsl);
                            flag = false;
                            break;
                        }
                        record = getCellEndRow(sheet, record + 1, 5);
                    }
                    if (flag) {
                        fillTotalData(sheet, rownum, rowstart, rowend, 0, cdsl*4+7, true,e,cdsl);
                    }
                }
            }
            if (rowend.getRowNum() <= getCellEndRow(sheet, rowstart.getRowNum(), 9)) {
                fillTotalData(sheet, rownum, rowstart, rowend, 1, cdsl*4+8, true,e,cdsl);
            } else {
                if (rowstart.getCell(9).getCellType() == Cell.CELL_TYPE_FORMULA || rowstart.getCell(9).getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    fillTotalData(sheet, rownum, rowstart, rowend, 1, cdsl*4+8, true,e,cdsl);
                } else {
                    int record = getCellEndRow(sheet, rowstart.getRowNum(), 9);
                    boolean flag = true;
                    while (rowend.getRowNum() > record) {
                        if (sheet.getRow(record + 1).getCell(9).getCellType() == Cell.CELL_TYPE_FORMULA || sheet.getRow(record + 1).getCell(9).getCellType() == Cell.CELL_TYPE_NUMERIC) {
                            fillTotalData(sheet, rownum, rowstart, rowend, 1, cdsl*4+8, true,e,cdsl);
                            flag = false;
                            break;
                        }
                        record = getCellEndRow(sheet, record + 1, 9);
                    }
                    if (flag) {
                        fillTotalData(sheet, rownum, rowstart, rowend, 1, cdsl*4+8, true,e,cdsl);
                    }
                }
            }
        }
    }

    /**
     * 将结果数据填到表格的右部
     * @param sheet
     * @param rownum
     * @param rowstart
     * @param rowend
     * @param left_right
     * @param fillcolum
     * @param hasdata
     */
    private void fillTotalData(XSSFSheet sheet, int rownum, XSSFRow rowstart, XSSFRow rowend, int left_right, int fillcolum, boolean hasdata,FormulaEvaluator e,int cdsl) {
        if (hasdata) {
            sheet.getRow(rownum)
                    .getCell(fillcolum)
                    .setCellFormula("IFERROR("+
                            "COUNT("
                            + rowstart.getCell(5 + left_right * 4)
                            .getReference()
                            + ":"
                            + rowend.getCell(8 + left_right * 4)
                            .getReference() + ")"+",\"-\")");// 总点数P=COUNT(F8:I17)

            sheet.getRow(rownum + 1)
                    .getCell(fillcolum)
                    .setCellFormula("IFERROR("+
                            sheet.getRow(rownum).getCell(fillcolum)
                                    .getReference()
                            + "-COUNTIF("
                            + rowstart.getCell(5 + left_right * 4)
                            .getReference()
                            + ":"
                            + rowend.getCell(8 + left_right * 4)
                            .getReference() + ",\">\"&$F$4)"+",\"-\")");// 合格数P=P11-COUNTIF(F8:I17,">"&$F$4)

            sheet.getRow(rownum + 2)
                    .getCell(fillcolum)
                    .setCellFormula("IFERROR("+
                            "AVERAGE("
                            + rowstart.getCell(5 + left_right * 4)
                            .getReference()
                            + ":"
                            + rowend.getCell(8 + left_right * 4)
                            .getReference() + ")"+",\"-\")");// 平均值P=AVERAGE(F8:I17)
            sheet.getRow(rownum + 3)
                    .getCell(fillcolum)
                    .setCellFormula("IFERROR("+
                            "STDEV("
                            + rowstart.getCell(5 + left_right * 4)
                            .getReference()
                            + ":"
                            + rowend.getCell(8 + left_right * 4)
                            .getReference() + ")"+",\"-\")");// 均方差P=STDEV(F8:I17)

            sheet.getRow(rownum + 4)
                    .getCell(fillcolum)
                    .setCellFormula("IF("+sheet.getRow(rownum)
                            .getCell(fillcolum).getReference()+"=0,\"-\","+
                            "MIN("
                            + rowstart.getCell(5 + left_right * 4)
                            .getReference()
                            + ":"
                            + rowend.getCell(8 + left_right * 4)
                            .getReference() + ")"+")");// 最小值P=MIN(F8:I17)=IF(MIN(P92,Q92)=0,"",MIN(P92,Q92))

            sheet.getRow(rownum + 5)
                    .getCell(fillcolum)
                    .setCellFormula("IF("+sheet.getRow(rownum)
                            .getCell(fillcolum).getReference()+"=0,\"-\","+
                            "MAX("
                            + rowstart.getCell(5 + left_right * 4)
                            .getReference()
                            + ":"
                            + rowend.getCell(8 + left_right * 4)
                            .getReference() + ")"+")");// 最大值P=MAX(F8:I17)

            sheet.getRow(rownum + 6)
                    .getCell(fillcolum)
                    .setCellFormula("IFERROR("+
                            sheet.getRow(rownum + 1).getCell(fillcolum)
                                    .getReference()
                            + "*100/"
                            + sheet.getRow(rownum)
                            .getCell(fillcolum).getReference()+",\"-\")");// 合格率P=P12*100/P11
            if(left_right == 1){
                fillExtraTotalData(sheet, rownum - 1, hasdata,e,cdsl);
            }

        } else {
            sheet.getRow(rownum).getCell(fillcolum).setCellValue("-");
            sheet.getRow(rownum + 1).getCell(fillcolum).setCellValue("-");
            sheet.getRow(rownum + 2).getCell(fillcolum).setCellValue("-");
            sheet.getRow(rownum + 3).getCell(fillcolum).setCellValue("-");
            sheet.getRow(rownum + 4).getCell(fillcolum).setCellValue("-");
            sheet.getRow(rownum + 5).getCell(fillcolum).setCellValue("-");
            sheet.getRow(rownum + 6).getCell(fillcolum).setCellValue("-");
            if(left_right == 1){
                fillExtraTotalData(sheet, rownum - 1, true,e,cdsl);
            }
        }

    }

    /**
     * 在表格右边外面的位置，填写统计数据
     * @param sheet
     * @param rownum
     * @param hasdata
     * @param e
     */
    private void fillExtraTotalData(XSSFSheet sheet, int rownum, boolean hasdata, FormulaEvaluator e,int cdsl) {
        if (hasdata) {
            sheet.getRow(rownum)
                    .createCell(cdsl*4+9)
                    .setCellFormula(
                            "IF("
                                    + sheet.getRow(rownum + 1).getCell(cdsl*4+7)
                                    .getReference()
                                    + "=\"-\",0,"
                                    + sheet.getRow(rownum + 1).getCell(cdsl*4+7)
                                    .getReference()
                                    + ")+IF("
                                    + sheet.getRow(rownum + 1).getCell(cdsl*4+8)
                                    .getReference()
                                    + "=\"-\",0,"
                                    + sheet.getRow(rownum + 1).getCell(cdsl*4+8)
                                    .getReference() + ")");
            double value = e.evaluate(sheet.getRow(rownum).getCell(cdsl*4+9)).getNumberValue();
            sheet.getRow(rownum).getCell(cdsl*4+9).setCellFormula(null);
            sheet.getRow(rownum).getCell(cdsl*4+9).setCellValue(value);
            // "IF("+sheet.getRow(rownum+1).getCell(15).getReference()+"=\"-\""+"+"
            // +sheet.getRow(rownum+1).getCell(16).getReference());//总点数R=P31+Q31=IF(P323="-",0,P323)+IF(Q323="-",0,Q323)

            sheet.getRow(rownum)
                    .createCell(cdsl*4+10)
                    .setCellFormula(
                            "IF("
                                    + sheet.getRow(rownum + 2).getCell(cdsl*4+7)
                                    .getReference()
                                    + "=\"-\",0,"
                                    + sheet.getRow(rownum + 2).getCell(cdsl*4+7)
                                    .getReference()
                                    + ")+IF("
                                    + sheet.getRow(rownum + 2).getCell(cdsl*4+8)
                                    .getReference()
                                    + "=\"-\",0,"
                                    + sheet.getRow(rownum + 2).getCell(cdsl*4+8)
                                    .getReference() + ")");
            value = e.evaluate(sheet.getRow(rownum).getCell(cdsl*4+10)).getNumberValue();
            sheet.getRow(rownum).getCell(cdsl*4+10).setCellFormula(null);
            sheet.getRow(rownum).getCell(cdsl*4+10).setCellValue(value);

            sheet.getRow(rownum)
                    .createCell(cdsl*4+11)
                    .setCellFormula(
                            "IF("+ sheet.getRow(rownum).getCell(cdsl*4+9).getReference()+"=0,0,"+
                            sheet.getRow(rownum).getCell(cdsl*4+10).getReference()
                                    + "/"
                                    + sheet.getRow(rownum).getCell(cdsl*4+9)
                                    .getReference() + "*100)");// 合格率T=S30/R30*100

            sheet.getRow(rownum)
                    .createCell(cdsl*4+12)
                    .setCellFormula(
                            "IF(MIN("
                                    + sheet.getRow(rownum + 5).getCell(cdsl*4+7)
                                    .getReference()
                                    + ","
                                    + sheet.getRow(rownum + 5).getCell(cdsl*4+8)
                                    .getReference()
                                    + ")=0,\"\",MIN("
                                    + sheet.getRow(rownum + 5).getCell(cdsl*4+7)
                                    .getReference()
                                    + ","
                                    + sheet.getRow(rownum + 5).getCell(cdsl*4+8)
                                    .getReference() + "))");
            value = e.evaluate(sheet.getRow(rownum).getCell(cdsl*4+12)).getNumberValue();
            sheet.getRow(rownum).getCell(cdsl*4+12).setCellFormula(null);
            if(value < 0.0001){
                sheet.getRow(rownum).getCell(cdsl*4+12).setCellValue("");
            }
            else{
                sheet.getRow(rownum).getCell(cdsl*4+12).setCellValue(value);
            }

            sheet.getRow(rownum)
                    .createCell(cdsl*4+11)
                    .setCellFormula(
                            "IF(MAX("
                                    + sheet.getRow(rownum + 6).getCell(cdsl*4+7)
                                    .getReference()
                                    + ","
                                    + sheet.getRow(rownum + 6).getCell(cdsl*4+8)
                                    .getReference()
                                    + ")=0,\"\",MAX("
                                    + sheet.getRow(rownum + 6).getCell(cdsl*4+7)
                                    .getReference()
                                    + ","
                                    + sheet.getRow(rownum + 6).getCell(cdsl*4+8)
                                    .getReference() + "))");
            value = e.evaluate(sheet.getRow(rownum).getCell(cdsl*4+11)).getNumberValue();
            sheet.getRow(rownum).getCell(cdsl*4+11).setCellFormula(null);
            if(value < 0.0001){
                sheet.getRow(rownum).getCell(cdsl*4+11).setCellValue("");
            }
            else{
                sheet.getRow(rownum).getCell(cdsl*4+11).setCellValue(value);
            }
        } else {
            sheet.getRow(rownum).createCell(cdsl*4+9).setCellValue("-");
            sheet.getRow(rownum).createCell(cdsl*4+10).setCellValue("-");
            sheet.getRow(rownum).createCell(cdsl*4+11).setCellValue("-");
            sheet.getRow(rownum).createCell(cdsl*4+12).setCellValue("-");
            sheet.getRow(rownum).createCell(cdsl*4+11).setCellValue("-");
        }

    }

    /**
     * 根据给定的单元格起始行号，得到合并单元格的最后一行行号 如果给定的初始行号不是合并单元格，那么函数返回初始行号
     * @param sheet
     * @param cellstartrow
     * @param cellstartcol
     * @return
     */
    private int getCellEndRow(XSSFSheet sheet, int cellstartrow, int cellstartcol) {
        int sheetmergerCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetmergerCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            if (cellstartrow >= ca.getFirstRow() && cellstartrow <= ca.getLastRow() && cellstartcol >= ca.getFirstColumn() && cellstartcol <= ca.getLastColumn()) {
                return ca.getLastRow();
            }
        }
        return cellstartrow;


    }


    /**
     *
     * @param sheet
     * @return
     */
    private boolean shouldBeCalculate(XSSFSheet sheet) {
        sheet.getRow(7).getCell(0).setCellType(CellType.STRING);
        if(sheet.getRow(7).getCell(0)==null ||"".equals(sheet.getRow(7).getCell(0).getStringCellValue())){
            return false;
        }

        return true;
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
    private void DBtoExcelLM(String proname, String htd, List<Map<String, Object>> data,List<Map<String, Object>> sdqlData, XSSFWorkbook wb, String sheetname, int cdsl, String sjz, String zx) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        SimpleDateFormat outputDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        if (data != null && !data.isEmpty()) {
            createTable(getNum(data,cdsl)/2, wb, sheetname, cdsl);
            XSSFSheet sheet = wb.getSheet(sheetname);

            String time = String.valueOf(data.get(0).get("createTime")) ;
            Date parse = simpleDateFormat.parse(time);
            String sj = outputDateFormat.format(parse);

            String fbgcname = "";
            if (sheetname.contains("隧道")){
                fbgcname = "隧道路面";
            }else if (sheetname.contains("桥")){
                fbgcname = "桥面系";
            }else {
                fbgcname = "路面面层";
            }

            String name = data.get(0).get("name").toString();
            sheet.getRow(1).getCell(5).setCellValue(proname);
            sheet.getRow(1).getCell(cdsl*4+4).setCellValue(htd);
            sheet.getRow(2).getCell(5).setCellValue("路面工程");
            sheet.getRow(3).getCell(cdsl*4+4).setCellValue(sj);
            sheet.getRow(2).getCell(cdsl*4+4).setCellValue(fbgcname+"("+zx+")");
            sheet.getRow(3).getCell(5).setCellValue(Double.parseDouble(sjz));

            List<Map<String, Object>> lmdata = handleLmData(data,sdqlData);

            List<Map<String, Object>> zf = new ArrayList<>();
            List<Map<String, Object>> yf = new ArrayList<>();
            for (Map<String, Object> stringObjectMap : lmdata) {
                if (stringObjectMap.get("cd").toString().equals("左幅")){
                    zf.add(stringObjectMap);
                }
                if (stringObjectMap.get("cd").toString().equals("右幅")){
                    yf.add(stringObjectMap);
                }
            }
            writezyf(sheet,zf,proname, htd, name,sj,sheetname,cdsl,sjz);
            writezyf(sheet,yf,proname, htd, name,sj,sheetname,cdsl,sjz);

        }
    }

    /**
     *
     * @param sheet
     * @param lmdata
     * @param proname
     * @param htd
     * @param name
     * @param sj
     * @param sheetname
     * @param cdsl
     * @param sjz
     */
    private void writezyf(XSSFSheet sheet, List<Map<String, Object>> lmdata, String proname, String htd, String name, String sj, String sheetname, int cdsl, String sjz) {
        int index = 0;
        int tableNum = 0;
        List<Map<String, Object>> rowAndcol = new ArrayList<>();
        int startRow = -1, endRow = -1, startCol = -1, endCol = -1;
        int a = 0;
        int b = 0;
        if (cdsl == 2){
            a = 47;
            b=40;
        }else if (cdsl == 3){
            a = 27;
            b=20;
        }else if (cdsl == 4 || cdsl ==5){
            a = 37;
            b=30;
        }
        for (Map<String, Object> lm : lmdata) {
            if(index > (b-1)){
                tableNum ++;
                fillTitleCellData(sheet, tableNum, proname, htd, name,sj,sheetname,cdsl,sjz);
                index = 0;
            }
            if (!lm.get("iri").toString().equals("") && !lm.get("iri").toString().isEmpty()) {
                double n = Double.valueOf(lm.get("zdzh").toString()) / 1000;
                int m = (int) n;
                //fillCommonCellData(sheet, tableNum, index, lm,cdsl,"zx");
                sheet.getRow(tableNum * a + 7 + index % b).getCell(0).setCellValue(m);
                sheet.getRow(tableNum * a + 7 + index % b).getCell(4).setCellValue(Double.valueOf(lm.get("zdzh").toString()));
                String[] sfc = lm.get("iri").toString().split(",");
                for (int i = 0 ; i < sfc.length ; i++) {
                    if (lm.get("cd").equals("左幅")){
                        if (!sfc[i].equals("-")) {
                            sheet.getRow(tableNum * a + 7 + index % b).getCell(5 + i).setCellValue(Double.parseDouble(sfc[i]));
                        }else {
                            sheet.getRow(tableNum * a + 7 + index % b).getCell(5 + i).setCellValue(sfc[i]);
                        }
                    }else {
                        if (!sfc[i].equals("-")){
                            sheet.getRow(tableNum * a + 7 + index % b).getCell((2*cdsl+5)+i).setCellValue(Double.parseDouble(sfc[i]));
                        }else {
                            sheet.getRow(tableNum * a + 7 + index % b).getCell((2*cdsl+5)+i).setCellValue(sfc[i]);
                        }

                    }
                }

            }else {
                for (int i = 0; i < cdsl; i++) {
                    double n = Double.valueOf(lm.get("zdzh").toString()) / 1000;
                    int m = (int) n;
                    sheet.getRow(tableNum * a + 7 + index % b).getCell(0).setCellValue(m);
                    sheet.getRow(tableNum * a + 7 + index % b).getCell(4).setCellValue(Double.valueOf(lm.get("zdzh").toString()));
                    if (lm.get("cd").equals("左幅")){
                        sheet.getRow(tableNum * a + 7 + index % b).getCell(5 + i).setCellValue(lm.get("name").toString());

                        startRow = tableNum * a + 7 + index % b ;
                        endRow = tableNum * a + 7 + index % b ;

                        startCol = 5;
                        endCol = 2*cdsl+4;

                    }else {
                        sheet.getRow(tableNum * a + 7 + index % b).getCell((2*cdsl+5) + i).setCellValue(lm.get("name").toString());
                        startRow = tableNum * a + 7 + index % b ;
                        endRow = tableNum * a + 7 + index % b ;

                        startCol = 2*cdsl+5;
                        endCol = 4*cdsl+4;
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
     * @param sdqlData
     * @return
     */
    private List<Map<String, Object>> handleLmData(List<Map<String, Object>> data, List<Map<String, Object>> sdqlData) {

        for (Map<String, Object> datum : data) {
            for (Map<String, Object> zfsdqlDatum : sdqlData) {
                if (datum.get("qdzh").toString().equals(zfsdqlDatum.get("qdzh")) && datum.get("cd").toString().equals(zfsdqlDatum.get("cd"))){
                    datum.put("iri","");
                    datum.put("name",zfsdqlDatum.get("name"));
                }
            }
        }
        Collections.sort(data, new Comparator<Map<String, Object>>() {
            @Override
            public int compare(Map<String, Object> o1, Map<String, Object> o2) {
                // 名字相同时按照 qdzh 排序
                Double qdzh1 = Double.parseDouble(o1.get("qdzh").toString());
                Double qdzh2 = Double.parseDouble(o2.get("qdzh").toString());
                return qdzh1.compareTo(qdzh2);
            }
        });
        return data;
    }

    /**
     *
     * @param dataList
     * @return
     */
    private List<Map<String, Object>> addMissingData(List<Map<String, Object>> dataList,int cdsl) {

        List<Map<String, Object>> zf = new ArrayList<>();
        List<Map<String, Object>> yf = new ArrayList<>();
        for (Map<String, Object> stringObjectMap : dataList) {
            if (stringObjectMap.get("cd").toString().contains("左幅")){
                zf.add(stringObjectMap);
            }
            if (stringObjectMap.get("cd").toString().contains("右幅")){
                yf.add(stringObjectMap);
            }
        }

        List<Map<String, Object>> groupaddz = groupadd(zf,cdsl);
        List<Map<String, Object>> groupaddy = groupadd(yf,cdsl);

        groupaddz.addAll(groupaddy);
        return groupaddz;

    }

    /**
     *
     * @param dataList
     * @param cdsl
     * @return
     */
    private List<Map<String, Object>> groupadd(List<Map<String, Object>> dataList,int cdsl) {
        // 按照 name 进行分组
        Map<String, List<Map<String, Object>>> groupMap = new HashMap<>();
        for (Map<String, Object> data : dataList) {
            String name = (String) data.get("name");
            if (groupMap.containsKey(name)) {
                groupMap.get(name).add(data);
            } else {
                List<Map<String, Object>> list = new ArrayList<>();
                list.add(data);
                groupMap.put(name, list);
            }
        }
        List<Map<String, Object>> result = new ArrayList<>();
        // 遍历 groupMap 中的所有键值对
        for (Map.Entry<String, List<Map<String, Object>>> entry : groupMap.entrySet()) {
            String key = entry.getKey();
            List<Map<String, Object>> valueList = entry.getValue();
            int size = valueList.size();
            double zdzh = Double.parseDouble(valueList.get(0).get("qdzh").toString());
            double finzdzh = Double.parseDouble(valueList.get(size-1).get("qdzh").toString());
            List<Map<String, Object>> maps = decrementNumberByStep(zdzh,finzdzh,valueList,cdsl);
            result.addAll(maps);
        }
        return result;

    }

    /**
     * 添加桩号
     * @param zdzh
     * @param finzdzh
     * @param mapList
     * @return
     */
    private List<Map<String, Object>> decrementNumberByStep(double zdzh,double finzdzh, List<Map<String, Object>> mapList,int cds) {
        int cdsl = cds * 2;
        String iri = "-";
        StringBuilder iriBuilder = new StringBuilder();
        for (int i = 0; i < cdsl; i++) {
            if (i == cdsl - 1) {
                iriBuilder.append("-");
            } else {
                iriBuilder.append("-,");
            }
        }
        iri = iriBuilder.toString();

        List<Map<String, Object>> zhlist = new ArrayList<>();
        List<Map<String, Object>> finzdzhlist = new ArrayList<>();
        int c = (int)zdzh;
        int a = c/1000;
        int b = a*1000;//49000

        for (int i = c; i>b;i-=100){
            Map map = new HashMap();
            map.put("cd",mapList.get(0).get("cd"));
            map.put("createTime",mapList.get(0).get("createTime"));
            map.put("zdbs",mapList.get(0).get("zdbs"));
            map.put("iri",iri);
            map.put("name",mapList.get(0).get("name"));
            map.put("pzlx",mapList.get(0).get("pzlx"));
            map.put("qdzh",Double.parseDouble(String.valueOf(i-100)));
            map.put("zdzh",Double.parseDouble(String.valueOf(i)));
            zhlist.add(map);
        }

        int m = (int)finzdzh;
        int n = m/1000+1;
        int v = n*1000;

        for (int i = m+100; i<v;i+=100){
            Map map = new HashMap();
            map.put("cd",mapList.get(0).get("cd"));
            map.put("createTime",mapList.get(0).get("createTime"));
            map.put("zdbs",mapList.get(0).get("zdbs"));
            map.put("iri",iri);
            map.put("name",mapList.get(0).get("name"));
            map.put("pzlx",mapList.get(0).get("pzlx"));
            map.put("zdzh",Double.parseDouble(String.valueOf(i+100)));
            map.put("qdzh",Double.parseDouble(String.valueOf(i)));
            finzdzhlist.add(map);
        }

        mapList.addAll(zhlist);
        mapList.addAll(finzdzhlist);

        Collections.sort(mapList, new Comparator<Map<String, Object>>() {
            @Override
            public int compare(Map<String, Object> o1, Map<String, Object> o2) {
                String name1 = o1.get("zdzh").toString();
                String name2 = o2.get("zdzh").toString();
                return name1.compareTo(name2);
            }
        });
        return mapList;

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
        if (cdsl == 2){
            b=40;
        }else if (cdsl == 3){
            b=20;
        }else if (cdsl == 4 || cdsl ==5){
            b=30;
        }
        if (data!=null && !data.isEmpty()){
            //List<Map<String, Object>> maps = mergedList(data,cdsl);
            createTable(getNum(data,cdsl),wb,sheetname,cdsl);
            writesdqlzyf(wb,data,proname, htd,sheetname,cdsl,sjz,b,zx);

        }

    }

    /**
     *
     * @param dataLists
     * @param cds
     * @return
     */
    private List<Map<String, Object>> mergedList(List<Map<String, Object>> dataLists,int cds) {
        //处理拼接的iri
        int cdsl = cds * 2;
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
                                item -> item.get("name") + "-" + item.get("qdzh") + "-" + item.get("zdzh")
                        ));
        List<Map<String, Object>> toBeRemoved = new ArrayList<>();
        for (Map.Entry<String, List<Map<String, Object>>> entry : groupedData.entrySet()) {
            String groupName = entry.getKey(); // 获取分组名字
            List<Map<String, Object>> groupData = entry.getValue(); // 获取分组数据

            if (groupData.size() > 1) {
                Map<String, Object> firstItem = groupData.get(0);
                Map<String, Object> secondItem = groupData.get(1);
                if (firstItem.get("cd").equals("左幅")) {
                    // 将第二条数据的iri拼接在第一条数据的iri后面
                    String iri = firstItem.get("iri").toString() + "," + secondItem.get("iri").toString();
                    firstItem.put("iri", iri);
                } else {
                    // 将第二条数据的iri拼接在第一条数据的iri前面
                    String iri = secondItem.get("iri").toString() + "," + firstItem.get("iri").toString();
                    firstItem.put("iri", iri);
                }
                // 将secondItem条数据删除
                toBeRemoved.add(secondItem);

            } else if (groupData.size() == 1) {
                Map<String, Object> item = groupData.get(0);
                if (item.get("cd").equals("左幅")) {
                    // 在iri后面拼接逗号
                    String iri = item.get("iri").toString() +","+ iris;
                    item.put("iri", iri);
                } else {
                    // 在iri前面拼接逗号
                    String iri = iris +","+ item.get("iri").toString();
                    item.put("iri", iri);
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
                Double qdzh1 = Double.parseDouble(o1.get("qdzh").toString());
                Double qdzh2 = Double.parseDouble(o2.get("qdzh").toString());
                return qdzh1.compareTo(qdzh2);
            }
        });


        return dataLists;
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
        System.out.println(data);
        XSSFSheet sheet = wb.getSheet(sheetname);
        String time = String.valueOf(data.get(0).get("createTime")) ;
        Date parse = simpleDateFormat.parse(time);
        String sj = outputDateFormat.format(parse);

        sheet.getRow(1).getCell(5).setCellValue(proname);
        sheet.getRow(2).getCell(5).setCellValue(zx);
        sheet.getRow(3).getCell(5).setCellValue(Double.parseDouble(sjz));

        sheet.getRow(3).getCell(cdsl*4+4).setCellValue(sj);
        sheet.getRow(1).getCell(cdsl*4+4).setCellValue(htd);
        String name = data.get(0).get("name").toString();
        String fbgcname = "";
        if (sheetname.contains("隧道")){
            fbgcname = "隧道路面";
        }else if (sheetname.contains("桥")){
            fbgcname = "桥面系";
        }else if(sheetname.contains("收费站")){
            fbgcname = "收费站";
        } else {
            fbgcname = "路面面层";
        }
        sheet.getRow(2).getCell(cdsl*4+4).setCellValue(fbgcname+"("+name+")");
        int index = 0;
        int tableNum = 0;

        for(int i =0; i < data.size(); i++){
            if (name.equals(data.get(i).get("name"))){
                if(index > b){
                    tableNum ++;
                    fillTitleCellData(sheet, tableNum, proname, htd, name,sj,sheetname,cdsl,sjz);
                    index = 0;
                }
                fillCommonCellData(sheet, tableNum, index, data.get(i),cdsl,zx);
                index ++;
            }else {
                name = data.get(i).get("name").toString();
                tableNum ++;
                index = 0;
                fillTitleCellData(sheet, tableNum, proname, htd, name,sj,sheetname,cdsl,sjz);
                fillCommonCellData(sheet, tableNum, index, data.get(i),cdsl,zx);
                index += 1;
            }
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
            a = 40;

        }else if (cdsl == 3){
            a = 20;

        }else if (cdsl == 4 || cdsl == 5){
            a = 30;
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
    private int getZDNum(List<Map<String, Object>> data,int cdsl) {
        int a = 0;
        if (cdsl == 2){
            a = 40;

        }else if (cdsl == 3){
            a = 20;

        }else if (cdsl == 4 || cdsl == 5){
            a = 30;
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
            if (value%a==0){
                num += value/a;
            }else {
                num += value/a+1;
            }
        }
        return num;
    }

    /**
     * 创建页
     * @param tableNum
     * @param wb
     * @param sheetname
     * @param cdsl
     */
    private void createTable(int tableNum, XSSFWorkbook wb, String sheetname, int cdsl) {
        int endrow = 0;
        int printrow = 0;
        if (cdsl == 2){
            endrow = 46;
            printrow=47;
        }else if (cdsl == 3){
            endrow = 26;
            printrow=27;
        }else if (cdsl == 4 || cdsl ==5){
            endrow = 36;
            printrow=37;
        }
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            RowCopy.copyRows(wb, sheetname, sheetname, 0, endrow, i* (endrow+1));
        }
        if(record >= 1){
            wb.setPrintArea(wb.getSheetIndex(sheetname), 0, cdsl*4+8, 0,(record) * printrow-1);
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
        int a = 0;
        int b = 0;
        if (cdsl == 2){
            a = 47;
            b=40;
        }else if (cdsl == 3){
            a = 27;
            b=20;
        }else if (cdsl == 4 || cdsl ==5){
            a = 37;
            b=30;
        }
        double n = Double.valueOf(row.get("zdzh").toString()) / 1000;
        int m = (int) n;
        sheet.getRow(tableNum * a + 7 + index % b).getCell(0).setCellValue(m);
        sheet.getRow(tableNum * a + 7 + index % b).getCell(4).setCellValue(Double.valueOf(row.get("zdzh").toString()));
        String[] sfc = row.get("iri").toString().split(",");
        if (!sfc[0].isEmpty()) {
            for (int i = 0 ; i < sfc.length ; i++) {
                if (!sfc[i].equals("-")){
                    sheet.getRow(tableNum * a + 7 + index % b).getCell(5+i).setCellValue(Double.parseDouble(sfc[i]));
                }else {
                    sheet.getRow(tableNum * a + 7 + index % b).getCell((2*cdsl+5)+i).setCellValue(sfc[i]);
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
        if (cdsl == 2){
            a = 47;
        }else if (cdsl == 3){
            a = 27;
        }else if (cdsl == 4 || cdsl ==5){
            a = 37;
        }
        String fbgcname = "";
        if (sheetname.contains("隧道")){
            fbgcname = "隧道路面";
        }else if (sheetname.contains("桥")){
            fbgcname = "桥面系";
        }else {
            fbgcname = "路面面层";
        }
        sheet.getRow(tableNum * a + 1).getCell(5).setCellValue(proname);
        sheet.getRow(tableNum * a + 1).getCell(cdsl*4+4).setCellValue(htd);
        sheet.getRow(tableNum * a + 2).getCell(5).setCellValue("路面工程");
        sheet.getRow(tableNum * a + 3).getCell(cdsl*4+4).setCellValue(sj);
        sheet.getRow(tableNum * a + 2).getCell(cdsl*4+4).setCellValue(fbgcname+"("+name+")");
        sheet.getRow(tableNum * a + 3).getCell(5).setCellValue(Double.parseDouble(sjz));
    }

    /**
     *将相同幅的IRI拼接
     * @param list
     * @return
     */
    private static List<Map<String, Object>> montageIRI(List<Map<String, Object>> list) {
        if (list == null || list.isEmpty()){
            return new ArrayList<>();
        }else {
            Map<String, List<String>> resultMapz = new TreeMap<>();
            for (Map<String, Object> map : list) {
                String qdzh = map.get("qdzh").toString();
                //String ziri = map.get("ziri").toString();
                //String yiri = map.get("yiri").toString();
                String ziri = "";
                String yiri = "";
                if (map.get("ziri") == null){
                    ziri = "-";
                }else {
                    ziri = map.get("ziri").toString();
                }
                if (map.get("yiri") == null){
                    yiri = "-";
                }else {
                    yiri = map.get("yiri").toString();
                }
                if (resultMapz.containsKey(qdzh)) {
                    resultMapz.get(qdzh).add(ziri);
                    resultMapz.get(qdzh).add(yiri);
                } else {
                    List<String> sfcList = new ArrayList<>();
                    sfcList.add(ziri);
                    sfcList.add(yiri);
                    resultMapz.put(qdzh, sfcList);
                }
            }

            List<Map<String, Object>> resultList = new LinkedList<>();
            for (Map.Entry<String, List<String>> entry : resultMapz.entrySet()) {
                Map<String, Object> map = new TreeMap<>();
                map.put("qdzh", entry.getKey());
                map.put("iri", String.join(",", entry.getValue()));
                for (Map<String, Object> item : list) {
                    if (item.get("qdzh").toString().equals(entry.getKey())) {
                        map.put("qdzh", item.get("qdzh"));
                        map.put("zdzh", item.get("zdzh"));
                        map.put("pzlx", item.get("pzlx"));
                        map.put("name", item.get("name"));
                        map.put("cd", item.get("cd").toString().substring(0,2));
                        map.put("createTime", item.get("createTime"));
                        break;
                    }
                }
                resultList.add(map);
            }
            return resultList;
        }
    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        List<Map<String, Object>> mapList = new ArrayList<>();
        List<Map<String,Object>> lxlist = jjgZdhPzdMapper.selectlx(proname,htd);

        if (lxlist.size()>0){
            for (Map<String, Object> map : lxlist) {
                String zx = map.get("lxbs").toString();
                int num = jjgZdhPzdMapper.selectcdnum(proname,htd,zx);
                List<Map<String, Object>> looksdjdb = lookjdb(proname, htd, zx,num/2);
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
            f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "18路面平整度.xlsx");
        }else {
            f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "61互通平整度-"+zx+".xlsx");
        }
        if (!f.exists()) {
            return new ArrayList<>();
        } else {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
            List<Map<String, Object>> jgmap = new ArrayList<>();
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                if (!wb.isSheetHidden(wb.getSheetIndex(wb.getSheetAt(j)))) {
                    XSSFSheet slSheet = wb.getSheetAt(j);

                    XSSFCell xmname = slSheet.getRow(1).getCell(5);//项目名
                    XSSFCell htdname = slSheet.getRow(1).getCell(cds*4+4);//合同段名
                    Map map = new HashMap();

                    if (proname.equals(xmname.toString()) && htd.equals(htdname.toString())) {
                        slSheet.getRow(5).getCell(cds*4+9).setCellType(CellType.STRING);//总点数
                        slSheet.getRow(5).getCell(cds*4+10).setCellType(CellType.STRING);//合格点数

                        slSheet.getRow(5).getCell(cds*4+11).setCellType(CellType.STRING);
                        slSheet.getRow(5).getCell(cds*4+12).setCellType(CellType.STRING);

                        double zds = Double.valueOf(slSheet.getRow(5).getCell(cds*4+9).getStringCellValue());
                        double hgds = Double.valueOf(slSheet.getRow(5).getCell(cds*4+10).getStringCellValue());
                        String zdsz1 = decf.format(zds);
                        String hgdsz1 = decf.format(hgds);
                        map.put("检测项目", zx);
                        map.put("路面类型", wb.getSheetName(j));
                        map.put("总点数", zdsz1);
                        map.put("设计值", slSheet.getRow(3).getCell(5).getStringCellValue());
                        map.put("合格点数", hgdsz1);
                        map.put("合格率", zds!=0 ? df.format(hgds/zds*100) : 0);
                        map.put("Max",slSheet.getRow(5).getCell(cds*4+11).getStringCellValue());
                        map.put("Min",slSheet.getRow(5).getCell(cds*4+12).getStringCellValue());
                    }
                    jgmap.add(map);

                }
            }
            return jgmap;
        }
    }

    @Override
    public void exportpzd(HttpServletResponse response, String cdsl) throws IOException {
        int cd = Integer.parseInt(cdsl);
        String fileName = "平整度实测数据";
        String[][] sheetNames = {
                {"左幅一车道","左幅二车道","右幅一车道","右幅二车道"},
                {"左幅一车道","左幅二车道","左幅三车道","右幅一车道","右幅二车道","右幅三车道"},
                {"左幅一车道","左幅二车道","左幅三车道","左幅四车道","右幅一车道","右幅二车道","右幅三车道","右幅四车道"},
                {"左幅一车道","左幅二车道","左幅三车道","左幅四车道","左幅五车道","右幅一车道","右幅二车道","右幅三车道","右幅四车道","右幅五车道"}
        };
        String[] sheetName = sheetNames[cd-2];
        ExcelUtil.writeExcelMultipleSheets(response, null, fileName, sheetName, new JjgZdhPzdVo());

    }

    @Override
    public void importpzd(MultipartFile file, CommonInfoVo commonInfoVo) throws IOException {
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
                        .head(JjgZdhPzdVo.class)
                        .headRowNumber(1)
                        .registerReadListener(
                                new ExcelHandler<JjgZdhPzdVo>(JjgZdhPzdVo.class) {
                                    @Override
                                    public void handle(List<JjgZdhPzdVo> dataList) {
                                        for(JjgZdhPzdVo pzdVo: dataList)
                                        {
                                            JjgZdhPzd pzd = new JjgZdhPzd();
                                            BeanUtils.copyProperties(pzdVo,pzd);
                                            pzd.setCreatetime(new Date());
                                            pzd.setProname(commonInfoVo.getProname());
                                            pzd.setHtd(commonInfoVo.getHtd());
                                            pzd.setQdzh(Double.parseDouble(pzdVo.getQdzh()));
                                            if (!pzdVo.getZdzh().isEmpty() && pzdVo.getZdzh()!=null){
                                                pzd.setZdzh(Double.parseDouble(pzdVo.getZdzh()));
                                            }
                                            if (sheetName.contains("一")){
                                                pzd.setVal(1);
                                            }else if (sheetName.contains("二")){
                                                pzd.setVal(2);
                                            }else if (sheetName.contains("三")){
                                                pzd.setVal(3);
                                            }else if (sheetName.contains("四")){
                                                pzd.setVal(4);
                                            }else if (sheetName.contains("五")){
                                                pzd.setVal(5);
                                            }
                                            pzd.setCd(sheetName);
                                            jjgZdhPzdMapper.insert(pzd);
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
