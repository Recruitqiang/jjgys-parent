package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.*;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.zdh.JjgZdhMcxsVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.*;
import glgc.jjgys.system.service.JjgZdhMcxsService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
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
import java.text.Collator;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-06-06
 */
@Service
public class JjgZdhMcxsServiceImpl extends ServiceImpl<JjgZdhMcxsMapper, JjgZdhMcxs> implements JjgZdhMcxsService {

    @Autowired
    private JjgZdhMcxsMapper jjgZdhMcxsMapper;

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

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;


    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        /**
         * 分三种情况
         * 主线，只生成一张鉴定表，分别要判断桥梁和隧道
         * 匝道，生成多张，分别判断桥梁和隧道
         * 连接线，同匝道。
         */
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();

        List<Map<String,Object>> lxlist = jjgZdhMcxsMapper.selectlx(proname,htd);
        for (Map<String, Object> map : lxlist) {
            String zx = map.get("lxbs").toString();
            int num = jjgZdhMcxsMapper.selectcdnum(proname,htd,zx);
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
     * 处理主线的数据
     * @param proname
     * @param htd
     * @param zx
     */
    private void handlezxData(String proname, String htd, String zx, int cdsl,String sjz) throws IOException, ParseException {
        if (zx.equals("主线")){
            /**
             * 还需要判断时几车道，待处理
             * 先将左幅的数据归类，里面还包含左1，左2，左3，左4
             * 每个左1，2中的桩号都是一样的
             * 所以要取出桩号，桩号还需要处理一下，然后和基础数据中比对，分为主线，桥梁和隧道，然后分别写入到工作簿中。
             */
            List<Map<String,Object>> datazf = jjgZdhMcxsMapper.selectzfList(proname,htd,zx);
            List<Map<String,Object>> datayf = jjgZdhMcxsMapper.selectyfList(proname,htd,zx);

            QueryWrapper<JjgHtd> wrapperhtd = new QueryWrapper<>();
            wrapperhtd.like("proname",proname);
            wrapperhtd.like("name",htd);
            List<JjgHtd> htdList = jjgHtdMapper.selectList(wrapperhtd);

            //if (htdList.get(0).getZy().equals("ZY")){
            String htdzhq = htdList.get(0).getZhq();
            String htdzhz = htdList.get(0).getZhz();

            //隧道
            List<JjgLqsSd> jjgLqsSdzf = jjgLqsSdMapper.selectsdzf(proname,htdzhq,htdzhz,"左幅");
            List<JjgLqsSd> jjgLqsSdyf = jjgLqsSdMapper.selectsdyf(proname,htdzhq,htdzhz,"右幅");

            //桥
            List<JjgLqsQl> jjgLqsQlzf = jjgLqsQlMapper.selectqlzf(proname,htdzhq,htdzhz,"左幅");
            List<JjgLqsQl> jjgLqsQlyf = jjgLqsQlMapper.selectqlyf(proname,htdzhq,htdzhz,"右幅");


            List<Map<String,Object>> hpsdzfdata = new ArrayList<>();
            if (jjgLqsSdzf.size()>0){
                for (JjgLqsSd jjgLqsSd : jjgLqsSdzf) {
                    String zhq = String.valueOf((jjgLqsSd.getZhq())/1000);
                    String zhz = String.valueOf((jjgLqsSd.getZhz())/1000);
                    hpsdzfdata.addAll(jjgZdhMcxsMapper.selectSdZfData(proname,htd,zhq,zhz,String.valueOf(jjgLqsSd.getZhq()),String.valueOf(jjgLqsSd.getZhz())));
                }
            }
            List<Map<String,Object>> hpsdyfdata = new ArrayList<>();
            if (jjgLqsSdyf.size()>0){
                for (JjgLqsSd jjgLqsSd : jjgLqsSdyf) {
                    String zhq = String.valueOf((jjgLqsSd.getZhq())/1000);
                    String zhz = String.valueOf((jjgLqsSd.getZhz())/1000);
                    hpsdyfdata.addAll(jjgZdhMcxsMapper.selectSdyfData(proname,htd,zhq,zhz,String.valueOf(jjgLqsSd.getZhq()),String.valueOf(jjgLqsSd.getZhz())));
                }
            }
            List<Map<String,Object>> hpqlzfdata = new ArrayList<>();
            if (jjgLqsQlzf.size()>0){
                for (JjgLqsQl jjgLqsQl : jjgLqsQlzf) {
                    String zhq = String.valueOf((jjgLqsQl.getZhq())/1000);
                    String zhz = String.valueOf((jjgLqsQl.getZhz())/1000);
                    hpqlzfdata.addAll(jjgZdhMcxsMapper.selectQlZfData(proname,htd,zhq,zhz,String.valueOf(jjgLqsQl.getZhq()),String.valueOf(jjgLqsQl.getZhz())));
                }
            }
            List<Map<String,Object>> hpqlyfdata = new ArrayList<>();
            if (jjgLqsQlyf.size()>0){
                for (JjgLqsQl jjgLqsQl : jjgLqsQlyf) {
                    String zhq = String.valueOf((jjgLqsQl.getZhq())/1000);
                    String zhz = String.valueOf((jjgLqsQl.getZhz())/1000);
                    hpqlyfdata.addAll(jjgZdhMcxsMapper.selectQlYfData(proname,htd,zhq,zhz,String.valueOf(jjgLqsQl.getZhq()),String.valueOf(jjgLqsQl.getZhz())));
                }
            }

            //处理数据
            List<Map<String, Object>> sdzxList = groupByZh(hpsdzfdata);
            List<Map<String, Object>> sdyxList = groupByZh(hpsdyfdata);
            List<Map<String, Object>> qlzxList = groupByZh(hpqlzfdata);
            List<Map<String, Object>> qlyxList = groupByZh(hpqlyfdata);

            List<Map<String, Object>> lmzfList = groupByZh(datazf);
            List<Map<String, Object>> lmyfList = groupByZh(datayf);

            writeExcelData(proname,htd,lmzfList,lmyfList,sdzxList,sdyxList,qlzxList,qlyxList,cdsl,sjz,zx);
        }else if (zx.contains("连接线")){

            //查询的是摩擦系数表中的连接线
            List<Map<String,Object>> dataljxzf = jjgZdhMcxsMapper.selectzfList(proname,htd,zx);
            List<Map<String,Object>> dataljxyf = jjgZdhMcxsMapper.selectyfList(proname,htd,zx);
            //连接线
            QueryWrapper<JjgLjx> wrapperljx = new QueryWrapper<>();
            wrapperljx.like("proname",proname);
            wrapperljx.like("sshtd",htd);
            List<JjgLjx> jjgLjxList = jjgLqsLjxMapper.selectList(wrapperljx);

            List<Map<String,Object>> sdmcxs = new ArrayList<>();
            List<Map<String,Object>> qlmcxs = new ArrayList<>();

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
                    Double sdq = jjgLqsSd.getZhq()+10;
                    String sdz = jjgLqsSd.getZhz().toString();
                    String sdzhq = String.valueOf(sdq);
                    String zhq1 = String.valueOf((jjgLqsSd.getZhq()));
                    String zhz1 = String.valueOf((jjgLqsSd.getZhz()));
                    sdmcxs.addAll(jjgZdhMcxsMapper.selectsdmcxs(proname,bz,lf,sdzhq,sdz,zx,zhq1,zhz1));
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
                    qlmcxs.addAll(jjgZdhMcxsMapper.selectqlmcxs(proname,bz,lf, qlzhq, qlzhz, zx, zhq1, zhz1));

                }
            }

            List<Map<String,Object>> zfqlmcxs = new ArrayList<>();
            List<Map<String,Object>> yfqlmcxs = new ArrayList<>();
            if (qlmcxs.size()>0){
                for (int i = 0; i < qlmcxs.size(); i++) {
                    if (qlmcxs.get(i).get("cd").toString().contains("左幅")){
                        zfqlmcxs.add(qlmcxs.get(i));
                    }
                    if (qlmcxs.get(i).get("cd").toString().contains("右幅")){
                        yfqlmcxs.add(qlmcxs.get(i));
                    }
                }
            }
            List<Map<String,Object>> zfsdmcxs = new ArrayList<>();
            List<Map<String,Object>> yfsdmcxs = new ArrayList<>();
            if (sdmcxs.size()>0){
                for (Map<String, Object> sdmcx : sdmcxs) {
                    if (sdmcx.get("cd").toString().contains("左幅")){
                        zfsdmcxs.add(sdmcx);
                    }
                    if (sdmcx.get("cd").toString().contains("右幅")){
                        yfsdmcxs.add(sdmcx);
                    }
                }
            }

            List<Map<String, Object>> qlzfsj = groupByZh1(zfqlmcxs);
            List<Map<String, Object>> qlyfsj = groupByZh1(yfqlmcxs);
            List<Map<String, Object>> sdzfsj = groupByZh1(zfsdmcxs);
            List<Map<String, Object>> sdyfsj = groupByZh1(yfsdmcxs);

            List<Map<String, Object>> allzfsj = mergeList(dataljxzf);
            List<Map<String, Object>> allyfsj = mergeList(dataljxyf);
            writeExcelData(proname,htd,allzfsj,allyfsj,sdzfsj,sdyfsj,qlzfsj,qlyfsj,cdsl,sjz,zx);


        }else {
            //匝道
            List<Map<String,Object>> datazdzf = jjgZdhMcxsMapper.selectzfList(proname,htd,zx);
            List<Map<String,Object>> datazdyf = jjgZdhMcxsMapper.selectyfList(proname,htd,zx);
            //{zdzh=10.0, sfc=64, cd=左幅一车道, zdbs=A, qdzh=10.0, createTime=2023-06-12, name=淮宁湾立交左幅, htd=LM-1, proname=陕西高速, id=b364a4a742e921480ba480efa407b56d, lxbs=淮宁湾立交}

            /**
             * 先去匝道表中查询起始桩号,查出多个，然后根据当前这条数据的起始桩号，去隧道表中查有无隧道数据，附带一个条件bz去和隧道表中的bz匹配
             */
            QueryWrapper<JjgLqsHntlmzd> wrapperzd = new QueryWrapper<>();
            wrapperzd.like("proname",proname);
            wrapperzd.like("wz",zx);
            List<JjgLqsHntlmzd> zdList = jjgLqsHntlmzdMapper.selectList(wrapperzd);

            /**
             * 查出多条数据，分A，B,C匝道
             */
            List<Map<String,Object>> sdmcxs = new ArrayList<>();//左右幅或者单幅，当前匝道下全部的隧道数据
            List<Map<String,Object>> qlmcxs = new ArrayList<>();

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
                    Double sdq = jjgLqsSd.getZhq()+10;
                    String sdz = jjgLqsSd.getZhz().toString();
                    String sdzhq = String.valueOf(sdq);

                    String zhq1 = String.valueOf((jjgLqsSd.getZhq()));
                    String zhz1 = String.valueOf((jjgLqsSd.getZhz()));
                    sdmcxs.addAll(jjgZdhMcxsMapper.selectsdmcxs(proname,bz,lf,sdzhq,sdz,zx,zhq1,zhz1));
                }

                List<JjgLqsQl> jjgLqsql = jjgLqsQlMapper.selectqlList(proname,zhq,zhz,bz,wz,zdlf);
                for (JjgLqsQl jjgLqsQl : jjgLqsql) {
                    String lf = jjgLqsQl.getLf();
                    Double qlq = jjgLqsQl.getZhq()+10;
                    Double qlz = jjgLqsQl.getZhz();
                    String qlzhq = String.valueOf(qlq);
                    String qlzhz = String.valueOf(qlz);
                    String zhq1 = String.valueOf(jjgLqsQl.getZhq());
                    String zhz1 = String.valueOf(jjgLqsQl.getZhz());
                    //先留着，在摩擦系数表中是左右幅，没有单幅的，在桥梁表中有可能存在单幅的情况
                    /*if (lf.equals("单幅")){
                        lf = "左幅";
                    }*/
                    qlmcxs.addAll(jjgZdhMcxsMapper.selectqlmcxs(proname,bz,lf, qlzhq, qlzhz, zx, zhq1, zhz1));

                }

            }
            //{zdzh=130.0, sfc=57, cd=左幅一车道, zdbs=E, qdzh=130.0, createTime=2023-06-12, name=淮宁湾立交EK0+157.338匝道桥, htd=LM-1, proname=陕西高速, id=0a4480e6b314848249fe87bfba96861e, lxbs=淮宁湾立交}
            List<Map<String,Object>> zfqlmcxs = new ArrayList<>();
            List<Map<String,Object>> yfqlmcxs = new ArrayList<>();
            if (qlmcxs.size()>0){
                for (int i = 0; i < qlmcxs.size(); i++) {
                    if (qlmcxs.get(i).get("cd").toString().contains("左幅")){
                        zfqlmcxs.add(qlmcxs.get(i));
                    }
                    if (qlmcxs.get(i).get("cd").toString().contains("右幅")){
                        yfqlmcxs.add(qlmcxs.get(i));
                    }
                }
            }
            List<Map<String,Object>> zfsdmcxs = new ArrayList<>();
            List<Map<String,Object>> yfsdmcxs = new ArrayList<>();
            if (sdmcxs.size()>0){
                for (Map<String, Object> sdmcx : sdmcxs) {
                    if (sdmcx.get("cd").toString().contains("左幅")){
                        zfsdmcxs.add(sdmcx);
                    }
                    if (sdmcx.get("cd").toString().contains("右幅")){
                        yfsdmcxs.add(sdmcx);
                    }
                }
            }

            //{zdzh=130.0, sfc=57, cd=左幅一车道, zdbs=E, qdzh=130.0, createTime=2023-06-12, name=淮宁湾立交EK0+157.338匝道桥, htd=LM-1, proname=陕西高速, id=0a4480e6b314848249fe87bfba96861e, lxbs=淮宁湾立交},
            List<Map<String, Object>> qlzfsj = groupByZh1(zfqlmcxs);
            List<Map<String, Object>> qlyfsj = groupByZh1(yfqlmcxs);
            List<Map<String, Object>> sdzfsj = groupByZh1(zfsdmcxs);
            List<Map<String, Object>> sdyfsj = groupByZh1(yfsdmcxs);

            //{zdzh=10.0, sfc=64, cd=左幅一车道, zdbs=A, qdzh=10.0, createTime=2023-06-12, name=淮宁湾立交左幅, htd=LM-1, proname=陕西高速, id=b364a4a742e921480ba480efa407b56d, lxbs=淮宁湾立交}
            List<Map<String, Object>> allzfsj = mergeList(datazdzf);
            List<Map<String, Object>> allyfsj = mergeList(datazdyf);
            //writeExcelData(proname,htd,allzfsj,allyfsj,sdzfsj,sdyfsj,qlzfsj,qlyfsj,cdsl,sjz,zx);
            writeExcelData(proname,htd,allzfsj,allyfsj,sdzfsj,sdyfsj,qlzfsj,qlyfsj,cdsl,sjz,zx);

        }

    }


    /**
     *
     * @param data
     * @return
     */
    private static List<Map<String,Object>>  sortList(List<Map<String,Object>> data){

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
     * @param proname
     * @param htd
     * @param data
     * @param zfsdqlData
     * @param wb
     * @param sheetname
     * @param cdsl
     * @param sjz
     * @param zx
     * @throws ParseException
     */
    private void DBtoExcelZd(String proname, String htd, List<Map<String, Object>> data, List<Map<String, Object>> zfsdqlData, XSSFWorkbook wb, String sheetname,int cdsl,String sjz,String zx) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        SimpleDateFormat outputDateFormat  = new SimpleDateFormat("yyyy.MM.dd");
        if (data!=null && !data.isEmpty()){
            createTableZd(getNum1(data),wb,sheetname,cdsl);
            XSSFSheet sheet = wb.getSheet(sheetname);
            if (cdsl ==2){
                sheet.getRow(1).getCell(1).setCellValue(proname);
                sheet.getRow(1).getCell(7).setCellValue(htd);
            }else if(cdsl == 3) {
                sheet.getRow(1).getCell(2).setCellValue(proname);
                sheet.getRow(1).getCell(9).setCellValue(htd);
            }else if (cdsl ==4){
                sheet.getRow(1).getCell(2).setCellValue(proname);
                sheet.getRow(1).getCell(12).setCellValue(htd);
            }
            String name = data.get(0).get("zdbs").toString()+"匝道";
            int index = 0;
            int tableNum = 0;
            String time1 = String.valueOf(data.get(0).get("createTime")) ;
            Date parse = simpleDateFormat.parse(time1);
            String time = outputDateFormat.format(parse);

            fillZdTitleCellData(sheet, tableNum, proname, htd, name,time,sheetname,cdsl,sjz);

            /**
             *处理数据，将data和zfsdqlData中相同的qdzh的name为空
             * 还得查一下收费站的数据
             */
            List<Map<String, Object>> lmdata = handlezdData(proname,data,zfsdqlData,zx);

            if (lmdata.size()>0) {
                List<Map<String, Object>> rowAndcol = new ArrayList<>();
                int startRow = -1, endRow = -1, startCol = -1, endCol = -1;

                String zdbs = lmdata.get(0).get("zdbs").toString();
                for (Map<String, Object> lm : lmdata) {
                    if (lm.get("zdbs").toString().equals(zdbs)) {
                        if (index > 99) {
                            tableNum++;
                            fillZdTitleCellData(sheet, tableNum, proname, htd, lm.get("zdbs").toString() + "匝道", time, sheetname, cdsl, sjz);
                            index = 0;
                        }
                        if (!lm.get("sfc").toString().equals("") && !lm.get("sfc").toString().isEmpty()) {
                            String[] sfc = lm.get("sfc").toString().split(",");
                            for (int i = 0; i < sfc.length; i++) {
                                if (index < 69) {
                                    sheet.getRow(tableNum * 44 + 6 + index % 38).getCell((cdsl + 1) * (index / 38)).setCellValue((Double.parseDouble(lm.get("qdzh").toString())));
                                    if (sfc[i].equals("-")){
                                        sheet.getRow(tableNum * 44 + 6 + index % 38).getCell((cdsl + 1) * (index / 38) + 1 + i).setCellValue("-");
                                    }else {
                                        sheet.getRow(tableNum * 44 + 6 + index % 38).getCell((cdsl + 1) * (index / 38) + 1 + i).setCellValue(Double.parseDouble(sfc[i]));
                                    }


                                } else {
                                    sheet.getRow(tableNum * 44 + 6 + index % 69).getCell((cdsl * 2 + 2) * (index / 69)).setCellValue((Double.parseDouble(lm.get("qdzh").toString())));
                                    if (sfc[i].equals("-")){
                                        sheet.getRow(tableNum * 44 + 6 + index % 69).getCell((cdsl * 2 + 2) * (index / 69) + 1 + i).setCellValue("-");
                                    }else {
                                        sheet.getRow(tableNum * 44 + 6 + index % 69).getCell((cdsl * 2 + 2) * (index / 69) + 1 + i).setCellValue(Double.parseDouble(sfc[i]));
                                    }

                                }
                            }
                        } else {
                            for (int i = 0; i < cdsl; i++) {
                                if (index < 69) {
                                    sheet.getRow(tableNum * 44 + 6 + index % 38).getCell((cdsl + 1) * (index / 38)).setCellValue((Double.parseDouble(lm.get("qdzh").toString())));
                                    sheet.getRow(tableNum * 44 + 6 + index % 38).getCell((cdsl + 1) * (index / 38) + 1 + i).setCellValue(lm.get("name").toString());

                                    startRow = tableNum * 44 + 6 + index % 38;
                                    endRow = tableNum * 44 + 6 + index % 38;
                                    startCol = (cdsl + 1) * (index / 38) + 1;
                                    endCol = (cdsl + 1) * (index / 38) + cdsl;

                                } else {
                                    sheet.getRow(tableNum * 44 + 6 + index % 69).getCell((cdsl * 2 + 2) * (index / 69)).setCellValue((Double.parseDouble(lm.get("qdzh").toString())));
                                    sheet.getRow(tableNum * 44 + 6 + index % 69).getCell((cdsl * 2 + 2) * (index / 69) + 1 + i).setCellValue(lm.get("name").toString());

                                    startRow = tableNum * 44 + 6 + index % 69;
                                    endRow = tableNum * 44 + 6 + index % 69;
                                    startCol = 2 * cdsl + 3;
                                    endCol = 3 * cdsl + 2;

                                }
                            }
                            //可以在这块记录一个行和列
                            Map<String, Object> map = new HashMap<>();
                            map.put("startRow", startRow);
                            map.put("endRow", endRow);
                            map.put("startCol", startCol);
                            map.put("endCol", endCol);
                            map.put("name", lm.get("name"));
                            map.put("tableNum", tableNum);
                            rowAndcol.add(map);

                        }
                        index++;
                    } else {
                        zdbs = lm.get("zdbs").toString();
                        tableNum++;
                        index = 0;
                        fillZdTitleCellData(sheet, tableNum, proname, htd, zdbs + "匝道", time, sheetname, cdsl, sjz);
                        if (index > 99) {
                            tableNum++;
                            fillTitleCellData(sheet, tableNum, proname, htd, zdbs + "匝道", time, sheetname, cdsl, sjz);
                            index = 0;
                        }
                        if (!lm.get("sfc").toString().equals("") && !lm.get("sfc").toString().isEmpty()) {
                            String[] sfc = lm.get("sfc").toString().split(",");
                            for (int i = 0; i < sfc.length; i++) {
                                if (index < 69) {
                                    sheet.getRow(tableNum * 44 + 6 + index % 38).getCell((cdsl + 1) * (index / 38)).setCellValue((Double.parseDouble(lm.get("qdzh").toString())));
                                    if (sfc[i].equals("-")){
                                        sheet.getRow(tableNum * 44 + 6 + index % 38).getCell((cdsl + 1) * (index / 38) + 1 + i).setCellValue(Double.parseDouble("-"));
                                    }else {
                                        sheet.getRow(tableNum * 44 + 6 + index % 38).getCell((cdsl + 1) * (index / 38) + 1 + i).setCellValue(Double.parseDouble(sfc[i]));
                                    }


                                } else {
                                    sheet.getRow(tableNum * 44 + 6 + index % 69).getCell((cdsl * 2 + 2) * (index / 69)).setCellValue((Double.parseDouble(lm.get("qdzh").toString())));
                                    if (sfc[i].equals("-")){
                                        sheet.getRow(tableNum * 44 + 6 + index % 69).getCell((cdsl * 2 + 2) * (index / 69) + 1 + i).setCellValue("-");
                                    }else {
                                        sheet.getRow(tableNum * 44 + 6 + index % 69).getCell((cdsl * 2 + 2) * (index / 69) + 1 + i).setCellValue(Double.parseDouble(sfc[i]));
                                    }

                                }
                            }
                        } else {
                            for (int i = 0; i < cdsl; i++) {
                                if (index < 69) {
                                    sheet.getRow(tableNum * 44 + 6 + index % 38).getCell((cdsl + 1) * (index / 38)).setCellValue((Double.parseDouble(lm.get("qdzh").toString())));
                                    sheet.getRow(tableNum * 44 + 6 + index % 38).getCell((cdsl + 1) * (index / 38) + 1 + i).setCellValue(lm.get("name").toString());

                                    startRow = tableNum * 44 + 6 + index % 38;
                                    endRow = tableNum * 44 + 6 + index % 38;
                                    startCol = (cdsl + 1) * (index / 38) + 1;
                                    endCol = (cdsl + 1) * (index / 38) + cdsl;

                                } else {
                                    sheet.getRow(tableNum * 44 + 6 + index % 69).getCell((cdsl * 2 + 2) * (index / 69)).setCellValue((Double.parseDouble(lm.get("qdzh").toString())));
                                    sheet.getRow(tableNum * 44 + 6 + index % 69).getCell((cdsl * 2 + 2) * (index / 69) + 1 + i).setCellValue(lm.get("name").toString());

                                    startRow = tableNum * 44 + 6 + index % 69;
                                    endRow = tableNum * 44 + 6 + index % 69;
                                    startCol = 2 * cdsl + 3;
                                    endCol = 3 * cdsl + 2;

                                }
                            }
                            //可以在这块记录一个行和列
                            Map<String, Object> map = new HashMap<>();
                            map.put("startRow", startRow);
                            map.put("endRow", endRow);
                            map.put("startCol", startCol);
                            map.put("endCol", endCol);
                            map.put("name", lm.get("name"));
                            map.put("tableNum", tableNum);
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

    }


    /**
     * 合并匝道的sfc值，需要根据name和qdzh
     * @param list
     * @return
     */
    private static List<Map<String, Object>> mergeList(List<Map<String, Object>> list) {
        if (list == null || list.isEmpty()){
            return new ArrayList<>();
        }else {
            Map<String, Map<String, List<String>>> map = new HashMap<>();
            for (Map<String, Object> item : list) {
                String zdbs = String.valueOf(item.get("zdbs"));
                String qdzh = String.valueOf(item.get("qdzh"));
                //String sfc = String.valueOf(item.get("sfc"));
                String sfc = "";
                if (item.get("sfc") == null){
                    sfc = "-";
                }else {
                    sfc = item.get("sfc").toString();
                }
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
                    values.put("qdzh", qdzh);
                    values.put("sfc", sfc);
                    // 遍历整个list，查找相同的name和createTime
                    boolean flag = true;
                    for (Map<String, Object> item : list) {
                        if (!String.valueOf(item.get("qdzh")).equals(qdzh) || !String.valueOf(item.get("zdbs")).equals(entry.getKey())) {
                            continue;
                        }
                        if (flag) { // 第一次找到匹配的元素，将name和createTime保存到values中
                            values.put("name", item.get("name"));
                            values.put("createTime", item.get("createTime"));
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
                    Double qdzh1 = Double.parseDouble(o1.get("qdzh").toString());
                    Double qdzh2 = Double.parseDouble(o2.get("qdzh").toString());
                    return qdzh1.compareTo(qdzh2);
                }
            });
            return result;

        }
    }



    /**
     *
     * @param proname
     * @param htd
     * @param sdzxList
     * @param sdyxList
     * @param qlzxList
     * @param qlyxList
     * @param cdsl
     */
    private void writeExcelData(String proname, String htd, List<Map<String, Object>> lmzflist, List<Map<String, Object>> lmyflist, List<Map<String, Object>> sdzxList, List<Map<String, Object>> sdyxList, List<Map<String, Object>> qlzxList, List<Map<String, Object>> qlyxList, int cdsl,String sjz,String zx) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        String fname="";
        if (zx.equals("主线")){
            fname = "19路面摩擦系数.xlsx";
        }else {
            fname = "62互通摩擦系数-"+zx+".xlsx";
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
        String sheetlmname = "";
        String sheetqname = "";
        String sheetsname = "";
        if (zx.equals("主线")){
            if (cdsl == 5){
                filename = "摩擦系数-5车道.xlsx";
            }else if (cdsl == 4){
                filename = "摩擦系数-4车道.xlsx";
            }else if (cdsl == 3){
                filename = "摩擦系数-3车道.xlsx";
            }else if (cdsl == 2){
                filename = "摩擦系数-2车道.xlsx";
            }
            sheetlmname="路面";
            sheetqname="桥";
            sheetsname="隧道";
        }else if(zx.contains("连接线")) {
            if (cdsl == 3){
                filename = "摩擦系数3.xlsx";
            }else if (cdsl == 2){
                filename = "摩擦系数2.xlsx";
            }else if (cdsl == 4){
                filename = "摩擦系数4.xlsx";
            }
            sheetlmname="路面";
            sheetqname="桥";
            sheetsname="隧道";
        }else {
            if (cdsl == 3){
                filename = "摩擦系数3.xlsx";
            }else if (cdsl == 2){
                filename = "摩擦系数2.xlsx";
            }else if (cdsl == 4){
                filename = "摩擦系数4.xlsx";
            }
            sheetlmname="匝道路面";
            sheetqname="匝道桥";
            sheetsname="匝道隧道";
        }
        String path = reportPath + File.separator + filename;
        Files.copy(Paths.get(path), new FileOutputStream(f));

        FileInputStream out = new FileInputStream(f);
        wb = new XSSFWorkbook(out);

        List<Map<String,Object>> zfsdqlData = new ArrayList<>();
        List<Map<String,Object>> yfsdqlData = new ArrayList<>();

        if (sdzxList.size() >0 && !sdzxList.isEmpty()){
            for (Map<String, Object> map : sdzxList) {
                Map<String,Object> map1 = new HashMap<>();
                map1.put("zh",map.get("qdzh").toString());
                map1.put("name",map.get("name").toString());
                map1.put("cd","左幅");
                if (map.get("zdbs")!=null){
                    map1.put("zdbs",map.get("zdbs").toString());
                }
                zfsdqlData.add(map1);
            }
            DBtoExcel(proname,htd,sdzxList,wb,"左幅-"+sheetsname,cdsl,sjz,zx);
        }
        if (sdyxList.size() >0 && !sdyxList.isEmpty()){
            for (Map<String, Object> map : sdyxList) {
                Map<String,Object> map1 = new HashMap<>();
                map1.put("zh",map.get("qdzh").toString());
                map1.put("name",map.get("name").toString());
                map1.put("cd","右幅");
                if (map.get("zdbs")!=null){
                    map1.put("zdbs",map.get("zdbs").toString());
                }
                yfsdqlData.add(map1);
            }
            DBtoExcel(proname,htd,sdyxList,wb,"右幅-"+sheetsname,cdsl,sjz,zx);
        }
        if (qlzxList.size() >0 && !qlzxList.isEmpty()){
            for (Map<String, Object> map : qlzxList) {
                Map<String,Object> map1 = new HashMap<>();
                map1.put("zh",map.get("qdzh").toString());
                map1.put("name",map.get("name").toString());
                map1.put("cd","左幅");
                if (map.get("zdbs")!=null){
                    map1.put("zdbs",map.get("zdbs").toString());
                }
                zfsdqlData.add(map1);
            }
            DBtoExcel(proname,htd,qlzxList,wb,"左幅-"+sheetqname,cdsl,sjz,zx);
        }
        if (qlyxList.size() >0 && !qlyxList.isEmpty()){
            for (Map<String, Object> map : qlyxList) {
                Map<String,Object> map1 = new HashMap<>();
                map1.put("zh",map.get("qdzh").toString());
                map1.put("name",map.get("name").toString());
                map1.put("cd","右幅");
                if (map.get("zdbs")!=null){
                    map1.put("zdbs",map.get("zdbs").toString());
                }
                yfsdqlData.add(map1);
            }
            DBtoExcel(proname,htd,qlyxList,wb,"右幅-"+sheetqname,cdsl,sjz,zx);
        }


        List<Map<String, Object>> zsdql = sortList(zfsdqlData);
        List<Map<String, Object>> ysdql = sortList(yfsdqlData);

        /**
         * 这块需要分别将隧道和桥梁的桩号取出，给写入路面的时候使用
         */

        if (lmzflist.size() >0 && !lmzflist.isEmpty()){
            if (zx.equals("主线")){
                DBtoExcelLm(proname,htd,lmzflist,zsdql,wb,"左幅-"+sheetlmname,cdsl,sjz);
            }else {
                DBtoExcelZd(proname,htd,lmzflist,zsdql,wb,"左幅-"+sheetlmname,cdsl,sjz,zx);
            }

        }
        if (lmyflist.size() >0 && !lmyflist.isEmpty()){
            if (zx.equals("主线")){
                DBtoExcelLm(proname,htd,lmyflist,ysdql,wb,"右幅-"+sheetlmname,cdsl,sjz);
            }else {
                DBtoExcelZd(proname,htd,lmyflist,ysdql,wb,"右幅-"+sheetlmname,cdsl,sjz,zx);
            }

        }
        String[] arr;
        if (zx.equals("主线")){
            arr = new String[]{"左幅-路面","右幅-路面","左幅-桥","右幅-桥","左幅-隧道","右幅-隧道"};
        }else {
            arr = new String[]{"左幅-路面","右幅-路面","左幅-桥","右幅-桥","左幅-隧道","右幅-隧道","左幅-匝道路面","右幅-匝道路面","左幅-匝道隧道","右幅-匝道隧道","右幅-匝道桥","左幅-匝道桥"};
        }
        for (int i = 0;i<arr.length;i++){
            if (shouldBeCalculate(wb.getSheet(arr[i]))){
                calculateSheet(wb,wb.getSheet(arr[i]),cdsl);
                JjgFbgcCommonUtils.updateFormula(wb, wb.getSheet(arr[i]));
            }else {
                wb.removeSheetAt(wb.getSheetIndex(arr[i]));
            }
        }
        wb.removeSheetAt(wb.getSheetIndex("保证率系数"));
        System.out.println("成功生成鉴定表");
        FileOutputStream fileOut = new FileOutputStream(f);
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
        out.close();
        wb.close();
    }


    /**
     * 是否为空
     * @param sheet
     * @return
     */
    private boolean shouldBeCalculate(XSSFSheet sheet) {
        sheet.getRow(6).getCell(0).setCellType(CellType.STRING);
        if(sheet.getRow(6).getCell(0)==null ||"".equals(sheet.getRow(6).getCell(0).getStringCellValue())){
            return false;
        }
        return true;
    }

    /**
     *
     * @param proname
     * @param htd
     * @param data
     * @param zfsdqlData
     * @param wb
     * @param sheetname
     */
    private void DBtoExcelLm(String proname, String htd, List<Map<String, Object>> data, List<Map<String, Object>> zfsdqlData, XSSFWorkbook wb, String sheetname,int cdsl,String sjz) throws ParseException {
        System.out.println(sheetname+"开始数据写入");
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        SimpleDateFormat outputDateFormat  = new SimpleDateFormat("yyyy.MM.dd");
        if (data!=null && !data.isEmpty()){
            createTable2(getNum2(data,cdsl),wb,sheetname,cdsl);
            XSSFSheet sheet = wb.getSheet(sheetname);
            if (cdsl ==2){
                sheet.getRow(1).getCell(1).setCellValue(proname);
                sheet.getRow(1).getCell(7).setCellValue(htd);
            }else if (cdsl == 5){
                sheet.getRow(1).getCell(2).setCellValue(proname);
                sheet.getRow(1).getCell(10).setCellValue(htd);
            }else {
                sheet.getRow(1).getCell(cdsl+2).setCellValue(proname);
                if (cdsl == 3){
                    sheet.getRow(1).getCell(9).setCellValue(htd);
                }else {
                    sheet.getRow(1).getCell(cdsl*2+4).setCellValue(htd);
                }
            }
            String name = data.get(0).get("name").toString();
            int index = 0;
            int tableNum = 0;
            String time1 = String.valueOf(data.get(0).get("createTime")) ;
            Date parse = simpleDateFormat.parse(time1);
            String time = outputDateFormat.format(parse);

            fillTitleCellData(sheet, tableNum, proname, htd, name,time,sheetname,cdsl,sjz);
            List<Map<String, Object>> lmdata = handleLmData(data,zfsdqlData);
            List<Map<String, Object>> rowAndcol = new ArrayList<>();
            int startRow = -1, endRow = -1, startCol = -1, endCol = -1;
            int s1 = 0;
            int s2 = 0;
            int s3  =0;
            if (cdsl == 5){
                s1 = 69;
                s2 = 38;
                s3 = 6;
            }else {
                s1 = 99;
                s2 = 69;
                s3 = cdsl*2+2;
            }
            for (Map<String, Object> lm : lmdata) {
                //if (index > s1) {
                if(index % s1 == 0 && index != 0){
                    tableNum++;
                    fillTitleCellData(sheet, tableNum, proname, htd, name, time, sheetname,cdsl,sjz);
                    index = 0;
                }
                if (!lm.get("sfc").toString().equals("") && !lm.get("sfc").toString().isEmpty()) {
                    String[] sfc = lm.get("sfc").toString().split(",");
                    for (int i = 0; i < sfc.length; i++) {
                        if (index < s2) {
                            sheet.getRow(tableNum * 44 + 6 + index % 38).getCell((cdsl+1) * (index / 38)).setCellValue((Double.parseDouble(lm.get("qdzh").toString())) * 1000);
                            if (sfc[i].equals("-")){
                                sheet.getRow(tableNum * 44 + 6 + index % 38).getCell((cdsl+1) * (index / 38) + 1 + i).setCellValue(sfc[i]);
                            }else {
                                sheet.getRow(tableNum * 44 + 6 + index % 38).getCell((cdsl+1) * (index / 38) + 1 + i).setCellValue(Double.parseDouble(sfc[i]));
                            }

                        } else {
                            sheet.getRow(tableNum * 44 + 6 + index % s2).getCell(s3  * (index / s2)).setCellValue((Double.parseDouble(lm.get("qdzh").toString())) * 1000);
                            if (sfc[i].equals("-")){
                                sheet.getRow(tableNum * 44 + 6 + index % s2).getCell(s3  * (index / s2) + 1 + i).setCellValue(sfc[i]);
                            }else {
                                sheet.getRow(tableNum * 44 + 6 + index % s2).getCell(s3  * (index / s2) + 1 + i).setCellValue(Double.parseDouble(sfc[i]));
                            }
                        }
                    }
                } else {
                    for (int i = 0; i < cdsl; i++) {
                        if (index < s2) {
                            sheet.getRow(tableNum * 44 + 6 + index % 38).getCell((cdsl+1) * (index / 38)).setCellValue((Double.parseDouble(lm.get("qdzh").toString())) * 1000);
                            sheet.getRow(tableNum * 44 + 6 + index % 38).getCell((cdsl+1) * (index / 38) + 1 + i).setCellValue(lm.get("name").toString());

                            startRow = tableNum * 44 + 6 + index % 38 ;
                            endRow = tableNum * 44 + 6 + index % 38 ;

                            startCol = (cdsl+1)  * (index / 38) + 1;
                            endCol = (cdsl+1)  * (index / 38) + cdsl;

                        } else {
                            sheet.getRow(tableNum * 44 + 6 + index % s2).getCell(s3  * (index / s2)).setCellValue((Double.parseDouble(lm.get("qdzh").toString())) * 1000);
                            sheet.getRow(tableNum * 44 + 6 + index % s2).getCell(s3  * (index / s2) + 1 + i).setCellValue(lm.get("name").toString());

                            startRow = tableNum * 44 + 6 + index % s2 ;
                            endRow = tableNum * 44 + 6 + index % s2 ;
                            if (cdsl == 5){
                                startCol = 7;
                                endCol = 11;
                            }else {
                                startCol = 2*cdsl+3;
                                endCol = 3*cdsl+2;
                            }
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
        System.out.println(sheetname+"数据写入完成");

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
     * @param zfsdqlData
     * @return
     */
    private List<Map<String, Object>> handleLmData(List<Map<String, Object>> data, List<Map<String, Object>> zfsdqlData) {
        for (Map<String, Object> datum : data) {
            for (Map<String, Object> zfsdqlDatum : zfsdqlData) {
                if (datum.get("qdzh").toString().equals(zfsdqlDatum.get("zh"))){
                    datum.put("sfc","");
                    datum.put("name",zfsdqlDatum.get("name"));
                }
            }
        }
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
        //if (zfsdqlData.size()>0){
        //String lf = zfsdqlData.get(0).get("cd").toString();
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
                    if (datum.get("zdbs").toString().equals(zfsdqlDatum.get("zdbs")) && datum.get("qdzh").toString().equals(zfsdqlDatum.get("zh"))) {
                        datum.put("sfc", "");
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
                Double qdzh1 = Double.parseDouble(o1.get("qdzh").toString());
                Double qdzh2 = Double.parseDouble(o2.get("qdzh").toString());
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
            for (double i = start; i <= end; i += 10) {
                Map map = new HashMap();
                map.put("qdzh",i);
                map.put("name",zdsfzname);
                map.put("zdbs",sszd);
                map.put("sfc","");
                result.add(map);
            }
            return result;
        } else {
            return new ArrayList();

        }
    }


    /**
     *
     * @param sheet
     */
    private void calculateSheet(XSSFWorkbook wb,XSSFSheet sheet,int cds) {
        XSSFRow row = null;
        XSSFRow rowstart = null;
        XSSFRow rowend = null;
        boolean flag = false;
        XSSFRow recordrow = null;
        String name = "";

        FormulaEvaluator e = new XSSFFormulaEvaluator(wb);

        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            // 下一张表
            if (!"".equals(row.getCell(0).toString()) && row.getCell(0).toString().contains("质量鉴定表") && flag) {
                flag = false;
            }
            if(flag){
                if(i == rowend.getRowNum()){
                    /*
                     * 此处对总计进行计算
                     */
                    calculateTotalData(sheet, rowstart, rowend,e,cds);
                    if(i+38 < sheet.getPhysicalNumberOfRows()){
                        rowstart = sheet.getRow(i+1);
                        rowend = sheet.getRow(i+38);
                    }
                    else{
                        break;
                    }
                }
            }

            if ("桩号".equals(row.getCell(0).toString())) {
                rowstart = sheet.getRow(i+3);
                rowend = sheet.getRow(i+40);
                i += 2;
                flag = true;
                if(recordrow == null){
                    recordrow = rowstart;
                }
                int cdnum = 0;
                if (cds == 2 || cds == 3 ){
                    cdnum = cds+2;
                }else if (cds == 5) {
                    cdnum = 5;
                }else {
                    cdnum = cds+3;
                }
                if (cds == 5){
                    if(!"".equals(name) && !name.equals(sheet.getRow(i-1-2-1).getCell(cdnum).toString())){

                        sheet.getRow(recordrow.getRowNum()-4).createCell(12).setCellFormula("SUM("
                                +recordrow.createCell(12).getReference()+":"
                                +sheet.getRow(i-4-2).createCell(12).getReference()+")");

                        double value = e.evaluate(sheet.getRow(recordrow.getRowNum()-4).getCell(12)).getNumberValue();
                        sheet.getRow(recordrow.getRowNum()-4).getCell(12).setCellFormula(null);
                        sheet.getRow(recordrow.getRowNum()-4).getCell(12).setCellValue(value);

                        sheet.getRow(recordrow.getRowNum()-4).createCell(13).setCellFormula("SUM("
                                +recordrow.createCell(13).getReference()+":"
                                +sheet.getRow(i-4-2).createCell(13).getReference()+")");
                        value = e.evaluate(sheet.getRow(recordrow.getRowNum()-4).getCell(13)).getNumberValue();
                        sheet.getRow(recordrow.getRowNum()-4).getCell(13).setCellFormula(null);
                        sheet.getRow(recordrow.getRowNum()-4).getCell(13).setCellValue(value);

                        sheet.getRow(recordrow.getRowNum()-4).createCell(14).setCellFormula(
                                sheet.getRow(recordrow.getRowNum()-4).getCell(13).getReference()+"*100/"
                                        + sheet.getRow(recordrow.getRowNum()-4).getCell(12).getReference());

                        value = e.evaluate(sheet.getRow(recordrow.getRowNum()-4).getCell(14)).getNumberValue();
                        sheet.getRow(recordrow.getRowNum()-4).getCell(14).setCellFormula(null);
                        sheet.getRow(recordrow.getRowNum()-4).getCell(14).setCellValue(value);

                        recordrow = rowstart;
                    }
                    name = sheet.getRow(i-1-2-1).getCell(cdnum).toString();

                }else {
                    if(!"".equals(name) && !name.equals(sheet.getRow(i-1-2-1).getCell(cdnum).toString())){

                        sheet.getRow(recordrow.getRowNum()-4).createCell(3*cds+3).setCellFormula("SUM("
                                +recordrow.createCell(3*cds+3).getReference()+":"
                                +sheet.getRow(i-4-2).createCell(3*cds+3).getReference()+")");

                        double value = e.evaluate(sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+3)).getNumberValue();
                        sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+3).setCellFormula(null);
                        sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+3).setCellValue(value);

                        sheet.getRow(recordrow.getRowNum()-4).createCell(3*cds+4).setCellFormula("SUM("
                                +recordrow.createCell(3*cds+4).getReference()+":"
                                +sheet.getRow(i-4-2).createCell(3*cds+4).getReference()+")");
                        value = e.evaluate(sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+4)).getNumberValue();
                        sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+4).setCellFormula(null);
                        sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+4).setCellValue(value);

                        sheet.getRow(recordrow.getRowNum()-4).createCell(3*cds+5).setCellFormula(
                                sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+4).getReference()+"*100/"
                                        + sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+3).getReference());

                        value = e.evaluate(sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+5)).getNumberValue();
                        sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+5).setCellFormula(null);
                        sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+5).setCellValue(value);

                        recordrow = rowstart;
                    }
                    name = sheet.getRow(i-1-2-1).getCell(cdnum).toString();
                }

            }
        }
        if (cds == 5){
            //计算最后一座桥或隧道
            sheet.getRow(recordrow.getRowNum()-4).createCell(12)
                    .setCellFormula("SUM(" +recordrow.createCell(12).getReference()+":"
                            +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(12).getReference()+")");
            double value = e.evaluate(sheet.getRow(recordrow.getRowNum()-4).getCell(12)).getNumberValue();
            sheet.getRow(recordrow.getRowNum()-4).getCell(12).setCellFormula(null);
            sheet.getRow(recordrow.getRowNum()-4).getCell(12).setCellValue(value);

            sheet.getRow(recordrow.getRowNum()-4).createCell(13).setCellFormula("SUM("
                    +recordrow.createCell(13).getReference()+":"
                    +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(13).getReference()+")");
            value = e.evaluate(sheet.getRow(recordrow.getRowNum()-4).getCell(13)).getNumberValue();
            sheet.getRow(recordrow.getRowNum()-4).getCell(13).setCellFormula(null);
            sheet.getRow(recordrow.getRowNum()-4).getCell(13).setCellValue(value);

            sheet.getRow(recordrow.getRowNum()-4).createCell(14).setCellFormula(
                    sheet.getRow(recordrow.getRowNum()-4).getCell(13).getReference()+"*100/"
                            + sheet.getRow(recordrow.getRowNum()-4).getCell(12).getReference());
            value = e.evaluate(sheet.getRow(recordrow.getRowNum()-4).getCell(14)).getNumberValue();
            sheet.getRow(recordrow.getRowNum()-4).getCell(14).setCellFormula(null);
            sheet.getRow(recordrow.getRowNum()-4).getCell(14).setCellValue(value);


            /**
             * 将本sheet的所有数据汇总到第一行 ok
             */


            sheet.getRow(0).createCell(12).setCellValue("总检测点数");
            sheet.getRow(0).createCell(13).setCellValue("总合格点数");
            sheet.getRow(0).createCell(14).setCellValue("合格率");
            sheet.getRow(0).createCell(15).setCellValue("最大SFC代表值");
            sheet.getRow(0).createCell(16).setCellValue("最小SFC代表值");

            sheet.getRow(1).createCell(12).setCellFormula("SUM("
                    +sheet.getRow(2).getCell(12).getReference()+":"
                    +sheet.getRow(sheet.getLastRowNum()).createCell(12).getReference()+")/2");//W1=SUM(W3:W825)/2
            value = e.evaluate(sheet.getRow(1).getCell(12)).getNumberValue();
            sheet.getRow(1).getCell(12).setCellFormula(null);
            sheet.getRow(1).getCell(12).setCellValue(value);

            sheet.getRow(1).createCell(13).setCellFormula("SUM("
                    +sheet.getRow(2).getCell(13).getReference()+":"
                    +sheet.getRow(sheet.getLastRowNum()).createCell(13).getReference()+")/2");//X1=SUM(X3:X825)/2
            value = e.evaluate(sheet.getRow(1).getCell(13)).getNumberValue();
            sheet.getRow(1).getCell(13).setCellFormula(null);
            sheet.getRow(1).getCell(13).setCellValue(value);

            sheet.getRow(1).createCell(14).setCellFormula(
                    sheet.getRow(1).getCell(13).getReference()+"*100/" //ok
                            +sheet.getRow(1).getCell(12).getReference());//ok
            value = e.evaluate(sheet.getRow(1).getCell(14)).getNumberValue();
            sheet.getRow(1).getCell(14).setCellFormula(null);
            sheet.getRow(1).getCell(14).setCellValue(value);


            sheet.getRow(1).createCell(15).setCellFormula("MAX("
                    +sheet.getRow(2).createCell(15).getReference()+":"
                    +sheet.getRow(sheet.getLastRowNum()).createCell(15).getReference()+")");//W1=MAX(W3:W825)/2
            value = e.evaluate(sheet.getRow(1).getCell(15)).getNumberValue();
            sheet.getRow(1).getCell(15).setCellFormula(null);
            sheet.getRow(1).getCell(15).setCellValue(value);

            sheet.getRow(1).createCell(16).setCellFormula("MIN("
                    +sheet.getRow(2).createCell(16).getReference()+":"
                    +sheet.getRow(sheet.getLastRowNum()).createCell(16).getReference()+")");//X1=MIN(X3:X825)/2
            value = e.evaluate(sheet.getRow(1).getCell(16)).getNumberValue();
            sheet.getRow(1).getCell(16).setCellFormula(null);
            sheet.getRow(1).getCell(16).setCellValue(value);
        }else {
            //计算最后一座桥或隧道
            sheet.getRow(recordrow.getRowNum()-4).createCell(3*cds+3)
                    .setCellFormula("SUM(" +recordrow.createCell(3*cds+3).getReference()+":"
                            +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(3*cds+3).getReference()+")");
            double value = e.evaluate(sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+3)).getNumberValue();
            sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+3).setCellFormula(null);
            sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+3).setCellValue(value);

            sheet.getRow(recordrow.getRowNum()-4).createCell(3*cds+4).setCellFormula("SUM("
                    +recordrow.createCell(3*cds+4).getReference()+":"
                    +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).createCell(3*cds+4).getReference()+")");
            value = e.evaluate(sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+4)).getNumberValue();
            sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+4).setCellFormula(null);
            sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+4).setCellValue(value);

            sheet.getRow(recordrow.getRowNum()-4).createCell(3*cds+5).setCellFormula(
                    sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+4).getReference()+"*100/"
                            + sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+3).getReference());
            value = e.evaluate(sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+5)).getNumberValue();
            sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+5).setCellFormula(null);
            sheet.getRow(recordrow.getRowNum()-4).getCell(3*cds+5).setCellValue(value);


            /**
             * 将本sheet的所有数据汇总到第一行 ok
             */


            sheet.getRow(0).createCell(3*cds+3).setCellValue("总检测点数");
            sheet.getRow(0).createCell(3*cds+4).setCellValue("总合格点数");
            sheet.getRow(0).createCell(3*cds+5).setCellValue("合格率");
            sheet.getRow(0).createCell(3*cds+6).setCellValue("最大SFC代表值");
            sheet.getRow(0).createCell(3*cds+7).setCellValue("最小SFC代表值");

            sheet.getRow(1).createCell(3*cds+3).setCellFormula("SUM("
                    +sheet.getRow(2).getCell(3*cds+3).getReference()+":"
                    +sheet.getRow(sheet.getLastRowNum()).createCell(3*cds+3).getReference()+")/2");//W1=SUM(W3:W825)/2
            value = e.evaluate(sheet.getRow(1).getCell(3*cds+3)).getNumberValue();
            sheet.getRow(1).getCell(3*cds+3).setCellFormula(null);
            sheet.getRow(1).getCell(3*cds+3).setCellValue(value);

            sheet.getRow(1).createCell(3*cds+4).setCellFormula("SUM("
                    +sheet.getRow(2).getCell(3*cds+4).getReference()+":"
                    +sheet.getRow(sheet.getLastRowNum()).createCell(3*cds+4).getReference()+")/2");//X1=SUM(X3:X825)/2
            value = e.evaluate(sheet.getRow(1).getCell(3*cds+4)).getNumberValue();
            sheet.getRow(1).getCell(3*cds+4).setCellFormula(null);
            sheet.getRow(1).getCell(3*cds+4).setCellValue(value);

            sheet.getRow(1).createCell(3*cds+5).setCellFormula(
                    sheet.getRow(1).getCell(3*cds+4).getReference()+"*100/" //ok
                            +sheet.getRow(1).getCell(3*cds+3).getReference());//ok
            value = e.evaluate(sheet.getRow(1).getCell(3*cds+5)).getNumberValue();
            sheet.getRow(1).getCell(3*cds+5).setCellFormula(null);
            sheet.getRow(1).getCell(3*cds+5).setCellValue(value);


            sheet.getRow(1).createCell(3*cds+6).setCellFormula("MAX("
                    +sheet.getRow(2).createCell(3*cds+6).getReference()+":"
                    +sheet.getRow(sheet.getLastRowNum()).createCell(3*cds+6).getReference()+")");//W1=MAX(W3:W825)/2
            value = e.evaluate(sheet.getRow(1).getCell(3*cds+6)).getNumberValue();
            sheet.getRow(1).getCell(3*cds+6).setCellFormula(null);
            sheet.getRow(1).getCell(3*cds+6).setCellValue(value);

            sheet.getRow(1).createCell(3*cds+7).setCellFormula("MIN("
                    +sheet.getRow(2).createCell(3*cds+7).getReference()+":"
                    +sheet.getRow(sheet.getLastRowNum()).createCell(3*cds+7).getReference()+")");//X1=MIN(X3:X825)/2
            value = e.evaluate(sheet.getRow(1).getCell(3*cds+7)).getNumberValue();
            sheet.getRow(1).getCell(3*cds+7).setCellFormula(null);
            sheet.getRow(1).getCell(3*cds+7).setCellValue(value);
        }


    }

    /**
     *  @param sheet
     * @param rowstart
     * @param rowend
     * @param e
     */
    private void calculateTotalData(XSSFSheet sheet, XSSFRow rowstart, XSSFRow rowend, FormulaEvaluator e,int cds) {
        if (cds == 5){
            //SFC平均值
            sheet.getRow(rowend.getRowNum()-5).getCell(9).setCellFormula("IF(ISERROR(AVERAGE("+rowstart.getCell(1).getReference()+":"+rowend.getCell(cds).getReference()+","
                    +rowstart.getCell(7).getReference()+":"+sheet.getRow(rowend.getRowNum()-7).getCell(11).getReference()+")),\"-\","+
                    "AVERAGE("+rowstart.getCell(1).getReference()+":"+rowend.getCell(cds).getReference()+","
                    +rowstart.getCell(7).getReference()+":"+sheet.getRow(rowend.getRowNum()-7).getCell(11).getReference()+"))");
            //标准差
            sheet.getRow(rowend.getRowNum()-4).getCell(9).setCellFormula("IF(ISERROR(ROUND(STDEV("+rowstart.getCell(1).getReference()+":"+rowend.getCell(cds).getReference()+","
                    +rowstart.getCell(7).getReference()+":"+sheet.getRow(rowend.getRowNum()-7).getCell(11).getReference()+"),3)),\"-\","+
                    "ROUND(STDEV("+rowstart.getCell(1).getReference()+":"+rowend.getCell(cds).getReference()+","
                    +rowstart.getCell(7).getReference()+":"+sheet.getRow(rowend.getRowNum()-7).getCell(11).getReference()+"),3))");

            //检测点数
            sheet.getRow(rowend.getRowNum()-1).getCell(7).setCellFormula("COUNT("+rowstart.getCell(1).getReference()+":"+rowend.getCell(cds).getReference()+","
                    +rowstart.getCell(7).getReference()+":"+sheet.getRow(rowend.getRowNum()-7).getCell(11).getReference()+")");

            //合格点数
            sheet.getRow(rowend.getRowNum()).getCell(7).setCellFormula("SUM(COUNTIF("+rowstart.getCell(1).getReference()+":"+rowend.getCell(cds).getReference()+",\">=\"&"
                    +sheet.getRow(rowend.getRowNum()-6).getCell(9).getReference()+"),COUNTIF("
                    +rowstart.getCell(7).getReference()+":"+sheet.getRow(rowend.getRowNum()-7).getCell(11).getReference()+",\">=\"&"
                    +sheet.getRow(rowend.getRowNum()-6).getCell(9).getReference()+"))");

            //SFC代表值
            sheet.getRow(rowend.getRowNum()-3).getCell(9).setCellFormula("IF(ISERROR(ROUND(IF("
                    +sheet.getRow(rowend.getRowNum()-1).getCell(7).getReference()+">100,("
                    +sheet.getRow(rowend.getRowNum()-5).getCell(9).getReference()+"-"
                    +sheet.getRow(rowend.getRowNum()-4).getCell(9).getReference()+"*(1.6449/SQRT("
                    +sheet.getRow(rowend.getRowNum()-1).getCell(7).getReference()+"))),("
                    +sheet.getRow(rowend.getRowNum()-5).getCell(9).getReference()+"-"
                    +sheet.getRow(rowend.getRowNum()-4).getCell(9).getReference()+"*VLOOKUP("
                    +sheet.getRow(rowend.getRowNum()-1).getCell(7).getReference()+",保证率系数!$A:$D,3))),2)),\"-\","+
                    "ROUND(IF("
                    +sheet.getRow(rowend.getRowNum()-1).getCell(7).getReference()+">100,("
                    +sheet.getRow(rowend.getRowNum()-5).getCell(9).getReference()+"-"
                    +sheet.getRow(rowend.getRowNum()-4).getCell(9).getReference()+"*(1.6449/SQRT("
                    +sheet.getRow(rowend.getRowNum()-1).getCell(7).getReference()+"))),("
                    +sheet.getRow(rowend.getRowNum()-5).getCell(9).getReference()+"-"
                    +sheet.getRow(rowend.getRowNum()-4).getCell(9).getReference()+"*VLOOKUP("
                    +sheet.getRow(rowend.getRowNum()-1).getCell(7).getReference()+",保证率系数!$A:$D,3))),2))");

            //结果
            sheet.getRow(rowend.getRowNum()-2).getCell(9).setCellFormula("IF("+sheet.getRow(rowend.getRowNum()-3).getCell(9).getReference()+"<>\"-\","+"IF("
                    +sheet.getRow(rowend.getRowNum()-3).getCell(9).getReference()+">="
                    +sheet.getRow(rowend.getRowNum()-6).getCell(9).getReference()+",\"合格\",\"不合格\"),\"-\")");

            //右上角统计的检测点数
            sheet.getRow(rowend.getRowNum()-4).createCell(12).setCellFormula(sheet.getRow(rowend.getRowNum()-1).getCell(7).getReference());
            double value = e.evaluate(sheet.getRow(rowend.getRowNum()-4).getCell(12)).getNumberValue();
            sheet.getRow(rowend.getRowNum()-4).getCell(12).setCellFormula(null);
            sheet.getRow(rowend.getRowNum()-4).getCell(12).setCellValue(value);

            //右上角统计的合格点数
            sheet.getRow(rowend.getRowNum()-4).createCell(13).setCellFormula(sheet.getRow(rowend.getRowNum()).getCell(7).getReference());
            value = e.evaluate(sheet.getRow(rowend.getRowNum()-4).getCell(13)).getNumberValue();
            sheet.getRow(rowend.getRowNum()-4).getCell(13).setCellFormula(null);
            sheet.getRow(rowend.getRowNum()-4).getCell(13).setCellValue(value);

            /**
             * 右上角
             * 下面计算代表值的变化范围
             */
            value = e.evaluate(sheet.getRow(rowend.getRowNum()-3).getCell(9)).getNumberValue();
            sheet.getRow(rowend.getRowNum()-5).createCell(15);
            sheet.getRow(rowend.getRowNum()-5).createCell(16);
            if(value == 0.0){
                sheet.getRow(rowend.getRowNum()-5).getCell(15).setCellValue("");
                sheet.getRow(rowend.getRowNum()-5).getCell(16).setCellValue("");
            }else{
                sheet.getRow(rowend.getRowNum()-5).getCell(15).setCellValue(value);
                sheet.getRow(rowend.getRowNum()-5).getCell(16).setCellValue(value);
            }
        }else {
            //SFC平均值
            sheet.getRow(rowend.getRowNum()-5).getCell(cds*2+3).setCellFormula("IF(ISERROR(AVERAGE("+rowstart.getCell(1).getReference()+":"+rowend.getCell(cds).getReference()+","
                    +rowstart.getCell(cds+2).getReference()+":"+sheet.getRow(rowend.getRowNum()-7).getCell(cds*2+1).getReference()+","
                    +rowstart.getCell(cds*2+3).getReference()+":"+sheet.getRow(rowend.getRowNum()-7).getCell(3*cds+2).getReference()+")),\"-\","+
                    "AVERAGE("+rowstart.getCell(1).getReference()+":"+rowend.getCell(cds).getReference()+","
                    +rowstart.getCell(cds+2).getReference()+":"+sheet.getRow(rowend.getRowNum()-7).getCell(cds*2+1).getReference()+","
                    +rowstart.getCell(cds*2+3).getReference()+":"+sheet.getRow(rowend.getRowNum()-7).getCell(3*cds+2).getReference()+"))");
            //标准差
            sheet.getRow(rowend.getRowNum()-4).getCell(cds*2+3).setCellFormula("IF(ISERROR(ROUND(STDEV("+rowstart.getCell(1).getReference()+":"+rowend.getCell(cds).getReference()+","
                    +rowstart.getCell(cds+2).getReference()+":"+sheet.getRow(rowend.getRowNum()-7).getCell(cds*2+1).getReference()+","
                    +rowstart.getCell(cds*2+3).getReference()+":"+sheet.getRow(rowend.getRowNum()-7).getCell(3*cds+2).getReference()+"),3)),\"-\","+
                    "ROUND(STDEV("+rowstart.getCell(1).getReference()+":"+rowend.getCell(cds).getReference()+","
                    +rowstart.getCell(cds+2).getReference()+":"+sheet.getRow(rowend.getRowNum()-7).getCell(cds*2+1).getReference()+","
                    +rowstart.getCell(cds*2+3).getReference()+":"+sheet.getRow(rowend.getRowNum()-7).getCell(3*cds+2).getReference()+"),3))");

            //检测点数
            sheet.getRow(rowend.getRowNum()-1).getCell(cds+3).setCellFormula("COUNT("+rowstart.getCell(1).getReference()+":"+rowend.getCell(cds).getReference()+","
                    +rowstart.getCell(cds+2).getReference()+":"+sheet.getRow(rowend.getRowNum()-7).getCell(cds*2+1).getReference()+","
                    +rowstart.getCell(cds*2+3).getReference()+":"+sheet.getRow(rowend.getRowNum()-7).getCell(3*cds+2).getReference()+")");

            //合格点数
            sheet.getRow(rowend.getRowNum()).getCell(cds+3).setCellFormula("SUM(COUNTIF("+rowstart.getCell(1).getReference()+":"+rowend.getCell(cds).getReference()+",\">=\"&"
                    +sheet.getRow(rowend.getRowNum()-6).getCell(cds*2+3).getReference()+"),COUNTIF("
                    +rowstart.getCell(cds+2).getReference()+":"+sheet.getRow(rowend.getRowNum()-7).getCell(cds*2+1).getReference()+",\">=\"&"
                    +sheet.getRow(rowend.getRowNum()-6).getCell(cds*2+3).getReference()+"),COUNTIF("
                    +rowstart.getCell(cds*2+3).getReference()+":"+sheet.getRow(rowend.getRowNum()-7).getCell(3*cds+2).getReference()+",\">=\"&"
                    +sheet.getRow(rowend.getRowNum()-6).getCell(cds*2+3).getReference()+"))");

            //SFC代表值
            sheet.getRow(rowend.getRowNum()-3).getCell(cds*2+3).setCellFormula("IF(ISERROR(ROUND(IF("
                    +sheet.getRow(rowend.getRowNum()-1).getCell(cds+3).getReference()+">100,("
                    +sheet.getRow(rowend.getRowNum()-5).getCell(cds*2+3).getReference()+"-"
                    +sheet.getRow(rowend.getRowNum()-4).getCell(cds*2+3).getReference()+"*(1.6449/SQRT("
                    +sheet.getRow(rowend.getRowNum()-1).getCell(cds+3).getReference()+"))),("
                    +sheet.getRow(rowend.getRowNum()-5).getCell(cds*2+3).getReference()+"-"
                    +sheet.getRow(rowend.getRowNum()-4).getCell(cds*2+3).getReference()+"*VLOOKUP("
                    +sheet.getRow(rowend.getRowNum()-1).getCell(cds+3).getReference()+",保证率系数!$A:$D,3))),2)),\"-\","+
                    "ROUND(IF("
                    +sheet.getRow(rowend.getRowNum()-1).getCell(cds+3).getReference()+">100,("
                    +sheet.getRow(rowend.getRowNum()-5).getCell(cds*2+3).getReference()+"-"
                    +sheet.getRow(rowend.getRowNum()-4).getCell(cds*2+3).getReference()+"*(1.6449/SQRT("
                    +sheet.getRow(rowend.getRowNum()-1).getCell(cds+3).getReference()+"))),("
                    +sheet.getRow(rowend.getRowNum()-5).getCell(cds*2+3).getReference()+"-"
                    +sheet.getRow(rowend.getRowNum()-4).getCell(cds*2+3).getReference()+"*VLOOKUP("
                    +sheet.getRow(rowend.getRowNum()-1).getCell(cds+3).getReference()+",保证率系数!$A:$D,3))),2))");

            //结果
            sheet.getRow(rowend.getRowNum()-2).getCell(cds*2+3).setCellFormula("IF("+sheet.getRow(rowend.getRowNum()-3).getCell(cds*2+3).getReference()+"<>\"-\","+"IF("
                    +sheet.getRow(rowend.getRowNum()-3).getCell(cds*2+3).getReference()+">="
                    +sheet.getRow(rowend.getRowNum()-6).getCell(cds*2+3).getReference()+",\"合格\",\"不合格\"),\"-\")");

            //右上角统计的检测点数
            sheet.getRow(rowend.getRowNum()-4).createCell(3*cds+3).setCellFormula(sheet.getRow(rowend.getRowNum()-1).getCell(cds+3).getReference());
            double value = e.evaluate(sheet.getRow(rowend.getRowNum()-4).getCell(3*cds+3)).getNumberValue();
            sheet.getRow(rowend.getRowNum()-4).getCell(3*cds+3).setCellFormula(null);
            sheet.getRow(rowend.getRowNum()-4).getCell(3*cds+3).setCellValue(value);

            //右上角统计的合格点数
            sheet.getRow(rowend.getRowNum()-4).createCell(3*cds+4).setCellFormula(sheet.getRow(rowend.getRowNum()).getCell(cds+3).getReference());
            value = e.evaluate(sheet.getRow(rowend.getRowNum()-4).getCell(3*cds+4)).getNumberValue();
            sheet.getRow(rowend.getRowNum()-4).getCell(3*cds+4).setCellFormula(null);
            sheet.getRow(rowend.getRowNum()-4).getCell(3*cds+4).setCellValue(value);

            /**
             * 右上角
             * 下面计算代表值的变化范围
             */
            value = e.evaluate(sheet.getRow(rowend.getRowNum()-3).getCell(cds*2+3)).getNumberValue();
            sheet.getRow(rowend.getRowNum()-5).createCell(3*cds+6);
            sheet.getRow(rowend.getRowNum()-5).createCell(3*cds+7);
            if(value == 0.0){
                sheet.getRow(rowend.getRowNum()-5).getCell(3*cds+6).setCellValue("");
                sheet.getRow(rowend.getRowNum()-5).getCell(3*cds+7).setCellValue("");
            }else{
                sheet.getRow(rowend.getRowNum()-5).getCell(3*cds+6).setCellValue(value);
                sheet.getRow(rowend.getRowNum()-5).getCell(3*cds+7).setCellValue(value);
            }
        }
    }

    /**
     * 根据给定的单元格起始行号及列号，得到合并单元格的最后一列列号， 如果给定的初始行号不是合并单元格，那么函数返回初始列号
     * @param sheet
     * @param cellstartrow
     * @param cellstartcol
     * @return
     */
    private int getCellEndColumn(XSSFSheet sheet, int cellstartrow, int cellstartcol) {
        int sheetmergerCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetmergerCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            if (cellstartrow >= ca.getFirstRow()
                    && cellstartrow <= ca.getLastRow()
                    && cellstartcol >= ca.getFirstColumn()
                    && cellstartcol <= ca.getLastColumn()) {
                return ca.getLastColumn();
            }
        }
        return cellstartcol;
    }

    /**
     *  @param proname
     * @param htd
     * @param data
     * @param wb
     * @param sheetname
     */
    private void DBtoExcel(String proname, String htd, List<Map<String, Object>> data, XSSFWorkbook wb, String sheetname, int cdsl,String sjz,String zx) throws ParseException {
        System.out.println(sheetname+"开始数据写入");
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        SimpleDateFormat outputDateFormat  = new SimpleDateFormat("yyyy.MM.dd");
        if (data!=null && !data.isEmpty()){
            createTable2(getNum2(data,cdsl),wb,sheetname,cdsl);
            XSSFSheet sheet = wb.getSheet(sheetname);
            String time = String.valueOf(data.get(0).get("createTime")) ;
            Date parse = simpleDateFormat.parse(time);
            String sj = outputDateFormat.format(parse);
            if (cdsl == 2){
                sheet.getRow(1).getCell(1).setCellValue(proname);
                sheet.getRow(1).getCell(7).setCellValue(htd);
                sheet.getRow(2).getCell(7).setCellValue(sj);
            }else if (cdsl == 5){
                sheet.getRow(1).getCell(1).setCellValue(proname);
                sheet.getRow(1).getCell(9).setCellValue(htd);
                sheet.getRow(2).getCell(9).setCellValue(sj);
            }else {
                sheet.getRow(1).getCell(2).setCellValue(proname);
                sheet.getRow(1).getCell(cdsl*3).setCellValue(htd);
                sheet.getRow(2).getCell(cdsl*3).setCellValue(sj);
            }
            String name = data.get(0).get("name").toString();
            int index = 0;
            int tableNum = 0;

            fillTitleCellData(sheet, tableNum, proname, htd, name,sj,sheetname,cdsl,sjz);
            if (cdsl == 5){
                for(int i =0; i < data.size(); i++){
                    if (name.equals(data.get(i).get("name"))){
                        if(index % 69 == 0 && index != 0){
                            tableNum ++;
                            fillTitleCellData(sheet, tableNum, proname, htd, name,sj,sheetname,cdsl,sjz);
                            index = 0;
                        }
                        fillFiveCommonCellData(sheet, tableNum, index, data.get(i),cdsl,zx);
                        index ++;

                    }else {
                        name = data.get(i).get("name").toString();
                        tableNum ++;
                        index = 0;
                        fillTitleCellData(sheet, tableNum, proname, htd, name,sj,sheetname,cdsl,sjz);
                        fillFiveCommonCellData(sheet, tableNum, index, data.get(i),cdsl,zx);
                        index += 1;
                    }
                }
            }else {
                for(int i =0; i < data.size(); i++){
                    if (name.equals(data.get(i).get("name"))){
                        if(index == 100){
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
        }
        System.out.println(sheetname+"数据写入完成");
    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param index
     * @param row
     */
    private void fillFiveCommonCellData(XSSFSheet sheet, int tableNum, int index, Map<String, Object> row,int cdsl,String zx) {
        int a = 0;
        if (zx.equals("主线")){
            a=1000;
        }else {
            a=1;
        }
        String[] sfc = row.get("sfc").toString().split(",");
        for (int i = 0 ; i < sfc.length ; i++) {
            if (index < 38){

                sheet.getRow(tableNum * 44 + 6 + index % 38).getCell((cdsl+1) * (index / 38)).setCellValue(Double.valueOf(Double.valueOf(row.get("qdzh").toString()) * a));
                if (sfc[i].equals("-")){
                    sheet.getRow(tableNum * 44 + 6 + index % 38).getCell((cdsl+1) * (index / 38)+1+i).setCellValue(sfc[i]);
                }else {
                    sheet.getRow(tableNum * 44 + 6 + index % 38).getCell((cdsl+1) * (index / 38)+1+i).setCellValue(Double.parseDouble(sfc[i]));
                }

            }else {
                sheet.getRow(tableNum * 44 + 6 + index % 38).getCell(6 * (index / 38)).setCellValue((Double.parseDouble(row.get("qdzh").toString())) * a);
                if (sfc[i].equals("-")){
                    sheet.getRow(tableNum * 44 + 6 + index % 38).getCell(6 * (index / 38)+1+i).setCellValue(sfc[i]);
                }else {
                    sheet.getRow(tableNum * 44 + 6 + index % 38).getCell(6 * (index / 38)+1+i).setCellValue(Double.parseDouble(sfc[i]));
                }
            }

        }

    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param index
     * @param row
     */
    private void fillCommonCellData(XSSFSheet sheet, int tableNum, int index, Map<String, Object> row,int cdsl,String zx) {
        int a = 0;
        if (zx.equals("主线")){
            a=1000;
        }else {
            a=1;
        }
        String[] sfc = row.get("sfc").toString().split(",");
        for (int i = 0 ; i < sfc.length ; i++) {
            if (index < 69){
                sheet.getRow(tableNum * 44 + 6 + index % 38).getCell((cdsl+1) * (index / 38)).setCellValue((Double.valueOf(row.get("qdzh").toString())) * a);
                if (sfc[i].equals("-")){
                    sheet.getRow(tableNum * 44 + 6 + index % 38).getCell((cdsl+1) * (index / 38)+1+i).setCellValue(sfc[i]);
                }else {
                    sheet.getRow(tableNum * 44 + 6 + index % 38).getCell((cdsl+1) * (index / 38)+1+i).setCellValue(Double.parseDouble(sfc[i]));
                }

                //sheet.getRow(tableNum * 44 + 6 + index % 38).getCell(9 * (index / 38)+5+i).setCellValue(Double.parseDouble(sfc[i]));

            }else {
                sheet.getRow(tableNum * 44 + 6 + index % 69).getCell((cdsl*2+2) * (index / 69)).setCellValue((Double.parseDouble(row.get("qdzh").toString())) * a);
                if (sfc[i].equals("-")){
                    sheet.getRow(tableNum * 44 + 6 + index % 69).getCell((cdsl*2+2) * (index / 69)+1+i).setCellValue(sfc[i]);
                }else {
                    sheet.getRow(tableNum * 44 + 6 + index % 69).getCell((cdsl*2+2) * (index / 69)+1+i).setCellValue(Double.parseDouble(sfc[i]));
                }
                //sheet.getRow(tableNum * 44 + 6 + index % 69).getCell(18 * (index / 69)+5+i).setCellValue(Double.parseDouble(sfc[i]));
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
     */
    private void fillTitleCellData(XSSFSheet sheet, int tableNum, String proname, String htd, String name,String time,String sheetname,int cdsl,String sjz) {
        String fbgcname = "";
        if (sheetname.contains("隧道")){
            fbgcname = "隧道路面";
        }else if (sheetname.contains("桥")){
            fbgcname = "桥面系";
        }else {
            fbgcname = "路面面层";
        }
        if (cdsl ==2){
            sheet.getRow(tableNum * 44 + 1).getCell(1).setCellValue(proname);
            sheet.getRow(tableNum * 44 + 2).getCell(1).setCellValue("路面工程");
            sheet.getRow(tableNum * 44 + 1).getCell(7).setCellValue(htd);
            sheet.getRow(tableNum * 44 + 2).getCell(7).setCellValue(time);
            sheet.getRow(tableNum * 44 + 1).getCell(4).setCellValue(fbgcname+"("+name+")");
            sheet.getRow(tableNum * 44 + 37).getCell(cdsl*2+3).setCellValue(Double.parseDouble(sjz));
        }else if (cdsl == 5){
            sheet.getRow(tableNum * 44 + 1).getCell(1).setCellValue(proname);
            sheet.getRow(tableNum * 44 + 2).getCell(1).setCellValue("路面工程");
            sheet.getRow(tableNum * 44 + 1).getCell(9).setCellValue(htd);
            sheet.getRow(tableNum * 44 + 2).getCell(9).setCellValue(time);
            sheet.getRow(tableNum * 44 + 1).getCell(5).setCellValue(fbgcname+"("+name+")");
            sheet.getRow(tableNum * 44 + 37).getCell(9).setCellValue(Double.parseDouble(sjz));
        }else {
            sheet.getRow(tableNum * 44 + 1).getCell(2).setCellValue(proname);
            sheet.getRow(tableNum * 44 + 2).getCell(2).setCellValue("路面工程");
            if (cdsl == 3){
                sheet.getRow(tableNum * 44 + 1).getCell(5).setCellValue(fbgcname+"("+name+")");
            }
            if (cdsl == 4){
                sheet.getRow(tableNum * 44 + 1).getCell(7).setCellValue(fbgcname+"("+name+")");
            }
            sheet.getRow(tableNum * 44 + 1).getCell(cdsl*3).setCellValue(htd);
            sheet.getRow(tableNum * 44 + 2).getCell(cdsl*3).setCellValue(time);
            sheet.getRow(tableNum * 44 + 37).getCell(cdsl*2+3).setCellValue(Double.parseDouble(sjz));
        }
    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param proname
     * @param htd
     * @param name
     * @param time
     * @param sheetname
     * @param cdsl
     * @param sjz
     */
    private void fillZdTitleCellData(XSSFSheet sheet, int tableNum, String proname, String htd, String name,String time,String sheetname,int cdsl,String sjz) {
        String fbgcname = "";
        if (sheetname.contains("隧道")){
            fbgcname = "隧道路面";
        }else if (sheetname.contains("桥")){
            fbgcname = "桥面系";
        }else {
            fbgcname = "匝道路面";
        }
        if (cdsl == 2){
            sheet.getRow(tableNum * 44 + 1).getCell(1).setCellValue(proname);
            sheet.getRow(tableNum * 44 + 2).getCell(1).setCellValue("路面工程");
            sheet.getRow(tableNum * 44 + 1).getCell(7).setCellValue(htd);
            sheet.getRow(tableNum * 44 + 2).getCell(7).setCellValue(time);
            sheet.getRow(tableNum * 44 + 1).getCell(4).setCellValue(fbgcname + "(" + name + ")");
        }else if (cdsl == 3){
            sheet.getRow(tableNum * 44 + 1).getCell(2).setCellValue(proname);
            sheet.getRow(tableNum * 44 + 2).getCell(2).setCellValue("路面工程");
            sheet.getRow(tableNum * 44 + 1).getCell(9).setCellValue(htd);
            sheet.getRow(tableNum * 44 + 2).getCell(9).setCellValue(time);
            sheet.getRow(tableNum * 44 + 1).getCell(5).setCellValue(fbgcname + "(" + name + ")");
        }else if (cdsl == 4){
            sheet.getRow(tableNum * 44 + 1).getCell(2).setCellValue(proname);
            sheet.getRow(tableNum * 44 + 2).getCell(2).setCellValue("路面工程");
            sheet.getRow(tableNum * 44 + 1).getCell(12).setCellValue(htd);
            sheet.getRow(tableNum * 44 + 2).getCell(12).setCellValue(time);
            sheet.getRow(tableNum * 44 + 1).getCell(7).setCellValue(fbgcname + "(" + name + ")");
        }
        sheet.getRow(tableNum * 44 + 37).getCell(cdsl*2+3).setCellValue(Double.parseDouble(sjz));


    }

    private int getNum2(List<Map<String, Object>> data,int cdsl) {
        int c=  0;
        if (cdsl == 5){
            c = 69;
        }else {
            c = 100;
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
            if (value%c==0){
                num += value/c;
            }else {
                num += value/c+1;
            }
        }
        return num;
    }

    /**
     * 获取table的数量
     * @param data
     * @return
     */
    private int getNum(List<Map<String, Object>> data) {
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
            if (value%100==0){
                num += value/100;
            }else {
                num += value/100+1;
            }
        }
        return num;
    }

    /**
     *
     * @param data
     * @return
     */
    private int getNum1(List<Map<String, Object>> data) {
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
            if (value%100==0){
                num += value/100;
            }else {
                num += value/100+1;
            }
        }
        return num;
    }

    /**
     *
     * @param tableNum
     * @param wb
     * @param sheetname
     */
    private void createTable(int tableNum, XSSFWorkbook wb, String sheetname,int cdsl) {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            RowCopy.copyRows(wb, sheetname, sheetname, 0, 43, i* 44);
        }
        if(record >= 1){
            wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 3*cdsl+2, 0,(record) * 44-1);//2车道
        }

    }

    private void createTable2(int tableNum, XSSFWorkbook wb, String sheetname,int cdsl) {
        if (cdsl == 5){
            int record = 0;
            record = tableNum;
            for (int i = 1; i < record; i++) {
                RowCopy.copyRows(wb, sheetname, sheetname, 0, 43, i* 44);
            }
            if(record >= 1){
                wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 11, 0,(record) * 44-1);//2车道
            }
        }else {
            int record = 0;
            record = tableNum;
            for (int i = 1; i < record; i++) {
                RowCopy.copyRows(wb, sheetname, sheetname, 0, 43, i* 44);
            }
            if(record >= 1){
                wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 3*cdsl+2, 0,(record) * 44-1);//2车道
            }
        }


    }

    /**
     *
     * @param tableNum
     * @param wb
     * @param sheetname
     * @param cdsl
     */
    private void createTableZd(int tableNum, XSSFWorkbook wb, String sheetname,int cdsl) {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            RowCopy.copyRows(wb, sheetname, sheetname, 0, 43, i* 44);
        }
        if(record >= 1){

            wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 3*cdsl+2, 0,(record) * 44-1);//2车道
        }

    }

    /**
     *
     * @param list
     * @return
     */
    private static List<Map<String, Object>> groupByZh(List<Map<String, Object>> list) {
        if (list == null || list.isEmpty()){
            return new ArrayList<>();
        }else {
            Map<String, List<String>> resultMap = new TreeMap<>();
            for (Map<String, Object> map : list) {
                String zh = map.get("qdzh").toString();
                String sfc = "";
                if (map.get("sfc") == null){
                    sfc = "-";
                }else {
                    sfc = map.get("sfc").toString();
                }
                //String sfc = map.get("sfc").toString();
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
                map.put("qdzh", entry.getKey());
                List<String> sfcList = entry.getValue();
                /*if (sfcList.size() == 1) {
                    map.put("sfc", sfcList.get(0) + ",-");
                } else {
                    map.put("sfc", String.join(",", sfcList));
                }*/
                map.put("sfc", String.join(",", entry.getValue()));
                for (Map<String, Object> item : list) {
                    if (item.get("qdzh").toString().equals(entry.getKey())) {
                        map.put("name", item.get("name"));
                        map.put("createTime", item.get("createTime"));
                        break;
                    }
                }
                resultList.add(map);
            }
            Collections.sort(resultList, new Comparator<Map<String, Object>>() {
                @Override
                public int compare(Map<String, Object> o1, Map<String, Object> o2) {
                    double name1 = Double.valueOf(o1.get("qdzh").toString());
                    double name2 = Double.valueOf(o2.get("qdzh").toString());
                    return Double.compare(name1, name2);
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
    private static List<Map<String, Object>> groupByZh1(List<Map<String, Object>> list) {
        if (list == null || list.isEmpty()){
            return new ArrayList<>();
        }else {
            Map<String, List<String>> resultMap = new TreeMap<>();
            for (Map<String, Object> map : list) {
                String zh = map.get("qdzh").toString();
                //String sfc = map.get("sfc").toString();
                String sfc = "";
                if (map.get("sfc") == null){
                    sfc = "-";
                }else {
                    sfc = map.get("sfc").toString();
                }
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
                map.put("qdzh", entry.getKey());
                map.put("sfc", String.join(",", entry.getValue()));
                for (Map<String, Object> item : list) {
                    if (item.get("qdzh").toString().equals(entry.getKey())) {
                        map.put("name", item.get("name"));
                        map.put("createTime", item.get("createTime"));
                        map.put("zdbs", item.get("zdbs"));
                        break;
                    }
                }
                resultList.add(map);
            }
            Collections.sort(resultList, new Comparator<Map<String, Object>>() {
                @Override
                public int compare(Map<String, Object> o1, Map<String, Object> o2) {
                    double name1 = Double.valueOf(o1.get("qdzh").toString());
                    double name2 = Double.valueOf(o2.get("qdzh").toString());
                    return Double.compare(name1, name2);
                }
            });

            return resultList;
        }
    }



    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        List<Map<String, Object>> mapList = new ArrayList<>();
        List<Map<String,Object>> lxlist = jjgZdhMcxsMapper.selectlx(proname,htd);

        if (lxlist.size()>0){
            for (Map<String, Object> map : lxlist) {
                String zx = map.get("lxbs").toString();
                int num = jjgZdhMcxsMapper.selectcdnum(proname,htd,zx);
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
        int pronamecl = 0;
        int htdcl = 0;
        if (cds == 2){
            pronamecl = 1;
            htdcl = 7;
        }else {
            pronamecl = 2;
            if (cds == 5){
                htdcl = 14;
            }else {
                htdcl =  cds*3;
            }
        }
        if (zx.equals("主线")){
            f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "19路面摩擦系数.xlsx");
        }else {
            f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "62互通摩擦系数-"+zx+".xlsx");
        }
        if (!f.exists()) {
            return new ArrayList<>();
        } else {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
            List<Map<String, Object>> jgmap = new ArrayList<>();
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                if (!wb.isSheetHidden(wb.getSheetIndex(wb.getSheetAt(j)))) {
                    XSSFSheet slSheet = wb.getSheetAt(j);

                    XSSFCell xmname = slSheet.getRow(1).getCell(pronamecl);//项目名
                    XSSFCell htdname = slSheet.getRow(1).getCell(htdcl);//合同段名
                    Map map = new HashMap();

                    if (proname.equals(xmname.toString()) && htd.equals(htdname.toString())) {
                        slSheet.getRow(1).getCell(3*cds+3).setCellType(CellType.STRING);//总点数
                        slSheet.getRow(1).getCell(3*cds+4).setCellType(CellType.STRING);//合格点数

                        slSheet.getRow(1).getCell(3*cds+6).setCellType(CellType.STRING);//合格点数
                        slSheet.getRow(1).getCell(3*cds+7).setCellType(CellType.STRING);//合格点数

                        double zds = Double.valueOf(slSheet.getRow(1).getCell(3*cds+3).getStringCellValue());
                        double hgds = Double.valueOf(slSheet.getRow(1).getCell(3*cds+4).getStringCellValue());
                        String zdsz1 = decf.format(zds);
                        String hgdsz1 = decf.format(hgds);
                        map.put("检测项目", zx);
                        map.put("路面类型", wb.getSheetName(j));
                        map.put("总点数", zdsz1);
                        map.put("设计值", slSheet.getRow(37).getCell(cds*2+3).getStringCellValue());
                        map.put("合格点数", hgdsz1);
                        map.put("Min", slSheet.getRow(1).getCell(3*cds+7).getStringCellValue());
                        map.put("Max", slSheet.getRow(1).getCell(3*cds+6).getStringCellValue());
                    }
                    jgmap.add(map);

                }
            }
            return jgmap;
        }
    }

    @Override
    public void exportmcxs(HttpServletResponse response, String cdsl) throws IOException {
        int cd = Integer.parseInt(cdsl);
        String fileName = "摩擦系数实测数据";
        String[][] sheetNames = {
                {"左幅一车道","左幅二车道","右幅一车道","右幅二车道"},
                {"左幅一车道","左幅二车道","左幅三车道","右幅一车道","右幅二车道","右幅三车道"},
                {"左幅一车道","左幅二车道","左幅三车道","左幅四车道","右幅一车道","右幅二车道","右幅三车道","右幅四车道"},
                {"左幅一车道","左幅二车道","左幅三车道","左幅四车道","左幅五车道","右幅一车道","右幅二车道","右幅三车道","右幅四车道","右幅五车道"}
        };
        String[] sheetName = sheetNames[cd-2];
        ExcelUtil.writeExcelMultipleSheets(response, null, fileName, sheetName, new JjgZdhMcxsVo());
    }


    @Override
    public void importmcxs(MultipartFile file, CommonInfoVo commonInfoVo) throws IOException {
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
                        .head(JjgZdhMcxsVo.class)
                        .headRowNumber(1)
                        .registerReadListener(
                                new ExcelHandler<JjgZdhMcxsVo>(JjgZdhMcxsVo.class) {
                                    @Override
                                    public void handle(List<JjgZdhMcxsVo> dataList) {
                                        for(JjgZdhMcxsVo mcxsVo: dataList)
                                        {
                                            JjgZdhMcxs mcxs = new JjgZdhMcxs();
                                            BeanUtils.copyProperties(mcxsVo,mcxs);
                                            mcxs.setCreatetime(new Date());
                                            mcxs.setProname(commonInfoVo.getProname());
                                            mcxs.setHtd(commonInfoVo.getHtd());
                                            mcxs.setQdzh(Double.parseDouble(mcxsVo.getQdzh()));
                                            if (!mcxsVo.getZdzh().isEmpty() && mcxsVo.getZdzh()!=null){
                                                mcxs.setZdzh(Double.parseDouble(mcxsVo.getZdzh()));
                                            }
                                            mcxs.setCd(sheetName);
                                            if (sheetName.contains("一")){
                                                mcxs.setVal(1);
                                            }else if (sheetName.contains("二")){
                                                mcxs.setVal(2);
                                            }else if (sheetName.contains("三")){
                                                mcxs.setVal(3);
                                            }else if (sheetName.contains("四")){
                                                mcxs.setVal(4);
                                            }else if (sheetName.contains("五")){
                                                mcxs.setVal(5);
                                            }
                                            jjgZdhMcxsMapper.insert(mcxs);
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
