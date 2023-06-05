package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.base.MultipleLists;
import glgc.jjgys.model.project.*;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.lmgc.JjgFbgcLmgcLmhpVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.*;
import glgc.jjgys.system.service.JjgFbgcLmgcLmhpService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
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
 * @author wq
 * @since 2023-04-22
 */
@Service
public class JjgFbgcLmgcLmhpServiceImpl extends ServiceImpl<JjgFbgcLmgcLmhpMapper, JjgFbgcLmgcLmhp> implements JjgFbgcLmgcLmhpService {

    @Autowired
    private JjgFbgcLmgcLmhpMapper jjgFbgcLmgcLmhpMapper;

    @Autowired
    private JjgLqsHntlmzdMapper jjgLqsHntlmzdMapper;

    @Autowired
    private JjgLqsLjxMapper jjgLqsLjxMapper;

    @Autowired
    private JjgLqsQlMapper jjgLqsQlMapper;

    @Autowired
    private JjgLqsSdMapper jjgLqsSdMapper;


    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws Exception {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //先查询需要生成几个鉴定表 根据lmlx
        List<Map<String,String>> lmlx = jjgFbgcLmgcLmhpMapper.selectlx(proname,htd,fbgc); //[{lxlx=主线}, {lxlx=岳口枢纽立交}, {lxlx=延壶路小桥连接线}, {lxlx=延长互通式立交}]
        if (lmlx.size()>0){
            for (Map<String, String> map : lmlx) {
                String lxlx = map.get("lxlx");
                DBtoExcelData(proname,htd,fbgc,lxlx);
            }
        }

    }

    /**
     * @param proname
     * @param htd
     * @param fbgc
     * @param lxlx
     */
    private void DBtoExcelData(String proname, String htd, String fbgc, String lxlx) throws Exception {
        /**
         * 为什么合同段都是路基的？
         *
         * 服务区只会匝道，桥梁和隧道都有可能有
         * 路线类型是主线，则只在桥梁和隧道清单种查找？
         * 如果路线类型不是主线，是不是分为： 互通式立交(互通) 立交(匝道) 连接线 服务区
         *      要是匝道，去桥梁和隧道清单查找是否有对应的数据，然后剩下的数据根据 混凝土路面及匝道清单 写入到路面的工作薄中，？混凝土路面及匝道清单可以不要了
         *      但是有一个左右幅的问题，匝道的实测数据是没显示左右幅，在清单中分，如果清单中左右幅的起止桩号是一样的，是无法区分的。
         *      如果有隧道的数据，写入到隧道清单吧，匝道隧道清单就不需要了。
         *
         *  路线类型是收费站：有无隧道和桥梁的情况，没有的话直接写到路面的工作薄上？
         *  路线类型是连接线：就去连接线隧道，连接线桥梁清单去查找。
         *
         *
         *  自动化数据，都只是在桥梁和隧道清单中匹配？
         */
        XSSFWorkbook wb = null;
        if (lxlx.equals("主线")){
            File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"24路面横坡.xlsx");
            File fdir = new File(filepath + File.separator + proname + File.separator + htd);
            if (!fdir.exists()) {
                //创建文件根目录
                fdir.mkdirs();
            }
            File directory = new File("service-system/src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String path = reportPath + File.separator + "横坡.xlsx";
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);

            QueryWrapper<JjgLqsSd> wrappersdzf = new QueryWrapper<>();
            wrappersdzf.like("proname",proname);
            wrappersdzf.like("lf","左幅");
            wrappersdzf.like("wz","主线");
            List<JjgLqsSd> jjgLqsSdzf = jjgLqsSdMapper.selectList(wrappersdzf);

            QueryWrapper<JjgLqsSd> wrappersdyf = new QueryWrapper<>();
            wrappersdyf.like("proname",proname);
            wrappersdyf.like("lf","右幅");
            wrappersdyf.like("wz","主线");
            List<JjgLqsSd> jjgLqsSdyf = jjgLqsSdMapper.selectList(wrappersdyf);

            QueryWrapper<JjgLqsQl> wrapperqlzf = new QueryWrapper<>();
            wrapperqlzf.like("proname",proname);
            wrapperqlzf.like("lf","左幅");
            wrapperqlzf.like("wz","主线");
            List<JjgLqsQl> jjgLqsQlzf = jjgLqsQlMapper.selectList(wrapperqlzf);

            QueryWrapper<JjgLqsQl> wrapperqlyf = new QueryWrapper<>();
            wrapperqlyf.like("proname",proname);
            wrapperqlyf.like("lf","右幅");
            wrapperqlyf.like("wz","主线");
            List<JjgLqsQl> jjgLqsQlyf = jjgLqsQlMapper.selectList(wrapperqlyf);
            /**
             * 把基础表中所以有隧道的数据查询出来，然后根据起止桩号去在横坡实测数据匹配，是的话就是哪个隧道
             */

            List<Map<String,String>> allData = new ArrayList<>();
            //隧道左幅
            List<Map<String,String>> hpsdzfdata = new ArrayList<>();
            if (jjgLqsSdzf.size()>0){
                for (JjgLqsSd jjgLqsSd : jjgLqsSdzf) {
                    hpsdzfdata.addAll(jjgFbgcLmgcLmhpMapper.selectSdZfData(proname,htd,fbgc,jjgLqsSd.getZhq(),jjgLqsSd.getZhz()));
                    allData.addAll(jjgFbgcLmgcLmhpMapper.selectSdZfData(proname,htd,fbgc,jjgLqsSd.getZhq(),jjgLqsSd.getZhz()));
                }
            }

            //隧道右幅
            List<Map<String,String>> hpsdyfdata = new ArrayList<>();
            if (jjgLqsSdyf.size()>0){
                for (JjgLqsSd jjgLqsSd : jjgLqsSdyf) {
                    hpsdyfdata.addAll(jjgFbgcLmgcLmhpMapper.selectSdYfData(proname,htd,fbgc,jjgLqsSd.getZhq(),jjgLqsSd.getZhz()));
                    allData.addAll(jjgFbgcLmgcLmhpMapper.selectSdYfData(proname,htd,fbgc,jjgLqsSd.getZhq(),jjgLqsSd.getZhz()));
                }
            }

            //桥梁左幅
            List<Map<String,String>> hpqlzfdata = new ArrayList<>();
            if (jjgLqsQlzf.size()>0){
                for (JjgLqsQl jjgLqsQl : jjgLqsQlzf) {
                    hpqlzfdata.addAll(jjgFbgcLmgcLmhpMapper.selectQlZfData(proname,htd,fbgc,jjgLqsQl.getZhq(),jjgLqsQl.getZhz()));
                    allData.addAll(jjgFbgcLmgcLmhpMapper.selectQlZfData(proname,htd,fbgc,jjgLqsQl.getZhq(),jjgLqsQl.getZhz()));
                }
            }

            //桥梁右幅
            List<Map<String,String>> hpqlyfdata = new ArrayList<>();
            if (jjgLqsQlyf.size()>0){
                for (JjgLqsQl jjgLqsQl : jjgLqsQlyf) {
                    hpqlyfdata.addAll(jjgFbgcLmgcLmhpMapper.selectQlYfData(proname,htd,fbgc,jjgLqsQl.getZhq(),jjgLqsQl.getZhz()));
                    allData.addAll(jjgFbgcLmgcLmhpMapper.selectQlYfData(proname,htd,fbgc,jjgLqsQl.getZhq(),jjgLqsQl.getZhz()));
                }
            }


            List<Map<String,String>> hpalldata = jjgFbgcLmgcLmhpMapper.selectAllList(proname,htd,fbgc);
            //路面
            List<Map<String, String>> lmdata = diff(allData, hpalldata);
            for (Map<String, String> lmdatum : lmdata) {
                lmdatum.put("name","主线");
            }
            //路面左幅
            List<Map<String,String>> lmzfdata = new ArrayList<>();
            //路面右幅
            List<Map<String,String>> lmyfdata = new ArrayList<>();

            for (Map<String, String> lmdatum : lmdata) {
                if (lmdatum.get("wz").equals("左幅")){
                    lmzfdata.add(lmdatum);
                }else if (lmdatum.get("wz").equals("右幅")){
                    lmyfdata.add(lmdatum);
                }
            }

            if (hpsdzfdata.size()>0 && hpsdzfdata.get(0).get("lmlx").contains("沥青")){
                sdqlDBtoExcel(wb,hpsdzfdata,"沥青隧道左幅");
            }else {
                sdqlDBtoExcel(wb,hpsdzfdata,"混凝土隧道左幅");
            }

            if (hpsdyfdata.size()>0 && hpsdyfdata.get(0).get("lmlx").contains("沥青")){
                sdqlDBtoExcel(wb,hpsdyfdata,"沥青隧道右幅");
            }else {
                sdqlDBtoExcel(wb,hpsdyfdata,"混凝土隧道右幅");
            }

            if (hpqlzfdata.size()>0 && hpqlzfdata.get(0).get("lmlx").contains("沥青")){
                sdqlDBtoExcel(wb,hpqlzfdata,"沥青桥面左幅");
            }else {
                sdqlDBtoExcel(wb,hpqlzfdata,"混凝土桥面左幅");
            }
            if (hpqlyfdata.size()>0 && hpqlyfdata.get(0).get("lmlx").contains("沥青")){
                sdqlDBtoExcel(wb,hpqlyfdata,"沥青桥面右幅");
            }else {
                sdqlDBtoExcel(wb,hpqlyfdata,"混凝土桥面右幅");
            }

            if (lmzfdata.size()>0 && lmzfdata.get(0).get("lmlx").contains("沥青")){
                lmDBtoExcel(wb,lmzfdata,"沥青路面左幅");
            }else {
                lmDBtoExcel(wb,lmzfdata,"混凝土路面左幅");
            }

            if (lmyfdata.size()>0 && lmyfdata.get(0).get("lmlx").contains("沥青")){
                lmDBtoExcel(wb,lmyfdata,"沥青路面右幅");
            }else {
                lmDBtoExcel(wb,lmyfdata,"混凝土路面右幅");
            }

            //设置公式
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                calculateThicknessSheet(wb.getSheetAt(j));
                JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
            }
            JjgFbgcCommonUtils.deleteEmptySheets(wb);
            FileOutputStream fileOut = new FileOutputStream(f);
            wb.write(fileOut);
            fileOut.flush();
            fileOut.close();
            out.close();
            wb.close();


        }else {
            /**
             * 根据实测数据种的路线类型，分为几大类：匝道，收费站，连接线
             * 匝道：
             *      1，根据路线类型和Z/Y去桥梁清单种查询有无数据，若有，则要返回 路幅，起止桩号，
             *      2，根据路线类型和Z/Y去隧道清单种查询有无数据，若有，则要返回 路幅，起止桩号，
             *      3，剩下的数据，需要去混凝土路面及匝道清单查询，Z/Y带上，写入到路面的工作薄中
             */
            /**
             * 现在这样，将横坡实测数据的lxlx相关的数据查询出来。
             * 不论是匝道还是连接线，一样对待处理
             */
            //在横坡实测数据中查询出对应的lxlx的数据
            QueryWrapper<JjgFbgcLmgcLmhp> wrapperhphnt = new QueryWrapper<>();
            wrapperhphnt.like("proname",proname);
            wrapperhphnt.like("lxlx",lxlx);
            List<JjgFbgcLmgcLmhp> hplxlxhnt = jjgFbgcLmgcLmhpMapper.selectList(wrapperhphnt);
            if (hplxlxhnt.size()>0){
                separateHpAllData(hplxlxhnt,lxlx);
            }

        }
    }

    /**
     * 将沥青和水泥的数据分开处理
     * @param data
     * @param lxlx
     */
    private void separateHpAllData(List<JjgFbgcLmgcLmhp> data, String lxlx) throws IOException, ParseException {
        String proname = data.get(0).getProname();
        String htd = data.get(0).getHtd();
        String fbgc = data.get(0).getFbgc();
        List<JjgFbgcLmgcLmhp> lqdata = new ArrayList<>();
        List<JjgFbgcLmgcLmhp> sndata = new ArrayList<>();
        for (JjgFbgcLmgcLmhp alldatum : data) {
            if (alldatum.getLmlx().equals("沥青混凝土")){
                lqdata.add(alldatum);
            }else {
                sndata.add(alldatum);
            }
        }
        //沥青和水泥可能同时都有
        List<MultipleLists> listslq = separateLqData(lqdata, lxlx, "沥青混凝土");
        List<MultipleLists> listssn = separateLqData(sndata, lxlx, "水泥混凝土");

        List<Map<String, String>> lqlmzfdata = new ArrayList<>();
        List<Map<String, String>> lqlmyfdata = new ArrayList<>();

        List<Map<String, String>> lqqmzfdata = new ArrayList<>();
        List<Map<String, String>> lqqmyfdata = new ArrayList<>();

        List<Map<String, String>> lqsdzfdata = new ArrayList<>();
        List<Map<String, String>> lqsdyfdata = new ArrayList<>();

        for (MultipleLists multipleLists : listslq) {
            List<Map<String, String>> listql = multipleLists.getListql();
            for (Map<String, String> qlmap : listql) {
                if (qlmap.get("wz").equals("左幅") || qlmap.get("wz").equals("单幅")){
                    lqqmzfdata.add(qlmap);
                }else if (qlmap.get("wz").equals("右幅")){
                    lqqmyfdata.add(qlmap);
                }

            }
            List<Map<String, String>> listsd = multipleLists.getListsd();
            for (Map<String, String> sdmap : listsd) {
                if (sdmap.get("wz").equals("左幅") || sdmap.get("wz").equals("单幅")){
                    lqsdzfdata.add(sdmap);
                }else if (sdmap.get("wz").equals("右幅")){
                    lqsdyfdata.add(sdmap);
                }
            }
            List<Map<String, String>> listsurplus = multipleLists.getListsurplus();
            for (Map<String, String> map : listsurplus) {
                map.put("name",map.get("lxlx"));
                if (map.get("wz").equals("左幅") || map.get("wz").equals("单幅") ){
                    lqlmzfdata.add(map);
                }else if (map.get("wz").equals("右幅")){
                    lqlmyfdata.add(map);
                }
            }
        }

        List<Map<String, String>> snlmzfdata = new ArrayList<>();
        List<Map<String, String>> snlmyfdata = new ArrayList<>();

        List<Map<String, String>> snqmzfdata = new ArrayList<>();
        List<Map<String, String>> snqmyfdata = new ArrayList<>();

        List<Map<String, String>> snsdzfdata = new ArrayList<>();
        List<Map<String, String>> snsdyfdata = new ArrayList<>();
        for (MultipleLists lists : listssn) {

            List<Map<String, String>> listql = lists.getListql();
            for (Map<String, String> qlmap : listql) {
                if (qlmap.get("wz").equals("左幅") || qlmap.get("wz").equals("单幅")){
                    snqmzfdata.add(qlmap);
                }else if (qlmap.get("wz").equals("右幅")){
                    snqmyfdata.add(qlmap);
                }

            }
            List<Map<String, String>> listsd = lists.getListsd();
            for (Map<String, String> sdmap : listsd) {
                if (sdmap.get("wz").equals("左幅") || sdmap.get("wz").equals("单幅")){
                    snsdzfdata.add(sdmap);
                }else if (sdmap.get("wz").equals("右幅")){
                    snsdyfdata.add(sdmap);
                }
            }
            List<Map<String, String>> listsurplus = lists.getListsurplus();
            for (Map<String, String> map : listsurplus) {
                map.put("name",map.get("lxlx"));
                if (map.get("wz").equals("左幅") || map.get("wz").equals("单幅") ){
                    snlmzfdata.add(map);
                }else if (map.get("wz").equals("右幅")){
                    snlmyfdata.add(map);
                }
            }

        }

        XSSFWorkbook wb = null;
        File f = new File(filepath+File.separator+proname+File.separator+htd+File.separator+"24路面横坡-"+lxlx+".xlsx");
        File fdir = new File(filepath + File.separator + proname + File.separator + htd);
        if (!fdir.exists()) {
            //创建文件根目录
            fdir.mkdirs();
        }
        File directory = new File("service-system/src/main/resources/static");
        String reportPath = directory.getCanonicalPath();
        String path = reportPath + File.separator + "横坡.xlsx";
        Files.copy(Paths.get(path), new FileOutputStream(f));
        FileInputStream out = new FileInputStream(f);
        wb = new XSSFWorkbook(out);

        writeData(wb,lqlmzfdata,"沥青路面左幅");
        writeData(wb,lqlmyfdata,"沥青路面右幅");
        writeDataQlAndSd(wb,lqqmzfdata,"沥青桥面左幅");
        writeDataQlAndSd(wb,lqqmyfdata,"沥青桥面右幅");
        writeDataQlAndSd(wb,lqsdzfdata,"沥青隧道左幅");
        writeDataQlAndSd(wb,lqsdyfdata,"沥青隧道右幅");

        writeData(wb,snlmzfdata,"混凝土路面左幅");
        writeData(wb,snlmyfdata,"混凝土路面右幅");
        writeDataQlAndSd(wb,snqmzfdata,"混凝土桥面左幅");
        writeDataQlAndSd(wb,snqmyfdata,"混凝土桥面右幅");
        writeDataQlAndSd(wb,snsdzfdata,"混凝土隧道左幅");
        writeDataQlAndSd(wb,snsdyfdata,"混凝土隧道右幅");

        for (int j = 0; j < wb.getNumberOfSheets(); j++) {
            calculateThicknessSheet(wb.getSheetAt(j));
            JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
        }
        JjgFbgcCommonUtils.deleteEmptySheets(wb);

        FileOutputStream fileOut = new FileOutputStream(f);
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
        out.close();
        wb.close();


    }

    /**
     * 写入数据
     * @param wb
     * @param data
     * @param sheetname
     */
    private void writeDataQlAndSd(XSSFWorkbook wb, List<Map<String, String>> data, String sheetname) throws ParseException {
        /**
         * 要根据name去判断，不同的话，需要写入不同的页里面
         */
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        SimpleDateFormat mysqlDateFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
        if (data.size()>0 && data != null){
            createTable(wb,getNum(data),sheetname);
            String proname = data.get(0).get("proname");
            String htd = data.get(0).get("htd");
            String fbgc = data.get(0).get("fbgc");
            XSSFSheet sheet = wb.getSheet(sheetname);

            String testtime = simpleDateFormat.format(data.get(0).get("jcsj"));

            int index = 0;
            int tableNum = 0;
            String name = data.get(0).get("name");
            fillTitleCellData(sheet,tableNum,proname,htd,fbgc,data.get(0).get("lmlx"),name,testtime);
            for(Map<String, String> row:data){
                if (name.equals(row.get("name"))){
                    //比较检测时间，拿到最新的检测时间
                    testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.get("jcsj")));
                    if(index/29 == 1){
                        sheet.getRow(tableNum*35+3).getCell(8).setCellValue(testtime);
                        testtime = simpleDateFormat.format(row.get("jcsj"));
                        tableNum ++;
                        fillTitleCellData(sheet,tableNum,proname,htd,fbgc,row.get("lmlx"),row.get("name"),testtime);
                        index = 0;
                    }
                    //填写中间下方的普通单元格
                    fillCommonCellData(sheet, tableNum, index+6, row);
                    index++;
                }else {
                    name = row.get("name");
                    testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.get("jcsj")));
                    if(index/29 == 1){
                        sheet.getRow(tableNum*35+3).getCell(8).setCellValue(testtime);
                        testtime = simpleDateFormat.format(row.get("jcsj"));
                        tableNum ++;
                        fillTitleCellData(sheet,tableNum,proname,htd,fbgc,row.get("lmlx"),row.get("name"),testtime);
                        index = 0;
                    }else {
                        index = 0;
                        tableNum ++;
                        fillTitleCellData(sheet,tableNum,proname,htd,fbgc,row.get("lmlx"),row.get("name"),testtime);
                    }
                    //填写中间下方的普通单元格
                    fillCommonCellData(sheet, tableNum, index+6, row);
                    index++;
                }

            }
            sheet.getRow(tableNum*35+3).getCell(8).setCellValue(testtime);

        }

    }

    /**
     * 获取table的数量
     * @param data
     * @return
     */
    private int getNum(List<Map<String, String>> data) {
        Map<String, Integer> resultMap = new HashMap<>();
        for (Map<String, String> map : data) {
            String name = map.get("name");
            if (resultMap.containsKey(name)) {
                resultMap.put(name, resultMap.get(name) + 1);
            } else {
                resultMap.put(name, 1);
            }
        }
        int num = 0;
        for (Map.Entry<String, Integer> entry : resultMap.entrySet()) {
            int value = entry.getValue();
            if (value%29==0){
                num += value/29;
            }else {
                num += value/29+1;
            }
        }
        return num;
    }

    /**
     * 给鉴定表中写入数据
     * @param wb
     * @param data
     * @param sheetname
     */
    private void writeData(XSSFWorkbook wb, List<Map<String, String>> data, String sheetname) throws ParseException {
        SimpleDateFormat mysqlDateFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        if (data.size()>0 && data != null){
            createTable(wb,gettableNum(data.size()),sheetname);
            String proname = data.get(0).get("proname");
            String htd = data.get(0).get("htd");
            String fbgc = data.get(0).get("fbgc");
            XSSFSheet sheet = wb.getSheet(sheetname);

            String sj = data.get(0).get("jcsj");
            Date parse = mysqlDateFormat.parse(sj);
            String testtime = simpleDateFormat.format(parse);

            int index = 0;
            int tableNum = 0;
            String name = data.get(0).get("name");
            fillTitleCellData(sheet,tableNum,proname,htd,fbgc,data.get(0).get("lmlx"),name,testtime);
            for(Map<String, String> row:data){
                //比较检测时间，拿到最新的检测时间
                testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(mysqlDateFormat.parse(row.get("jcsj"))));
                if(index/29 == 1){
                    sheet.getRow(tableNum*35+3).getCell(8).setCellValue(testtime);
                    testtime = row.get("jcsj");
                    tableNum ++;
                    fillTitleCellData(sheet,tableNum,proname,htd,fbgc,row.get("lmlx"),row.get("name"),testtime);
                    index = 0;
                }
                //填写中间下方的普通单元格
                fillCommonCellData(sheet, tableNum, index+6, row);
                index++;
            }
            sheet.getRow(tableNum*35+3).getCell(8).setCellValue(testtime);

        }

    }


    /**
     *
     * @param size
     * @return
     */
    private int gettableNum(int size) {
        return size%29 ==0 ? size/29 : size/29+1;
    }

    /**
     * 处理沥青的数据
     * @param data
     * @param lxlx
     * @param lq
     */
    private List<MultipleLists> separateLqData(List<JjgFbgcLmgcLmhp> data, String lxlx, String lq) {
        List<JjgFbgcLmgcLmhp> zfdata = new ArrayList<>();
        List<JjgFbgcLmgcLmhp> yfdata = new ArrayList<>();
        List<JjgFbgcLmgcLmhp> dfdata = new ArrayList<>();

        for (JjgFbgcLmgcLmhp stringMap : data) {
            if (stringMap.getWz().equals("左幅")){
                zfdata.add(stringMap);
            }else if (stringMap.getWz().equals("右幅")){
                yfdata.add(stringMap);
            }else if (stringMap.getWz().equals("单幅")){
                dfdata.add(stringMap);
            }
        }
        /**
         * 横坡中全部的数据
         */
        List<MultipleLists> lists = new ArrayList<>();
        if (dfdata.size()>0){
            MultipleLists lqdfData = handleLqdfData(dfdata, lxlx, "单幅", lq);
            lists.add(lqdfData);
        }
        if (zfdata.size()>0){
            MultipleLists lqzfData = handleLqdfData(zfdata, lxlx, "左幅", lq);
            lists.add(lqzfData);
        }
        if (yfdata.size()>0){
            MultipleLists lqyfData = handleLqdfData(yfdata, lxlx, "右幅", lq);
            lists.add(lqyfData);
        }
        return lists;
    }

    /**
     * 分别处理单幅，左幅，右幅的数据
     * @param data
     * @param lxlx
     * @param df
     * @param lq
     */
    private MultipleLists handleLqdfData(List<JjgFbgcLmgcLmhp> data, String lxlx, String df, String lq) {
        List<Map<String, String>> queryqldata = new ArrayList<>();
        List<Map<String, String>> querysddata = new ArrayList<>();
        List<Map<String, String>> diff = new ArrayList<>();
        if (lxlx.contains("连接线")){
            /*List<Map<String, String>> queryqldata = new ArrayList<>();
            List<Map<String, String>> querysddata = new ArrayList<>();*/
            QueryWrapper<JjgLjx> wrapperljx = new QueryWrapper<>();
            wrapperljx.like("proname",data.get(0).getProname());
            wrapperljx.like("name",lxlx);
            wrapperljx.like("lf",df);
            List<JjgLjx> ljxdata = jjgLqsLjxMapper.selectList(wrapperljx);
            if (ljxdata.size()>0){
                for (JjgLjx jjgLjx : ljxdata) {
                    Map<String, String> map = new HashMap<>();
                    map.put("proname", data.get(0).getProname());
                    map.put("htd", data.get(0).getHtd());
                    map.put("fbgc", data.get(0).getFbgc());
                    map.put("zhq", jjgLjx.getZhq());
                    map.put("zhz", jjgLjx.getZhz());
                    map.put("bz", jjgLjx.getBz());
                    map.put("df", df);
                    map.put("lxlx", lxlx);
                    map.put("lq", lq);
                    List<JjgFbgcLmgcLmhp> selecthpljx = jjgFbgcLmgcLmhpMapper.selecthpljx(map);//根据起始桩号，查询连接线的数据
                    if (selecthpljx.size()>0){
                        QueryWrapper<JjgLqsQl> wrapperql = new QueryWrapper<>();
                        wrapperql.like("proname", data.get(0).getProname());
                        wrapperql.like("wz", lxlx);
                        wrapperql.like("lf", df);
                        wrapperql.like("bz", jjgLjx.getBz());
                        wrapperql.like("pzlx", lq);
                        List<JjgLqsQl> qlzfzhdata = jjgLqsQlMapper.selectList(wrapperql);//查询桥梁是否有无连接线
                        if (qlzfzhdata.size() > 0) {
                            Map<String, String> mapljx = new HashMap<>();
                            mapljx.put("proname", data.get(0).getProname());
                            mapljx.put("htd", data.get(0).getHtd());
                            mapljx.put("fbgc", data.get(0).getFbgc());
                            mapljx.put("zhq", qlzfzhdata.get(0).getZhq());
                            mapljx.put("zhz", qlzfzhdata.get(0).getZhz());
                            mapljx.put("zy", qlzfzhdata.get(0).getBz());
                            mapljx.put("wz", df);
                            mapljx.put("lxlx", lxlx);
                            mapljx.put("lq", lq);
                            queryqldata.addAll(jjgFbgcLmgcLmhpMapper.selecthpzdData(mapljx));
                        }

                        QueryWrapper<JjgLqsSd> wrapperljxql = new QueryWrapper<>();
                        wrapperljxql.like("proname", data.get(0).getProname());
                        wrapperljxql.like("wz", lxlx);
                        wrapperljxql.like("lf", df);
                        wrapperljxql.like("zdbz", jjgLjx.getBz());
                        wrapperljxql.like("pzlx", lq);
                        List<JjgLqsSd> sdzfzhdata = jjgLqsSdMapper.selectList(wrapperljxql);//正常的逻辑，这块只有一条数据，先这样
                        if (sdzfzhdata.size() > 0) {
                            /**
                             * 然后就拿着该匝道的桥梁起始桩号，去横坡表中查询哪些是属于桥梁的数据
                             */
                            Map<String, String> mapzdql = new HashMap<>();
                            mapzdql.put("proname", data.get(0).getProname());
                            mapzdql.put("htd", data.get(0).getHtd());
                            mapzdql.put("fbgc", data.get(0).getFbgc());
                            mapzdql.put("zhq", sdzfzhdata.get(0).getZhq());
                            mapzdql.put("zhz", sdzfzhdata.get(0).getZhz());
                            mapzdql.put("zy", sdzfzhdata.get(0).getZdbz());
                            mapzdql.put("wz", df);
                            mapzdql.put("lxlx", lxlx);
                            mapzdql.put("lq", lq);
                            //需要吧桥梁查询的数据加上
                            querysddata.addAll(jjgFbgcLmgcLmhpMapper.selecthpzdsdData(mapzdql));
                        }
                    }
                }
            }
            List<Map<String, String>> maps = convertListToMapList(data);
            List<Map<String, String>> diffql = diff(maps, queryqldata);
            diff.addAll(diff(diffql, querysddata));
            return new MultipleLists(queryqldata,querysddata,diff);
        }else {
            /**
             * 去混凝土路面及匝道清单中查询起始桩号，然后再去桥梁或者隧道清单查找有无桥梁和隧道
             * data单幅的全部数据,那么data处理后，剩下的数据就归路面了
             */
            /*List<Map<String, String>> queryqldata = new ArrayList<>();
            List<Map<String, String>> querysddata = new ArrayList<>();*/

            QueryWrapper<JjgLqsHntlmzd> wrapperzd = new QueryWrapper<>();
            wrapperzd.like("proname",data.get(0).getProname());
            wrapperzd.like("wz",lxlx);
            wrapperzd.like("lf",df);
            wrapperzd.like("pzlx",lq);
            List<JjgLqsHntlmzd> sdzhdata = jjgLqsHntlmzdMapper.selectList(wrapperzd);
            /**
             * 至此，可以获取到相关匝道的所有数据。
             */
            if (sdzhdata.size() > 0) {
                for (JjgLqsHntlmzd hntlmzd : sdzhdata) {
                    Map<String, String> map = new HashMap<>();
                    map.put("proname", data.get(0).getProname());
                    map.put("htd", data.get(0).getHtd());
                    map.put("fbgc", data.get(0).getFbgc());
                    map.put("zhq", hntlmzd.getZhq());
                    map.put("zhz", hntlmzd.getZhz());
                    map.put("zdlx", hntlmzd.getZdlx());
                    map.put("df", df);
                    map.put("lxlx", lxlx);
                    map.put("lq", lq);
                    List<JjgFbgcLmgcLmhp> writeData = jjgFbgcLmgcLmhpMapper.selecthpzd(map);
                    /**
                     * 这些数据全是符合清单中起止桩号的数据，所以还要进一步对比桥梁和清单中的数据。看是否存在桥梁和清单
                     */
                    if (writeData.size() > 0) {
                        /**
                         * 然后就是将这些数据分别和桥梁隧道清单查找，看是否有无桥梁隧道。
                         * 肯定还会有剩下的数据，需单独存储。
                         */
                        QueryWrapper<JjgLqsQl> wrapperql = new QueryWrapper<>();
                        wrapperql.like("proname", data.get(0).getProname());
                        wrapperql.like("wz", lxlx);
                        wrapperql.like("lf", df);
                        wrapperql.like("pzlx", lq);
                        wrapperql.like("bz", hntlmzd.getZdlx());
                        List<JjgLqsQl> qlzfzhdata = jjgLqsQlMapper.selectList(wrapperql);//正常的逻辑，这块只有一条数据，先这样
                        if (qlzfzhdata.size() > 0) {
                            /**
                             * 然后就拿着该匝道的桥梁起始桩号，去横坡表中查询哪些是属于桥梁的数据
                             */
                            Map<String, String> mapzd = new HashMap<>();
                            mapzd.put("proname", data.get(0).getProname());
                            mapzd.put("htd", data.get(0).getHtd());
                            mapzd.put("fbgc", data.get(0).getFbgc());
                            mapzd.put("zhq", qlzfzhdata.get(0).getZhq());
                            mapzd.put("zhz", qlzfzhdata.get(0).getZhz());
                            mapzd.put("zy", qlzfzhdata.get(0).getBz());
                            mapzd.put("wz", df);
                            mapzd.put("lxlx", lxlx);
                            mapzd.put("lq", lq);
                            //需要吧桥梁查询的数据加上
                            queryqldata.addAll(jjgFbgcLmgcLmhpMapper.selecthpzdData(mapzd));
                        }

                        //隧道
                        QueryWrapper<JjgLqsSd> wrappersd = new QueryWrapper<>();
                        wrappersd.like("proname", data.get(0).getProname());
                        wrappersd.like("wz", lxlx);
                        wrappersd.like("lf", df);
                        wrappersd.like("pzlx", lq);
                        wrappersd.like("zdbz", hntlmzd.getZdlx());
                        List<JjgLqsSd> sdzfzhdata = jjgLqsSdMapper.selectList(wrappersd);//正常的逻辑，这块只有一条数据，先这样
                        if (sdzfzhdata.size() > 0) {
                            /**
                             * 然后就拿着该匝道的桥梁起始桩号，去横坡表中查询哪些是属于桥梁的数据
                             */
                            Map<String, String> mapzdql = new HashMap<>();
                            mapzdql.put("proname", data.get(0).getProname());
                            mapzdql.put("htd", data.get(0).getHtd());
                            mapzdql.put("fbgc", data.get(0).getFbgc());
                            mapzdql.put("zhq", sdzfzhdata.get(0).getZhq());
                            mapzdql.put("zhz", sdzfzhdata.get(0).getZhz());
                            mapzdql.put("zy", sdzfzhdata.get(0).getZdbz());
                            mapzdql.put("wz", df);
                            mapzdql.put("lx", lq);
                            mapzdql.put("lxlx", lxlx);
                            //需要吧桥梁查询的数据加上
                            querysddata.addAll(jjgFbgcLmgcLmhpMapper.selecthpzdsdData(mapzdql));
                        }
                    }


                }
            }

            /**
             * 在这，将所有的数据和桥梁隧道对比
             */
            List<Map<String, String>> maps = convertListToMapList(data);
            List<Map<String, String>> diffql = diff(maps, queryqldata);
            diff.addAll(diff(diffql, querysddata));
            System.out.println(maps);
            System.out.println(queryqldata);
            System.out.println(diff(diffql, querysddata));
            return new MultipleLists(queryqldata,querysddata,diff);
        }
    }

    /**
     *List<JjgFbgcLmgcLmhp>转换成List<Map<String, String>>的方法
     * @param list
     * @param <T>
     * @return
     */
    private static <T> List<Map<String, String>> convertListToMapList(List<T> list) {
        List<Map<String, String>> result = new ArrayList<>();
        for (T item : list) {
            Map<String, String> map = new HashMap<>();
            Field[] fields = item.getClass().getDeclaredFields();
            for (Field field : fields) {
                field.setAccessible(true);
                try {
                    Object value = field.get(item);
                    map.put(field.getName(), value != null ? value.toString() : "");
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
            result.add(map);
        }
        return result;
    }


    /**
     *实体类转换成Map<String, String>
     * @param obj
     * @return
     * @throws Exception
     */
    private static Map<String, String> objectToMap(Object obj) throws Exception {
        if (obj == null) {
            return null;
        }
        Map<String, String> map = new HashMap<String, String>();

        Field[] declaredFields = obj.getClass().getDeclaredFields();
        for (Field field : declaredFields) {
            field.setAccessible(true);
            map.put(field.getName(), (String) field.get(obj));
        }

        return map;
    }

    /**
     *
     * @param sheet
     */
    private void calculateThicknessSheet(XSSFSheet sheet) {
        XSSFRow row = null;
        boolean flag = false;
        String name = "";
        ArrayList<XSSFRow> startRowList = new ArrayList<XSSFRow>();
        ArrayList<XSSFRow> endRowList = new ArrayList<XSSFRow>();
        sheet.getRow(0).getCell(12).setCellValue("合计");
        sheet.getRow(0).getCell(13).setCellValue("总点数");
        sheet.getRow(0).getCell(14).setCellFormula("SUM("
                +sheet.getRow(5).getCell(14).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).getCell(14).getReference()+")");
        sheet.getRow(0).getCell(15).setCellValue("合格点数");
        sheet.getRow(0).getCell(16).setCellFormula("SUM("
                +sheet.getRow(5).getCell(16).getReference()+":"
                +sheet.getRow(sheet.getPhysicalNumberOfRows()-1).getCell(16).getReference()+")");
        sheet.getRow(0).getCell(17).setCellValue("合格率");
        sheet.getRow(0).getCell(18).setCellFormula(
                sheet.getRow(0).getCell(16).getReference()+"/"
                        +sheet.getRow(0).getCell(14).getReference()+"*100");
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            if (flag && !"".equals(row.getCell(0).toString()) && row.getCell(0).toString().contains("质量鉴定表")) {
                flag = false;
            }
            if(flag && !"".equals(row.getCell(2).toString())){
                row.getCell(5).setCellFormula(row.getCell(4).getReference()+"-"
                        +row.getCell(3).getReference());//F=E7-D7
                row.getCell(7).setCellFormula(row.getCell(5).getReference()+"/"
                        +row.getCell(6).getReference()+"*100");//H=F7/G7*100
                row.getCell(10).setCellFormula("IF((("
                        +row.getCell(7).getReference()+"+"
                        +row.getCell(9).getReference()+")>="
                        +row.getCell(8).getReference()+")*AND("
                        +row.getCell(8).getReference()+">=("
                        +row.getCell(7).getReference()+"-"
                        +row.getCell(9).getReference()+")),\"√\",\"\")");//K=IF(((H7+0.3)>=I7)*AND(I7>=(H7-0.3)),"√","")
                row.getCell(11).setCellFormula("IF((("
                        +row.getCell(7).getReference()+"+"
                        +row.getCell(9).getReference()+")>="
                        +row.getCell(8).getReference()+")*AND("
                        +row.getCell(8).getReference()+">=("
                        +row.getCell(7).getReference()+"-"
                        +row.getCell(9).getReference()+")),\"\",\"×\")");//L=IF(((H7+0.3)>=I7)*AND(I7>=(H7-0.3)),"","×")
                row.getCell(12).setCellFormula(row.getCell(7).getReference()+"-"
                        +row.getCell(8).getReference());
            }
            if ("桩  号".equals(row.getCell(0).toString())) {
                i += 1;
                flag = true;
                if(!name.equals(sheet.getRow(i-3).getCell(8).toString()) && !"".equals(name)){
                    String column14 = "";
                    String column16 = "";
                    for(int r = 0;r < startRowList.size(); r++){
                        column14 += startRowList.get(r).getCell(7).getReference()+":"+endRowList.get(r).getCell(7).getReference()+",";
                        column16 += "COUNTIF("+startRowList.get(r).getCell(10).getReference()+":"+endRowList.get(r).getCell(10).getReference()+",\"√\")+";
                    }
                    startRowList.get(0).createCell(14).setCellFormula("COUNT("+column14.substring(0, column14.lastIndexOf(','))+")");
                    startRowList.get(0).createCell(16).setCellFormula(column16.substring(0, column16.lastIndexOf('+')));//=COUNTIF(AG6:AG50,"√")
                    startRowList.get(0).createCell(18).setCellFormula(startRowList.get(0).getCell(16).getReference()+"*100/"
                            +startRowList.get(0).getCell(14).getReference());//=COUNTIF(AH6:AH50,"×")

                    startRowList.clear();
                    endRowList.clear();
                }
                startRowList.add(sheet.getRow(i+1));
                endRowList.add(sheet.getRow(i+29));
                name = sheet.getRow(i-3).getCell(8).toString();
                //System.out.println("name = "+name);
            }
        }
        String column14 = "";
        String column16 = "";
        for(int r = 0;r < startRowList.size(); r++){
            column14 += startRowList.get(r).getCell(7).getReference()+":"+endRowList.get(r).getCell(7).getReference()+",";
            column16 += "COUNTIF("+startRowList.get(r).getCell(10).getReference()+":"+endRowList.get(r).getCell(10).getReference()+",\"√\")+";
        }
        startRowList.get(0).createCell(14).setCellFormula("COUNT("+column14.substring(0, column14.lastIndexOf(','))+")");
        startRowList.get(0).createCell(16).setCellFormula(column16.substring(0, column16.lastIndexOf('+')));//=COUNTIF(AG6:AG50,"√")
        startRowList.get(0).createCell(18).setCellFormula(startRowList.get(0).getCell(16).getReference()+"*100/"
                +startRowList.get(0).getCell(14).getReference());//=COUNTIF(AH6:AH50,"×")

        startRowList.clear();
        endRowList.clear();
    }

    /**
     *
     * @param wb
     * @param data
     * @param sheetname
     */
    private void lmDBtoExcel(XSSFWorkbook wb, List<Map<String, String>> data, String sheetname) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        if (data.size()>0){
            createTable(wb,gettableNum(data.size()),sheetname);
            String proname = data.get(0).get("proname");
            String htd = data.get(0).get("htd");
            String fbgc = data.get(0).get("fbgc");
            XSSFSheet sheet = wb.getSheet(sheetname);
            //Date jcsj = data.get(0).get("jcsj");
            String testtime = simpleDateFormat.format(data.get(0).get("jcsj"));
            int index = 0;
            int tableNum = 0;
            fillTitleCellData(sheet,tableNum,proname,htd,fbgc,data.get(0).get("lmlx"),"",testtime);
            for(Map<String, String> row:data){
                //比较检测时间，拿到最新的检测时间
                testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.get("jcsj")));
                if(index/29 == 1){
                    sheet.getRow(tableNum*35+3).getCell(8).setCellValue(testtime);
                    testtime = simpleDateFormat.format(row.get("jcsj"));
                    tableNum ++;
                    fillTitleCellData(sheet,tableNum,proname,htd,fbgc,row.get("lmlx"),"",testtime);
                    index = 0;
                }
                //填写中间下方的普通单元格
                fillCommonCellData(sheet, tableNum, index+6, row);
                index++;
            }
            sheet.getRow(tableNum*35+3).getCell(8).setCellValue(testtime);
        }
    }

    /**
     *
     * @param wb
     * @param data
     * @param sheetname
     */
    private void sdqlDBtoExcel(XSSFWorkbook wb, List<Map<String, String>> data, String sheetname) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        if (data.size()>0){
            createTable(wb,getNum(data),sheetname);
            //createTable(wb,gettableNum(data.size()),sheetname);
            String proname = data.get(0).get("proname");
            String htd = data.get(0).get("htd");
            String fbgc = data.get(0).get("fbgc");
            XSSFSheet sheet = wb.getSheet(sheetname);
            //Date jcsj = (Date) data.get(0).get("jcsj");
            String testtime = simpleDateFormat.format(data.get(0).get("jcsj"));
            int index = 0;
            int tableNum = 0;
            String name = data.get(0).get("name");
            fillTitleCellData(sheet,tableNum,proname,htd,fbgc,data.get(0).get("lmlx"),data.get(0).get("name"),testtime);
            for(Map<String, String> row:data){
                if (name.equals(row.get("name"))){
                    //比较检测时间，拿到最新的检测时间
                    testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.get("jcsj")));
                    if(index/29 == 1){
                        sheet.getRow(tableNum*35+3).getCell(8).setCellValue(testtime);
                        testtime = simpleDateFormat.format(row.get("jcsj"));
                        tableNum ++;
                        System.out.println(row.get("name")+"wqwq"+row);
                        fillTitleCellData(sheet,tableNum,proname,htd,fbgc,row.get("lmlx"),row.get("name"),testtime);
                        index = 0;
                    }
                    //填写中间下方的普通单元格
                    fillCommonCellData(sheet, tableNum, index+6, row);
                    index++;
                }else {
                    name = row.get("name");
                    testtime = JjgFbgcCommonUtils.getLastDate(testtime, simpleDateFormat.format(row.get("jcsj")));
                    if(index/29 == 1){
                        sheet.getRow(tableNum*35+3).getCell(8).setCellValue(testtime);
                        testtime = simpleDateFormat.format(row.get("jcsj"));
                        tableNum ++;
                        fillTitleCellData(sheet,tableNum,proname,htd,fbgc,row.get("lmlx"),row.get("name"),testtime);
                        index = 0;
                    }else {
                        index = 0;
                        tableNum++;
                        fillTitleCellData(sheet,tableNum,proname,htd,fbgc,row.get("lmlx"),row.get("name"),testtime);
                    }
                    //填写中间下方的普通单元格
                    fillCommonCellData(sheet, tableNum, index+6, row);
                    index++;

                }
            }
            sheet.getRow(tableNum*35+3).getCell(8).setCellValue(testtime);
        }

    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param index
     * @param row
     */
    private void fillCommonCellData(XSSFSheet sheet, int tableNum, int index, Map<String, String> row) {
        sheet.getRow(tableNum*35+index).getCell(0).setCellValue(getPile(row.get("zy"),String.valueOf(row.get("zh"))));
        sheet.getRow(tableNum*35+index).getCell(2).setCellValue(row.get("wz"));
        sheet.getRow(tableNum*35+index).getCell(3).setCellValue(Double.parseDouble(row.get("qsds")));
        sheet.getRow(tableNum*35+index).getCell(4).setCellValue(Double.parseDouble(row.get("hsds")));
        sheet.getRow(tableNum*35+index).getCell(6).setCellValue(Double.parseDouble(row.get("length")));
        sheet.getRow(tableNum*35+index).getCell(8).setCellValue(Double.parseDouble(row.get("sjz")));
        sheet.getRow(tableNum*35+index).getCell(9).setCellValue(Double.parseDouble(row.get("yxps")));
    }

    /**
     *
     * @param sheet
     * @param tableNum
     * @param proname
     * @param htd
     * @param fbgc
     * @param lmlx
     * @param name
     */
    private void fillTitleCellData(XSSFSheet sheet, int tableNum, String proname, String htd, String fbgc, String lmlx,String name,String testtime) {
        sheet.getRow(tableNum*35+1).getCell(2).setCellValue(proname);
        sheet.getRow(tableNum*35+1).getCell(8).setCellValue(htd);
        sheet.getRow(tableNum*35+2).getCell(2).setCellValue(fbgc);
        String mname = "";
        if (name.contains("隧道")){
            mname="隧道路面";
        }else if (name.contains("桥")){
            mname="桥面系";
        }else {
            mname="路面面层";
            name = "主线";
        }
        sheet.getRow(tableNum*35+2).getCell(8).setCellValue(mname+"("+name+")");
        sheet.getRow(tableNum*35+3).getCell(2).setCellValue(lmlx);
        sheet.getRow(tableNum*35+3).getCell(8).setCellValue(testtime);

    }


    /**
     * 取出两个list的不同
     * @param list1
     * @param list2
     * @return
     */
    private List<Map<String, String>> diff(List<Map<String, String>> list1, List<Map<String, String>> list2) {
        // 合并两个List
        List<Map<String, String>> mergeList = new ArrayList<>();
        mergeList.addAll(list1);
        mergeList.addAll(list2);

        // 按照id分组
        Map<String, List<Map<String, String>>> result =
                mergeList.stream().collect(Collectors.groupingBy(m -> m.get("id")));

        // 过滤相同的数据
        List<Map<String, String>> diffList = result.values()
                .stream()
                .filter(list -> list.size() == 1)
                .flatMap(Collection::stream)
                .collect(Collectors.toList());

        return diffList;
    }


    /**
     * 得到桩号
     * @param ZY
     * @param startpileNO
     * @return
     */
    private String getPile(String ZY, String startpileNO) {
        if(startpileNO.endsWith(".0")){
            startpileNO = startpileNO.substring(0, startpileNO.lastIndexOf(".0"));
        }
        String pileNO= "";
        String dot = "";
        if(startpileNO.contains(".")){
            String temp[] = startpileNO.split("[.]");
            startpileNO =temp[0];
            dot = "."+temp[1];
        }
        if(startpileNO.length() <= 3){
            for (int i = 0; i < 3-startpileNO.length(); i++) {
                startpileNO = "0" + startpileNO;
            }
            pileNO += ZY + "K0+" + startpileNO+ dot;
        }
        else{
            pileNO += ZY + "K" + startpileNO.substring(0,startpileNO.length()-3) + "+"
                    + startpileNO.substring(startpileNO.length()-3,startpileNO.length())+ dot;
        }
        return pileNO;
    }


    /**
     *
     * @param tableNum
     * @param wb
     * @param sheetname
     */
    private void createTable(XSSFWorkbook wb,int tableNum,String sheetname) {
        int record = 0;
        record = tableNum;
        for (int i = 1; i < record; i++) {
            RowCopy.copyRows(wb, sheetname, sheetname, 0, 34, i* 35);
        }
        if(record >= 1){
            wb.setPrintArea(wb.getSheetIndex(sheetname), 0, 11, 0,(record) * 35-1);
        }
    }


    @Override
    public List<Map<String, String>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        List<Map<String, String>> mapList = new ArrayList<>();
        List<Map<String,String>> mclist = jjgFbgcLmgcLmhpMapper.selectlx(proname,htd,fbgc);
        if (mclist.size()>0){
            for (Map<String, String> m : mclist) {
                for (String k : m.keySet()){
                    String mc = m.get(k);
                    List<Map<String, String>> looksdjdb = looksdjdb(proname, htd, mc);
                    mapList.addAll(looksdjdb);
                }
            }
            return mapList;
        }else {
            return null;
        }
    }

    /**
     *
     * @param proname
     * @param htd
     * @param mc
     * @return
     * @throws IOException
     */
    private List<Map<String, String>> looksdjdb(String proname, String htd, String mc) throws IOException {
        DecimalFormat df = new DecimalFormat(".00");
        DecimalFormat decf = new DecimalFormat("0.##");
        File f;
        if (mc.equals("主线")){
            f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "24路面横坡.xlsx");
        }else {
            f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "24路面横坡-"+mc+".xlsx");
        }

        if (!f.exists()) {
            return null;
        } else {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(f));
            List<Map<String, String> > jgmap = new ArrayList<>();
            for (int j = 0; j < wb.getNumberOfSheets(); j++) {
                if (!wb.isSheetHidden(wb.getSheetIndex(wb.getSheetAt(j)))) {
                    XSSFSheet slSheet = wb.getSheetAt(j);
                    XSSFCell xmname = slSheet.getRow(1).getCell(2);//项目名
                    XSSFCell htdname = slSheet.getRow(1).getCell(8);//合同段名
                    Map map = new HashMap();

                    if (proname.equals(xmname.toString()) && htd.equals(htdname.toString())) {
                        slSheet.getRow(0).getCell(14).setCellType(CellType.STRING);//总点数
                        slSheet.getRow(0).getCell(16).setCellType(CellType.STRING);//合格点数
                        slSheet.getRow(0).getCell(18).setCellType(CellType.STRING);//合格率
                        double zds = Double.valueOf(slSheet.getRow(0).getCell(14).getStringCellValue());
                        double hgds = Double.valueOf(slSheet.getRow(0).getCell(16).getStringCellValue());
                        double hgl = Double.valueOf(slSheet.getRow(0).getCell(18).getStringCellValue());
                        String zdsz1 = decf.format(zds);
                        String hgdsz1 = decf.format(hgds);
                        String hglz1 = df.format(hgl);
                        map.put("检测项目", mc);
                        map.put("路面类型", wb.getSheetName(j));
                        map.put("检测点数", zdsz1);
                        map.put("合格点数", hgdsz1);
                        map.put("合格率", hglz1);
                    }
                    jgmap.add(map);

                }
            }
            return jgmap;
        }
    }

    @Override
    public void exportlmhp(HttpServletResponse response) {
        String fileName = "10路面横坡实测数据";
        String sheetName = "实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcLmgcLmhpVo()).finish();

    }

    @Override
    public void importlmhp(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcLmgcLmhpVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcLmgcLmhpVo>(JjgFbgcLmgcLmhpVo.class) {
                                @Override
                                public void handle(List<JjgFbgcLmgcLmhpVo> dataList) {
                                    for(JjgFbgcLmgcLmhpVo lmgcLmhpVo: dataList)
                                    {
                                        JjgFbgcLmgcLmhp lmhp = new JjgFbgcLmgcLmhp();
                                        BeanUtils.copyProperties(lmgcLmhpVo,lmhp);
                                        double zh = Double.parseDouble(lmgcLmhpVo.getZh().replaceAll("[^0-9]", ""));
                                        lmhp.setZh(zh);
                                        lmhp.setCreatetime(new Date());
                                        lmhp.setProname(commonInfoVo.getProname());
                                        lmhp.setHtd(commonInfoVo.getHtd());
                                        lmhp.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcLmgcLmhpMapper.insert(lmhp);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }

    @Override
    public List<Map<String, String>> selectmc(String proname, String htd, String fbgc) {
        List<Map<String,String>> sdmclist = jjgFbgcLmgcLmhpMapper.selectlx(proname,htd,fbgc);
        return sdmclist;
    }
}
