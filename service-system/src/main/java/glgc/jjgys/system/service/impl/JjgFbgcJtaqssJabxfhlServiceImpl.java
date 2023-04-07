package glgc.jjgys.system.service.impl;

import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabxfhl;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabz;
import glgc.jjgys.model.project.JjgFbgcLjgcLjtsfysdHt;
import glgc.jjgys.model.projectvo.jagc.JjgFbgcJtaqssJabxfhlVo;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.ljgc.JjgFbgcLjgcLjtsfysdHtVo;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcJtaqssJabxfhlMapper;
import glgc.jjgys.system.service.JjgFbgcJtaqssJabxfhlService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import glgc.jjgys.system.utils.RowCopy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.DateFormat;
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
 * @since 2023-03-01
 */
@Service
public class JjgFbgcJtaqssJabxfhlServiceImpl extends ServiceImpl<JjgFbgcJtaqssJabxfhlMapper, JjgFbgcJtaqssJabxfhl> implements JjgFbgcJtaqssJabxfhlService {

    @Autowired
    private JjgFbgcJtaqssJabxfhlMapper jjgFbgcJtaqssJabxfhlMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    /**
     * 生成鉴定表
     * @param commonInfoVo
     */
    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException {
        XSSFWorkbook wb = null;
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        //获取数据
        QueryWrapper<JjgFbgcJtaqssJabxfhl> wrapper = new QueryWrapper<>();
        wrapper.like("proname", proname);
        wrapper.like("htd", htd);
        wrapper.like("fbgc", fbgc);
        wrapper.orderByAsc("wzjlx");
        List<JjgFbgcJtaqssJabxfhl> data = jjgFbgcJtaqssJabxfhlMapper.selectList(wrapper);
        //鉴定表要存放的路径
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "58交安钢防护栏.xlsx");
        //健壮性判断如果没有数据返回"请导入数据"
        if (data == null || data.size() == 0) {
            return;
        } else {
            //存放鉴定表的目录
            File fdir = new File(filepath + File.separator + proname + File.separator + htd);
            if (!fdir.exists()) {
                //创建文件根目录
                fdir.mkdirs();
            }
            File directory = new File("service-system/src/main/resources/static");
            String reportPath = directory.getCanonicalPath();
            String name = "钢防护栏.xlsx";
            String path = reportPath + File.separator + name;
            Files.copy(Paths.get(path), new FileOutputStream(f));
            FileInputStream out = new FileInputStream(f);
            wb = new XSSFWorkbook(out);
            createTable(gettableNum(data.size()),wb);
            if(DBtoExcel(data,wb)){

                for (int j = 0; j < wb.getNumberOfSheets(); j++) {   //表内公式  计算 显示结果
                    JjgFbgcCommonUtils.updateFormula(wb, wb.getSheetAt(j));
                }

                FileOutputStream fileOut = new FileOutputStream(f);
                wb.write(fileOut);
                fileOut.flush();
                fileOut.close();
            }
            out.close();
            wb.close();

        }


    }

    /**
     * 写入数据
     * @param data
     * @param wb
     * @return
     * @throws ParseException
     */
    private boolean DBtoExcel(List<JjgFbgcJtaqssJabxfhl> data, XSSFWorkbook wb) throws ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy.MM.dd");
        XSSFSheet sheet = wb.getSheet("防护栏");
        sheet.getRow(1).getCell(2).setCellValue(data.get(0).getProname());
        sheet.getRow(1).getCell(12).setCellValue(data.get(0).getHtd());

        String date = simpleDateFormat.format(data.get(0).getJcsj());
        for (int i = 1; i < data.size(); i++) {
            Date jcsj = data.get(i).getJcsj();
            date = JjgFbgcCommonUtils.getLastDate(date, simpleDateFormat.format(jcsj));
        }
        sheet.getRow(3).getCell(12).setCellValue(date);

        String error[] = getError(data);

        sheet.getRow(5).getCell(2).setCellValue(error[0]);
        sheet.getRow(5).getCell(6).setCellValue(error[1]);
        sheet.getRow(5).getCell(10).setCellValue(error[2]);
        sheet.getRow(5).getCell(13).setCellValue(error[3]);

        int index = 0;
        for (int i = 0; i < data.size(); i++) {
            sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 9, 0, 0));  //位置
            sheet.getRow(index + 7).getCell(0).setCellValue(data.get(i).getWzjlx());

            //波形梁钢护栏立柱埋入深度（mm）
            /*boolean isMerged = false;

            if (i % 5 == 0 && (i + 4) % 8 > 3) {
                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 7 + 15 - 1, 13, 13));
                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 7 + 15 - 1, 14, 14));
                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 7 + 15 - 1, 15, 15));
                isMerged = true;
            } else if (i % 5 == 0 && (i + 4) % 8 <= 3) {
                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 7 + (4 - (i + 4) % 8) * 3 - 1, 13, 13));
                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 7 + (4 - (i + 4) % 8) * 3 - 1, 14, 14));
                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 7 + (4 - (i + 4) % 8) * 3 - 1, 15, 15));
                isMerged = true;
            }

            if (i % 8 == 0 && i % 5 != 0) {
                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 7 + (5 - i % 5) * 3 - 1, 13, 13));
                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 7 + (5 - i % 5) * 3 - 1, 14, 14));
                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 7 + (5 - i % 5) * 3 - 1, 15, 15));
                isMerged = true;
            }*/

            if (data.get(i).getMrsdscz()!=null && !"".equals(data.get(i).getMrsdscz())) {
                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 7 + 2, 13, 13));
                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 7 + 2, 14, 14));
                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 7 + 2, 15, 15));
                sheet.getRow(index + 7).getCell(13).setCellValue(Double.valueOf(data.get(i).getMrsdscz()));
                sheet.getRow(index + 7).getCell(14).setCellFormula("IF(AND("
                        + sheet.getRow(index + 7).getCell(13).getReference() + ":"
                        + sheet.getRow(index + 9).getCell(13).getReference() + "-"
                        + data.get(i).getMrsdgdz().replace("≥", "") + ">=0),\"√\",\"\")");//M=IF(AND(N8:N10-1400>=0),"√","")
                sheet.getRow(index + 7).getCell(15).setCellFormula("IF(AND("
                        + sheet.getRow(index + 7).getCell(13).getReference() + ":"
                        + sheet.getRow(index + 9).getCell(13).getReference() + "-"
                        + data.get(i).getMrsdgdz().replace("≥", "") + ">=0),\"\",\"×\")");//M=IF(AND(L8:L12-1400>=0),"","×")

            } else {
                sheet.getRow(index + 7).getCell(13).setCellValue("-");
                sheet.getRow(index + 7).getCell(14).setCellValue("-");
                sheet.getRow(index + 7).getCell(15).setCellValue("-");
            }

                sheet.getRow(index + 7  ).getCell(1).setCellValue(Double.valueOf((1)));  //序号
                sheet.getRow(index + 8).getCell(1).setCellValue(Double.valueOf((2)));  //序号
                sheet.getRow(index + 9).getCell(1).setCellValue(Double.valueOf((3)));  //序号
                //波形梁板基底金属厚度（mm）
                sheet.getRow(index + 7 ).getCell(2).setCellValue(Double.valueOf(data.get(i).getJdhdscz1()));  //实测值
                sheet.getRow(index + 8).getCell(2).setCellValue(Double.valueOf(data.get(i).getJdhdscz2()));  //实测值
                sheet.getRow(index + 9).getCell(2).setCellValue(Double.valueOf(data.get(i).getJdhdscz3()));  //实测值
                //波形梁钢护栏立柱
                sheet.getRow(index+7).getCell(6).setCellValue(Double.valueOf(data.get(i).getLzbhscz1()));
                sheet.getRow(index+8).getCell(6).setCellValue(Double.valueOf(data.get(i).getLzbhscz2()));
                sheet.getRow(index+9).getCell(6).setCellValue(Double.valueOf(data.get(i).getLzbhscz3()));

                //sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 9, 10, 10));
//                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 9, 11, 11));
//                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 9, 12, 12));
                if (data.get(i).getZxgdscz1().equals("") || data.get(i).getZxgdscz1() == null) {
                    sheet.getRow(index + 7).getCell(10).setCellValue(Double.valueOf("0"));
                    sheet.getRow(index + 8).getCell(10).setCellValue(Double.valueOf("0"));
                    sheet.getRow(index + 9).getCell(10).setCellValue(Double.valueOf("0"));
                } else {
                    sheet.getRow(index + 7).getCell(10).setCellValue(Double.valueOf(data.get(i).getZxgdscz1()));
                    sheet.getRow(index + 8).getCell(10).setCellValue(Double.valueOf(data.get(i).getZxgdscz2()));
                    sheet.getRow(index + 9).getCell(10).setCellValue(Double.valueOf(data.get(i).getZxgdscz3()));
                }


                if (data.get(i).getZxgdgdz()== null || data.get(i).getZxgdyxpsz() == null || "".equals(data.get(i).getZxgdgdz()) || "".equals(data.get(i).getZxgdyxpsz())) {
                    ArrayList<String> list = gethlgdGDZ(data.get(i));
                    if (list.size() == 3) {
                        String zxgdgdz = data.get(i).getZxgdgdz();
                        String zxgdyxpsz = data.get(i).getZxgdyxpsz();
                        String zxgdyxpsf = data.get(i).getZxgdyxpsf();
                        zxgdgdz = list.get(0);
                        zxgdyxpsz = list.get(1);
                        zxgdyxpsf= list.get(2);
                        sheet.getRow(index + 7).getCell(11).setCellFormula("IF(AND("
                                + sheet.getRow(index + 7).getCell(10).getReference() + "-"
                                + data.get(i).getZxgdgdz() + "<="
                                + data.get(i).getZxgdyxpsz() + ","
                                + sheet.getRow(index + 7).getCell(10).getReference() + "-"
                                + data.get(i).getZxgdgdz() + ">=-"
                                + data.get(i).getZxgdyxpsf() + "),\"√\",\"\")");
                        sheet.getRow(index + 7).getCell(12).setCellFormula("IF(AND("
                                + sheet.getRow(index + 7).getCell(10).getReference() + "-"
                                + data.get(i).getZxgdgdz() + "<="
                                + data.get(i).getZxgdyxpsz() + ","
                                + sheet.getRow(index + 7).getCell(10).getReference() + "-"
                                + data.get(i).getZxgdgdz() + ">=-"
                                + data.get(i).getZxgdyxpsf() + "),\"\",\"×\")");
                    } else if (list.size() == 1) {
                        String zxgdgdz = data.get(i).getZxgdgdz();
                        zxgdgdz = list.get(0);
                        sheet.getRow(index + 7).getCell(11).setCellFormula("IF(AND("
                                + sheet.getRow(index + 7).getCell(10).getReference() + "-"
                                + data.get(i).getZxgdgdz() + ">=0"
                                + "),\"√\",\"\")");
                        sheet.getRow(index + 7).getCell(12).setCellFormula("IF(AND("
                                + sheet.getRow(index + 7).getCell(10).getReference() + "-"
                                + data.get(i).getZxgdgdz() + ">=0"
                                + "),\"\",\"×\")");
                    }
                } else {
                    sheet.getRow(index+7).getCell(11).setCellFormula("IF(AND("
                            +sheet.getRow(index+7).getCell(10).getReference()+"-"
                            +data.get(i).getZxgdgdz()+"<="  //21
                            +data.get(i).getZxgdyxpsz()+"," //22
                            +sheet.getRow(index+7).getCell(10).getReference()+"-"
                            +data.get(i).getZxgdgdz() +">=-" //21
                            +data.get(i).getZxgdyxpsf()+"),\"√\",\"\")");//23
                    sheet.getRow(index+7).getCell(12).setCellFormula("IF(AND("
                            +sheet.getRow(index+7).getCell(10).getReference()+"-"
                            +data.get(i).getZxgdgdz()+"<="
                            +data.get(i).getZxgdyxpsz()+","
                            +sheet.getRow(index+7).getCell(10).getReference()+"-"
                            +data.get(i).getZxgdgdz()+">=-"
                            +data.get(i).getZxgdyxpsf()+"),\"\",\"×\")");

                    sheet.getRow(index+8).getCell(11).setCellFormula("IF(AND("
                            +sheet.getRow(index+8).getCell(10).getReference()+"-"
                            +data.get(i).getZxgdgdz()+"<="  //21
                            +data.get(i).getZxgdyxpsz()+"," //22
                            +sheet.getRow(index+8).getCell(10).getReference()+"-"
                            +data.get(i).getZxgdgdz() +">=-" //21
                            +data.get(i).getZxgdyxpsf()+"),\"√\",\"\")");//23
                    sheet.getRow(index+8).getCell(12).setCellFormula("IF(AND("
                            +sheet.getRow(index+8).getCell(10).getReference()+"-"
                            +data.get(i).getZxgdgdz()+"<="
                            +data.get(i).getZxgdyxpsz()+","
                            +sheet.getRow(index+8).getCell(10).getReference()+"-"
                            +data.get(i).getZxgdgdz()+">=-"
                            +data.get(i).getZxgdyxpsf()+"),\"\",\"×\")");


                    sheet.getRow(index+9).getCell(11).setCellFormula("IF(AND("
                            +sheet.getRow(index+9).getCell(10).getReference()+"-"
                            +data.get(i).getZxgdgdz()+"<="  //21
                            +data.get(i).getZxgdyxpsz()+"," //22
                            +sheet.getRow(index+9).getCell(10).getReference()+"-"
                            +data.get(i).getZxgdgdz() +">=-" //21
                            +data.get(i).getZxgdyxpsf()+"),\"√\",\"\")");//23
                    sheet.getRow(index+9).getCell(12).setCellFormula("IF(AND("
                            +sheet.getRow(index+9).getCell(10).getReference()+"-"
                            +data.get(i).getZxgdgdz()+"<="
                            +data.get(i).getZxgdyxpsz()+","
                            +sheet.getRow(index+9).getCell(10).getReference()+"-"
                            +data.get(i).getZxgdgdz()+">=-"
                            +data.get(i).getZxgdyxpsf()+"),\"\",\"×\")");
                }

                //波形梁板基底金属厚度的平均值公式
                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 9, 3, 3));
                sheet.getRow(index + 7).getCell(3).setCellFormula("AVERAGE("
                        + sheet.getRow(index + 7).getCell(2).getReference() + ":"
                        + sheet.getRow(index + 11).getCell(2).getReference() + ")");   //D8=AVERAGE(C8:C12)
                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 9, 4, 4));
                //波形梁板基底金属厚度的结论
                sheet.getRow(index + 7).getCell(4).setCellFormula("IF(MIN("
                        + sheet.getRow(index + 7).getCell(2).getReference() + ":"
                        + sheet.getRow(index + 9).getCell(2).getReference() + ")>=2.95,IF("
                        + sheet.getRow(index + 7).getCell(3).getReference() + ">=3,\"√\",\"\"),\"\")"); //E8== IF(MIN(C8:C10)>=2.95,IF(D8>=3,"√",""),"")
                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 9, 5, 5));
                sheet.getRow(index + 7).getCell(5).setCellFormula("IF(MIN("
                        + sheet.getRow(index + 7).getCell(2).getReference() + ":"
                        + sheet.getRow(index + 9).getCell(2).getReference() + ")>=2.95,IF("
                        + sheet.getRow(index + 7).getCell(3).getReference() + ">=3,\"\",\"×\"),\"×\")"); //F8=IF(MIN(C8:C10)>=2.95,IF(D8>=3,"","×"),"×")


                //波形梁钢护栏立柱
                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 9, 7, 7));
                sheet.getRow(index + 7).getCell(7).setCellFormula("AVERAGE("
                        + sheet.getRow(index + 7).getCell(6).getReference() + ":"
                        + sheet.getRow(index + 11).getCell(6).getReference() + ")");   //H8==AVERAGE(G8:G12)

                String errorBXL = "";
                if (data.get(i).getWzjlx().contains("圆柱")) {
                    if (data.get(i).getWzjlx().contains(",")) {
                        errorBXL = error[1].substring(error[1].indexOf("圆柱") + 3, error[1].indexOf(","));
                    } else {
                        errorBXL = error[1].substring(error[1].indexOf("圆柱") + 3);
                    }
                } else if (data.get(i).getWzjlx().contains("方柱")) {
                    errorBXL = error[1].substring(error[1].indexOf("方柱") + 3);
                } else if (data.get(i).getWzjlx().contains("≥")) {
                    errorBXL = error[1].substring(error[1].indexOf("≥") + 1);
                } else {
                    errorBXL = error[1];
                }

                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 9, 8, 8));
                sheet.getRow(index + 7).getCell(8).setCellFormula("IF(AND("
                        + sheet.getRow(index + 7).getCell(7).getReference() + "-"
                        + errorBXL + ">=0"
                        + "),\"√\",\"\")");
                sheet.addMergedRegion(new CellRangeAddress(index + 7, index + 9, 9, 9));
                sheet.getRow(index + 7).getCell(9).setCellFormula("IF(AND("
                        + sheet.getRow(index + 7).getCell(7).getReference() + "-"
                        + errorBXL + ">=0"
                        + "),\"\",\"×\")");

                index +=3;

            /*
             * 所有数据填写完毕，此处计算合计
             */
            //波形梁板基底金属
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 4).getCell(3).setCellFormula("COUNT(D8:" + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(3).getReference() + ")");
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 3).getCell(3).setCellFormula("COUNTIF(E8:" + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(3).getReference() + ",\"√\")");//=COUNTIF(E8:E302,"√")
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 2).getCell(3).setCellFormula("COUNTIF(F8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(5).getReference() + ",\"×\")");//=COUNTIF(F8:F302,"×")
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 1).getCell(3).setCellFormula("ROUND("
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 3).getCell(3).getReference() + "/"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 4).getCell(3).getReference() + "*100,2)");//=D304/D303*100
            //波形梁钢护栏立柱
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 4).getCell(6).setCellFormula("COUNT(H8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(7).getReference() + ")");
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 3).getCell(6).setCellFormula("COUNTIF(I8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(8).getReference() + ",\"√\")");//=COUNTIF(I8:I302,"√")
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 2).getCell(6).setCellFormula("COUNTIF(J8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(9).getReference() + ",\"×\")");//=COUNTIF(J8:J302,"×")
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 1).getCell(6).setCellFormula("ROUND("
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 3).getCell(6).getReference() + "/"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 4).getCell(6).getReference() + "*100,2)");//=D304/D303*100
            //波形梁钢护栏横梁中心
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 4).getCell(10).setCellFormula("COUNT(K8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(10).getReference() + ")");
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 3).getCell(10).setCellFormula("COUNTIF(L8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(11).getReference() + ",\"√\")");//=COUNTIF(L8:L302,"√")
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 2).getCell(10).setCellFormula("COUNTIF(M8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(12).getReference() + ",\"×\")");//=COUNTIF(M8:M302,"×")
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 1).getCell(10).setCellFormula("ROUND("
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 3).getCell(10).getReference() + "/"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 4).getCell(10).getReference() + "*100,2)");//=D304/D303*100
            //波形梁钢护栏立柱埋入深度（mm）
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 4).getCell(13).setCellFormula("COUNT(N8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(13).getReference() + ")");
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 3).getCell(13).setCellFormula("COUNTIF(O8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(14).getReference() + ",\"√\")");//=COUNTIF(O8:O302,"√")
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 2).getCell(13).setCellFormula("COUNTIF(P8:"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 5).getCell(15).getReference() + ",\"×\")");//=COUNTIF(P8:P302,"×")
            sheet.getRow(sheet.getPhysicalNumberOfRows() - 1).getCell(13).setCellFormula("ROUND("
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 3).getCell(13).getReference() + "/"
                    + sheet.getRow(sheet.getPhysicalNumberOfRows() - 4).getCell(13).getReference() + "*100,2)");//=D304/D303*100

        }
        return true;
    }


    private ArrayList<String> gethlgdGDZ(JjgFbgcJtaqssJabxfhl data) {
        ArrayList<String> list = new ArrayList<String>();
        if(data.getWzjlx().contains("两波板")){
            String str = data.getZxgdgdz();//"两波板600±20\n三波板697±20";
            if(str.contains("两波板")){
                if(str.contains("两波板") && str.contains("三波板")){
                    if(str.indexOf("两波板") < str.indexOf("三波板")){
                        String str1 = str.substring(str.indexOf("两波板"), str.indexOf("三波板")).replaceAll("\n", "");
                        if(str1.contains("±")){
                            list.add(str1.substring(3, str.indexOf("±")));
                            list.add(str1.substring(str.indexOf("±")+1, str1.length()));
                            list.add(str1.substring(str.indexOf("±")+1, str1.length()));
                        }
                        else{
                            list.add(str1.substring(3, str.indexOf("±")));
                        }
                    }
                    else{
                        String str1 = str.substring(str.indexOf("两波板"), str.length()).replaceAll("\n", "");
                        if(str1.contains("±")){
                            list.add(str1.substring(3, str.indexOf("±")));
                            list.add(str1.substring(str.indexOf("±")+1, str1.length()));
                            list.add(str1.substring(str.indexOf("±")+1, str1.length()));
                        }
                        else{
                            list.add(str1.substring(3, str.indexOf("±")));
                        }
                    }
                }
                else{
                    String str1 = str.substring(str.indexOf("两波板"), str.length()).replaceAll("\n", "");
                    if(str1.contains("±")){
                        list.add(str1.substring(3, str.indexOf("±")));
                        list.add(str1.substring(str.indexOf("±")+1, str1.length()));
                        list.add(str1.substring(str.indexOf("±")+1, str1.length()));
                    }
                    else{
                        list.add(str1.substring(3, str.indexOf("±")));
                    }
                }
                return list;
            }
            else{
                return null;
            }
        }
        else{
            String str = data.getZxgdgdz();//"两波板600±20\n三波板697±20";
            if(str.contains("三波板")){
                if(str.contains("两波板") && str.contains("三波板")){
                    if(str.indexOf("两波板") < str.indexOf("三波板")){
                        String str1 = str.substring(str.indexOf("三波板"), str.length()).replaceAll("\n", "");
                        if(str1.contains("±")){
                            list.add(str1.substring(3, str.indexOf("±")));
                            list.add(str1.substring(str.indexOf("±")+1, str1.length()));
                            list.add(str1.substring(str.indexOf("±")+1, str1.length()));
                        }
                        else{
                            list.add(str1.substring(3, str.indexOf("±")));
                        }
                    }
                    else{
                        String str1 = str.substring(str.indexOf("三波板"), str.indexOf("两波板")).replaceAll("\n", "");
                        if(str1.contains("±")){
                            list.add(str1.substring(3, str.indexOf("±")));
                            list.add(str1.substring(str.indexOf("±")+1, str1.length()));
                            list.add(str1.substring(str.indexOf("±")+1, str1.length()));
                        }
                        else{
                            list.add(str1.substring(3, str.indexOf("±")));
                        }
                    }
                }
                else{
                    String str1 = str.substring(str.indexOf("三波板"), str.indexOf("两波板")).replaceAll("\n", "");
                    if(str1.contains("±")){
                        list.add(str1.substring(3, str.indexOf("±")));
                        list.add(str1.substring(str.indexOf("±")+1, str1.length()));
                        list.add(str1.substring(str.indexOf("±")+1, str1.length()));
                    }
                    else{
                        list.add(str1.substring(3, str.indexOf("±")));
                    }
                }
                return list;
            }
            else{
                return null;
            }
        }
    }

    /**
     * 获取偏差值
     * @param data
     * @return
     */
    private String[] getError(List<JjgFbgcJtaqssJabxfhl> data) {
        String[] error = {"","","",""};
        String temp = "";
        ArrayList<String> errorlist = new ArrayList<String>();
        for(int i =0; i < data.size(); i++){
            if(!errorlist.contains(data.get(i).getWzjlx())){//  按位置 存偏差
                errorlist.add((data.get(i).getWzjlx()));
                if(!"".equals(data.get(i).getJdhdgdz()) && !data.get(i).getJdhdgdz().equals("0")){
                    if(data.get(i).getJdhdgdz().contains("≥")){
                        temp+=data.get(i).getJdhdgdz();
                    }
                    if(!error[0].contains(temp)){
                        error[0] += temp;
                    }
                    temp = "";
                }
                if(data.get(i).getWzjlx().contains("圆柱")){  //位置中有圆柱
                    if(!"".equals(data.get(i).getLzbhgdz()) && !data.get(i).getLzbhgdz().equals("0")){
                        if(data.get(i).getLzbhgdz().contains("圆柱≥")){
                            temp+=data.get(i).getLzbhgdz();
                        }else if(data.get(i).getLzbhgdz().contains("≥")){
                            temp+="圆柱"+data.get(i).getLzbhgdz();
                        }
                        else{
                            temp+="圆柱≥"+data.get(i).getLzbhgdz();
                        }
                    }
                }
                else if(data.get(i).getWzjlx().contains("方柱")){
                    if(!"".equals(data.get(i).getLzbhgdz()) && !data.get(i).getLzbhgdz().equals("0")){
                        temp+="方柱"+data.get(i).getLzbhgdz();
                    }
                }
                if(!error[1].contains(temp)){
                    error[1] += temp;
                }temp = "";
                if(data.get(i).getWzjlx().contains("两波")){
                    if(!"".equals(data.get(i).getZxgdgdz()) && !data.get(i).getZxgdgdz().equals("0")){
                        temp+="两波板"+Double.valueOf(data.get(i).getZxgdgdz()).intValue();
                        if(!data.get(i).getZxgdyxpsz().equals("0") && !data.get(i).getZxgdyxpsf().equals("0")){
                            if(data.get(i).getZxgdyxpsz().equals(data.get(i).getZxgdyxpsf())){
                                temp+="±"+Double.valueOf(data.get(i).getZxgdyxpsz()).intValue();
                            }
                        }else{
                            if(!data.get(i).getZxgdyxpsz().equals("0")){
                                temp+="+"+Double.valueOf(data.get(i).getZxgdyxpsz()).intValue();
                            }
                            if(!data.get(i).getZxgdyxpsf().equals("0")){
                                temp+="-"+Double.valueOf(data.get(i).getZxgdyxpsf()).intValue();
                            }
                        }
                    }
                }
                else if(data.get(i).getWzjlx().contains("三波")){
                    if(!"".equals(data.get(i).getZxgdgdz()) && !data.get(i).getZxgdgdz().equals("0")){
                        temp+="三波板"+Double.valueOf(data.get(i).getZxgdgdz()).intValue();
                        if(!data.get(i).getZxgdyxpsz().equals("0") && !data.get(i).getZxgdyxpsf().equals("0")){
                            if(data.get(i).getZxgdyxpsz().equals(data.get(i).getZxgdyxpsf())){
                                temp+="±"+Double.valueOf(data.get(i).getZxgdyxpsz()).intValue();
                            }
                        }else{
                            if(!data.get(i).getZxgdyxpsz().equals("0")){
                                temp+="+"+Double.valueOf(data.get(i).getZxgdyxpsz()).intValue();
                            }
                            if(!data.get(i).getZxgdyxpsf().equals("0")){
                                temp+="-"+Double.valueOf(data.get(i).getZxgdyxpsf()).intValue();
                            }
                        }
                    }
                }
                if(!error[2].contains(temp)){
                    error[2] += temp;
                }temp = "";


                if(data.get(i).getWzjlx().contains("圆柱")){
                    if(!"".equals(data.get(i).getMrsdgdz()) && !data.get(i).getMrsdgdz().equals("0")){
                        if(data.get(i).getMrsdgdz().contains("≥")){
                            temp+="圆柱"+data.get(i).getMrsdgdz();
                        }
                        else{
                            temp+="圆柱≥"+data.get(i).getMrsdgdz();
                        }
                    }
                }
                else if(data.get(i).getWzjlx().contains("方柱")){
                    if(!"".equals(data.get(i).getMrsdgdz()) && !data.get(i).getMrsdgdz().equals("0")){
                        if(data.get(i).getMrsdgdz().contains("≥")){
                            temp+="方柱"+data.get(i).getMrsdgdz();
                        }
                        else{
                            temp+="方柱≥"+data.get(i).getMrsdgdz();
                        }
                    }
                }
                if(!error[3].contains(temp)){
                    error[3] += temp;
                }temp = "";
            }
        }
        return error;
    }


    /**
     * 创建模板页
     * @param tableNum
     * @param wb
     */
    private void createTable(int tableNum, XSSFWorkbook wb){
        int record = 0;
        record = tableNum;

        for (int i = 1; i < record; i++) {
            if (i < record - 1) {
                RowCopy.copyRows(wb, "防护栏", "防护栏", 7, 30, (i - 1) * 24 + 31);
            } else {
                RowCopy.copyRows(wb, "防护栏", "防护栏", 7, 27, (i - 1) * 24 + 31);
            }
        }
        if (record >= 1)
            RowCopy.copyRows(wb, "source", "防护栏", 0, 4, (record - 1) * 24 + 27);


        wb.setPrintArea(wb.getSheetIndex("防护栏"), 0, 15, 0, record * 24 + 6);

    }

    /**
     * 模板页数
     * @param tableNum
     * @return
     */
    private int gettableNum(int tableNum) {
        return tableNum %8 ==0|| tableNum % 8== 7 ? tableNum/8+2 : tableNum/8+1;
    }

    /**
     * 导出数据模板文件
     * @param response
     * @throws IOException
     */
    @Override
    public void exportjabxfhl(HttpServletResponse response) throws IOException {
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
        String filename = "交安波形防护栏实测数据.xls";// 设置下载时客户端Excel的名称
        filename = new String((filename).getBytes("GBK"), "ISO8859_1");
        response.setContentType("application/octet-stream;charset=UTF-8");
        response.addHeader("Content-Disposition", "attachment;filename=" + filename);
        OutputStream ouputStream = response.getOutputStream();
        workbook.write(ouputStream);
        ouputStream.flush();
        ouputStream.close();

    }

    /**
     * 导入数据
     * @param file
     * @param commonInfoVo
     * @throws IOException
     * @throws ParseException
     */
    @Override
    public void importjabxfhl(MultipartFile file, CommonInfoVo commonInfoVo) throws IOException, ParseException {
        // 将文件流传过来，变成workbook对象
        XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
        //获得文本
        XSSFSheet sheet = workbook.getSheetAt(0);
        //获得行数
        int rows = sheet.getPhysicalNumberOfRows();
        //获得列数
        int columns = 0;
        for(int i=1;i<rows;i++){
            XSSFRow row = sheet.getRow(i);
            columns  = row.getPhysicalNumberOfCells();
        }
        JjgFbgcJtaqssJabxfhlVo jjgFbgcJtaqssJabxfhlVo = new JjgFbgcJtaqssJabxfhlVo();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        List<String> titlelist = new ArrayList();
        for (int n=0;n<1;n++){
            for (int m=0;m<rows;m++){
                XSSFRow row = sheet.getRow(m);
                XSSFCell cell = row.getCell(n);
                titlelist.add(cell.toString());
            }
        }
        List<String> checklist = Arrays.asList("检测日期","位置及类型","基底厚度规定值(mm)","基底厚度实测值1(mm)",
                "基底厚度实测值2(mm)","基底厚度实测值3(mm)","立柱壁厚规定值(mm)","立柱壁厚实测值1(mm)","立柱壁厚实测值2(mm)",
                "立柱壁厚实测值3(mm)","中心高度规定值(mm)","中心高度允许偏差+(mm)",
                "中心高度允许偏差-(mm)","中心高度实测值1(mm)","中心高度实测值2(mm)","中心高度实测值3(mm)","埋入深度规定值(mm)","埋入深度实测值(mm)");
        if(checklist.equals(titlelist)){
            for (int j = 1;j<columns;j++){//列
                Map<String,Object> map = new HashMap<>();
                Field[] fields = jjgFbgcJtaqssJabxfhlVo.getClass().getDeclaredFields();
                JjgFbgcJtaqssJabxfhl jjgFbgcJtaqssJabxfhl = new JjgFbgcJtaqssJabxfhl();
                for(int k=0;k<rows;k++){//行
                    //列是不变的 行增加
                    XSSFRow row = sheet.getRow(k);
                    XSSFCell cell = row.getCell(j);

                    switch (cell.getCellType()){
                        case XSSFCell.CELL_TYPE_STRING ://String
                            cell.setCellType(CellType.STRING);
                            map.put(fields[k].getName(),cell.getStringCellValue());//属性赋值
                            break;
                        case XSSFCell.CELL_TYPE_BOOLEAN ://bealean
                            //cell.setCellType(XSSFCell.CELL_TYPE_STRING);
                            cell.setCellType(CellType.STRING);
                            //map.put(fields[k].getName(),Boolean.valueOf(cell.getBooleanCellValue()).toString());//属性赋值
                            map.put(fields[k].getName(),String.valueOf(cell.getStringCellValue()));//属性赋值
                            break;
                        case XSSFCell.CELL_TYPE_NUMERIC ://number
                            //默认日期读取出来是数字，判断是否是日期格式的数字
                            if(DateUtil.isCellDateFormatted(cell)){
                                //读取的数字是日期，转换一下格式
                                DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                                Date date = cell.getDateCellValue();
                                System.out.println(date);
                                map.put(fields[k].getName(),dateFormat.format(date));//属性赋值
                            }else {//不是日期直接赋值
                                System.out.println(cell);

                                //map.put(fields[k].getName(),Double.valueOf(cell.getNumericCellValue()).toString());//属性赋值
                                map.put(fields[k].getName(),String.valueOf(cell.getNumericCellValue()));//属性赋值
                            }
                            break;
                        case XSSFCell.CELL_TYPE_BLANK :
                            cell.setCellType(CellType.STRING);
                            map.put(fields[k].getName(),"");//属性赋值
                            break;
                        default:
                            System.out.println("未知类型------>"+cell);
                    }
                }
                jjgFbgcJtaqssJabxfhl.setJcsj(simpleDateFormat.parse((String) map.get("jcsj")));
                jjgFbgcJtaqssJabxfhl.setWzjlx((String) map.get("wzjlx"));
                jjgFbgcJtaqssJabxfhl.setJdhdgdz((String) map.get("jdhdgdz"));
                jjgFbgcJtaqssJabxfhl.setJdhdscz1((String) map.get("jdhdscz1"));
                jjgFbgcJtaqssJabxfhl.setJdhdscz2((String) map.get("jdhdscz2"));
                jjgFbgcJtaqssJabxfhl.setJdhdscz3((String) map.get("jdhdscz3"));
                jjgFbgcJtaqssJabxfhl.setLzbhgdz((String) map.get("lzbhgdz"));
                jjgFbgcJtaqssJabxfhl.setLzbhscz1((String) map.get("lzbhscz1"));
                jjgFbgcJtaqssJabxfhl.setLzbhscz2((String) map.get("lzbhscz2"));
                jjgFbgcJtaqssJabxfhl.setLzbhscz3((String) map.get("lzbhscz3"));
                jjgFbgcJtaqssJabxfhl.setZxgdgdz((String) map.get("zxgdgdz"));
                jjgFbgcJtaqssJabxfhl.setZxgdyxpsz((String) map.get("zxgdyxpsz"));
                jjgFbgcJtaqssJabxfhl.setZxgdyxpsf((String) map.get("zxgdyxpsf"));
                jjgFbgcJtaqssJabxfhl.setZxgdscz1((String) map.get("zxgdscz1"));
                jjgFbgcJtaqssJabxfhl.setZxgdscz2((String) map.get("zxgdscz2"));
                jjgFbgcJtaqssJabxfhl.setZxgdscz3((String) map.get("zxgdscz3"));
                jjgFbgcJtaqssJabxfhl.setMrsdgdz((String) map.get("mrsdgdz"));
                jjgFbgcJtaqssJabxfhl.setMrsdscz((String) map.get("mrsdscz"));

                jjgFbgcJtaqssJabxfhl.setProname(commonInfoVo.getProname());
                jjgFbgcJtaqssJabxfhl.setHtd(commonInfoVo.getHtd());
                jjgFbgcJtaqssJabxfhl.setFbgc(commonInfoVo.getFbgc());
                jjgFbgcJtaqssJabxfhl.setCreatetime(new Date());
                jjgFbgcJtaqssJabxfhlMapper.insert(jjgFbgcJtaqssJabxfhl);
            }
        }else {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        String fbgc = commonInfoVo.getFbgc();
        String title = "道路防护栏施工质量鉴定表（波形梁钢护栏）";
        String sheetname = "防护栏";
        //获取鉴定表文件
        File f = new File(filepath + File.separator + proname + File.separator + htd + File.separator + "58交安钢防护栏.xlsx");
        if (!f.exists()) {
            return null;
        } else {
            XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(f));
            //读取工作表
            XSSFSheet slSheet = xwb.getSheet(sheetname);
            XSSFCell bt = slSheet.getRow(0).getCell(0);//标题
            XSSFCell xmname = slSheet.getRow(1).getCell(2);//项目名
            XSSFCell htdname = slSheet.getRow(1).getCell(12);//合同段名
            XSSFCell hd = slSheet.getRow(2).getCell(12);//分布工程名
            List<Map<String, Object>> mapList = new ArrayList<>();
            Map<String, Object> jgmap1 = new HashMap<>();
            Map<String, Object> jgmap2 = new HashMap<>();
            Map<String, Object> jgmap3 = new HashMap<>();
            Map<String, Object> jgmap4 = new HashMap<>();
            DecimalFormat df = new DecimalFormat(".00");
            DecimalFormat decf = new DecimalFormat("0.##");
            if (proname.equals(xmname.toString()) && title.equals(bt.toString()) && htd.equals(htdname.toString()) && fbgc.equals(hd.toString())) {
                //获取到最后一行
                int lastRowNum = slSheet.getLastRowNum();
                slSheet.getRow(lastRowNum - 3).getCell(3).setCellType(XSSFCell.CELL_TYPE_STRING);//总点数
                slSheet.getRow(lastRowNum - 2).getCell(3).setCellType(XSSFCell.CELL_TYPE_STRING);//合格点数
                slSheet.getRow(lastRowNum - 1).getCell(3).setCellType(XSSFCell.CELL_TYPE_STRING);//不合格点数
                slSheet.getRow(lastRowNum).getCell(3).setCellType(XSSFCell.CELL_TYPE_STRING);//合格率
                double zds1 = Double.valueOf(slSheet.getRow(lastRowNum - 3).getCell(3).getStringCellValue());
                double hgds1 = Double.valueOf(slSheet.getRow(lastRowNum - 2).getCell(3).getStringCellValue());
                double bhgds1 = Double.valueOf(slSheet.getRow(lastRowNum - 2).getCell(3).getStringCellValue());
                double hgl1 = Double.valueOf(slSheet.getRow(lastRowNum).getCell(3).getStringCellValue());
                String zdsz1 = decf.format(zds1);
                String hgdsz1 = decf.format(hgds1);
                String bhgdsz1 = decf.format(bhgds1);
                String hglz1 = df.format(hgl1);
                jgmap1.put("检测项目", "波形梁板基底金属厚度（mm）");
                jgmap1.put("总点数", zdsz1);
                jgmap1.put("合格点数", hgdsz1);
                jgmap1.put("不合格点数", bhgdsz1);
                jgmap1.put("合格率", hglz1);

                slSheet.getRow(lastRowNum - 3).getCell(6).setCellType(XSSFCell.CELL_TYPE_STRING);//总点数
                slSheet.getRow(lastRowNum - 2).getCell(6).setCellType(XSSFCell.CELL_TYPE_STRING);//合格点数
                slSheet.getRow(lastRowNum - 1).getCell(6).setCellType(XSSFCell.CELL_TYPE_STRING);//不合格点数
                slSheet.getRow(lastRowNum).getCell(6).setCellType(XSSFCell.CELL_TYPE_STRING);//合格率
                double zds2 = Double.valueOf(slSheet.getRow(lastRowNum - 3).getCell(6).getStringCellValue());
                double hgds2 = Double.valueOf(slSheet.getRow(lastRowNum - 2).getCell(6).getStringCellValue());
                double bhgds2 = Double.valueOf(slSheet.getRow(lastRowNum - 2).getCell(6).getStringCellValue());
                double hgl2 = Double.valueOf(slSheet.getRow(lastRowNum).getCell(6).getStringCellValue());
                String zdsz2 = decf.format(zds2);
                String hgdsz2 = decf.format(hgds2);
                String bhgdsz2 = decf.format(bhgds2);
                String hglz2 = df.format(hgl2);
                jgmap2.put("检测项目", "波形梁钢护栏立柱壁厚（mm）");
                jgmap2.put("总点数", zdsz2);
                jgmap2.put("合格点数", hgdsz2);
                jgmap2.put("不合格点数", bhgdsz2);
                jgmap2.put("合格率", hglz2);

                slSheet.getRow(lastRowNum - 3).getCell(10).setCellType(XSSFCell.CELL_TYPE_STRING);//总点数
                slSheet.getRow(lastRowNum - 2).getCell(10).setCellType(XSSFCell.CELL_TYPE_STRING);//合格点数
                slSheet.getRow(lastRowNum - 1).getCell(10).setCellType(XSSFCell.CELL_TYPE_STRING);//不合格点数
                slSheet.getRow(lastRowNum).getCell(10).setCellType(XSSFCell.CELL_TYPE_STRING);//合格率
                double zds3 = Double.valueOf(slSheet.getRow(lastRowNum - 3).getCell(10).getStringCellValue());
                double hgds3 = Double.valueOf(slSheet.getRow(lastRowNum - 2).getCell(10).getStringCellValue());
                double bhgds3 = Double.valueOf(slSheet.getRow(lastRowNum - 2).getCell(10).getStringCellValue());
                double hgl3 = Double.valueOf(slSheet.getRow(lastRowNum).getCell(10).getStringCellValue());
                String zdsz3 = decf.format(zds3);
                String hgdsz3 = decf.format(hgds3);
                String bhgdsz3 = decf.format(bhgds3);
                String hglz3 = df.format(hgl3);
                jgmap3.put("检测项目", "波形梁钢护栏横梁中心高度（mm）");
                jgmap3.put("总点数", zdsz3);
                jgmap3.put("合格点数", hgdsz3);
                jgmap3.put("不合格点数", bhgdsz3);
                jgmap3.put("合格率", hglz3);

                slSheet.getRow(lastRowNum - 3).getCell(13).setCellType(XSSFCell.CELL_TYPE_STRING);//总点数
                slSheet.getRow(lastRowNum - 2).getCell(13).setCellType(XSSFCell.CELL_TYPE_STRING);//合格点数
                slSheet.getRow(lastRowNum - 1).getCell(13).setCellType(XSSFCell.CELL_TYPE_STRING);//不合格点数
                slSheet.getRow(lastRowNum).getCell(13).setCellType(XSSFCell.CELL_TYPE_STRING);//合格率
                double zds4 = Double.valueOf(slSheet.getRow(lastRowNum - 3).getCell(13).getStringCellValue());
                double hgds4 = Double.valueOf(slSheet.getRow(lastRowNum - 2).getCell(13).getStringCellValue());
                double bhgds4 = Double.valueOf(slSheet.getRow(lastRowNum - 2).getCell(13).getStringCellValue());
                double hgl4 = Double.valueOf(slSheet.getRow(lastRowNum).getCell(13).getStringCellValue());
                String zdsz4 = decf.format(zds4);
                String hgdsz4 = decf.format(hgds4);
                String bhgdsz4 = decf.format(bhgds4);
                String hglz4 = df.format(hgl4);
                jgmap4.put("检测项目", "波形梁钢护栏立柱埋入深度（mm）");
                jgmap4.put("总点数", zdsz4);
                jgmap4.put("合格点数", hgdsz4);
                jgmap4.put("不合格点数", bhgdsz4);
                jgmap4.put("合格率", hglz4);

                mapList.add(jgmap1);
                mapList.add(jgmap2);
                mapList.add(jgmap3);
                mapList.add(jgmap4);
                return mapList;
            }
            return null;
        }
    }
}
