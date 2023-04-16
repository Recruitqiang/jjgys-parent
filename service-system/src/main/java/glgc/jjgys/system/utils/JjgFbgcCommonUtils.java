package glgc.jjgys.system.utils;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;



import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.nio.channels.Channels;
import java.nio.channels.FileChannel;
import java.nio.channels.WritableByteChannel;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import org.apache.tools.zip.ZipEntry;
import org.apache.tools.zip.ZipOutputStream;

public class JjgFbgcCommonUtils {


    /**
     *
     * @param response
     * @param zipName 压缩包的名字
     * @param list 文件名
     */
    public static void batchDownloadFile(HttpServletRequest request, HttpServletResponse response,String zipName,List list,String filepath) throws UnsupportedEncodingException {
        //设置压缩包的名字
        String downloadName = URLEncoder.encode(zipName+".zip", "UTF-8");
        response.reset();
        response.setHeader("Content-disposition", "attachment; filename=" + downloadName);
        response.setContentType("application/zip;charset=utf-8");
        response.setCharacterEncoding("utf-8");

        //设置压缩流：直接写入response，实现边压缩边下载
        ZipOutputStream zipOs = null;
        //循环将文件写入压缩流
        DataOutputStream os = null;
        //文件
        File file;
        try {
            zipOs = new ZipOutputStream(new BufferedOutputStream(response.getOutputStream()));
            //设置压缩方法
            zipOs.setMethod(ZipOutputStream.DEFLATED);
            //遍历文件信息（主要获取文件名/文件路径等）
            for (int i=0;i<list.size();i++) {
                String name = zipName+"-"+list.get(i)+".xlsx";
                String path = filepath+File.separator+name;
                file = new File(path);
                if (!file.exists()) {
                    break;
                }
                //添加ZipEntry，并将ZipEntry写入文件流
                zipOs.putNextEntry(new ZipEntry(name));
                os = new DataOutputStream(zipOs);
                FileInputStream fs = new FileInputStream(file);
                byte[] b = new byte[100];
                int length;
                //读入需要下载的文件的内容，打包到zip文件
                while ((length = fs.read(b)) != -1) {
                    os.write(b, 0, length);
                }
                //关闭流
                fs.close();
                zipOs.closeEntry();

            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            //关闭流
            try {
                if (os != null) {
                    os.flush();
                    os.close();
                }
                if (zipOs != null) {
                    zipOs.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }


    public static void download(HttpServletResponse response, String p, String fileName) throws IOException {
        final Path path = Paths.get(p);
//        设置响应头
        response.setContentType("application/octet-stream");
        response.addHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
//        获取输出流
        final WritableByteChannel writableByteChannel = Channels.newChannel(response.getOutputStream());
//        读取文件流通道
        final FileChannel fileChannel = new FileInputStream(path.toFile()).getChannel();
//        写入数据到通道
        fileChannel.transferTo(0, fileChannel.size(), writableByteChannel);
        fileChannel.close();
        writableByteChannel.close();

    }

    public static ArrayList<String> getNOHiddenSheets(XSSFWorkbook wb){
        ArrayList<String> sheets = new ArrayList<String>();
        /*
         * 要打印的文件肯定都不是隐藏的，所以先选出来所有没有隐藏的sheet
         */
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            if(!wb.isSheetHidden(i)){
                sheets.add(wb.getSheetAt(i).getSheetName());
            }
        }
        return sheets;
    }

    public static ArrayList<String> getHasPrintAreaSheets(XSSFWorkbook wb){
        ArrayList<String> shownsheets;
        ArrayList<String> printsheets = new ArrayList<String>();

        XSSFSheet sheet = null;
        shownsheets = getNOHiddenSheets(wb);
        for (String sheetname : shownsheets) {
            sheet = wb.getSheet(sheetname);
            String pringArea = wb.getPrintArea(wb.getSheetIndex(sheet));
            if(pringArea != null){
                printsheets.add(sheetname);
            }
        }
        return printsheets;
    }

    public static void deleteEmptySheets(XSSFWorkbook wb){
        ArrayList<String> printSheets;
        ArrayList<String> delsheets = new ArrayList<String>();;
        printSheets = getHasPrintAreaSheets(wb);

        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            if(wb.getSheetAt(i).getSheetName().contains("温度修正")){
                continue;
            }
            if(wb.getSheetAt(i).getRow(0)!=null&&wb.getSheetAt(i).getRow(0).getCell(0).getStringCellValue().contains("构造深度质量鉴定表")
                    &&(wb.getSheetAt(i).getSheetName().equals("混凝土桥")||
                    wb.getSheetAt(i).getSheetName().equals("混凝土隧道")||
                    wb.getSheetAt(i).getSheetName().equals("混凝土路面"))){  //判断铺沙法构造深度

                if(wb.getSheetAt(i).getRow(6).getCell(0)!=null&&wb.getSheetAt(i).getRow(6).getCell(0).getStringCellValue().equals(""))
                {
                    delsheets.add(wb.getSheetAt(i).getSheetName());
                    continue;
                }
            }
            else if(wb.getSheetAt(i).getRow(0)!=null&&wb.getSheetAt(i).getRow(0).getCell(0).getStringCellValue().contains("构造深度质量鉴定表"))
            {
                if(wb.getSheetAt(i).getRow(1).getCell(2).getStringCellValue()==null
                        ||"".equals(wb.getSheetAt(i).getRow(1).getCell(2).getStringCellValue())
                ){
                    delsheets.add(wb.getSheetAt(i).getSheetName());
                    continue;
                }
            }

            if(wb.getSheetAt(i).getRow(0)!=null&&wb.getSheetAt(i).getRow(0).getCell(0).getStringCellValue().contains("桥面系构造深度")){  //桥面系构造深度
                if(wb.getSheetAt(i).getRow(6).getCell(1)==null&&wb.getSheetAt(i).getRow(6).getCell(2)==null)
                {
                    delsheets.add(wb.getSheetAt(i).getSheetName());
                    continue;
                }else if(wb.getSheetAt(i).getRow(6).getCell(1)!=null&&wb.getSheetAt(i).getRow(6).getCell(1).getCellType()!=1&&wb.getSheetAt(i).getRow(6).getCell(1).getNumericCellValue()==0&&
                        wb.getSheetAt(i).getRow(6).getCell(2)!=null&&wb.getSheetAt(i).getRow(6).getCell(2).getCellType()!=1&&wb.getSheetAt(i).getRow(6).getCell(2).getNumericCellValue()==0) {
                    delsheets.add(wb.getSheetAt(i).getSheetName());
                    continue;
                }else if(wb.getSheetAt(i).getRow(6).getCell(1)!=null&&wb.getSheetAt(i).getRow(6).getCell(1).getCellType()==1&&wb.getSheetAt(i).getRow(6).getCell(1).getStringCellValue().equals("")&&
                        wb.getSheetAt(i).getRow(6).getCell(2)!=null&&wb.getSheetAt(i).getRow(6).getCell(2).getCellType()==1&&wb.getSheetAt(i).getRow(6).getCell(2).getStringCellValue().equals("")) {
                    delsheets.add(wb.getSheetAt(i).getSheetName());
                    continue;
                }
            }
            if(wb.getSheetAt(i).getRow(0)!=null&&wb.getSheetAt(i).getRow(0).getCell(0).getStringCellValue().contains("桥面系摩擦系数")){  //桥面系构造深度
                if(wb.getSheetAt(i).getRow(6).getCell(3)==null&&wb.getSheetAt(i).getRow(6).getCell(4)==null)
                {
                    delsheets.add(wb.getSheetAt(i).getSheetName());
                    continue;
                }else if(wb.getSheetAt(i).getRow(6).getCell(3)!=null&&wb.getSheetAt(i).getRow(6).getCell(3).getCellType()!=1&&wb.getSheetAt(i).getRow(6).getCell(3).getNumericCellValue()==0&&
                        wb.getSheetAt(i).getRow(6).getCell(4)!=null&&wb.getSheetAt(i).getRow(6).getCell(4).getCellType()!=1&&wb.getSheetAt(i).getRow(6).getCell(4).getNumericCellValue()==0) {
                    delsheets.add(wb.getSheetAt(i).getSheetName());
                    continue;
                }else if(wb.getSheetAt(i).getRow(6).getCell(3)!=null&&wb.getSheetAt(i).getRow(6).getCell(3).getCellType()==1&&wb.getSheetAt(i).getRow(6).getCell(3).getStringCellValue().equals("")&&
                        wb.getSheetAt(i).getRow(6).getCell(4)!=null&&wb.getSheetAt(i).getRow(6).getCell(4).getCellType()==1&&wb.getSheetAt(i).getRow(6).getCell(4).getStringCellValue().equals("")) {
                    delsheets.add(wb.getSheetAt(i).getSheetName());
                    continue;
                }
            }
            if(wb.getSheetAt(i).getRow(0)!=null&&wb.getSheetAt(i).getRow(0).getCell(0).getStringCellValue().contains("车辙"))
            {
                if(wb.getSheetAt(i).getRow(1).getCell(2).getStringCellValue()==null
                        ||"".equals(wb.getSheetAt(i).getRow(1).getCell(2).getStringCellValue())){
                    delsheets.add(wb.getSheetAt(i).getSheetName());
                    continue;
                }
                if(wb.getSheetAt(i).getRow(6).getCell(0)==null
                        ||
                        (wb.getSheetAt(i).getRow(6).getCell(0).getCellType()==1&&wb.getSheetAt(i).getRow(6).getCell(0).getStringCellValue().equals(""))
                        ||
                        (wb.getSheetAt(i).getRow(6).getCell(1).getCellType()==1&&wb.getSheetAt(i).getRow(6).getCell(1).getStringCellValue().equals("")))
                {
                    delsheets.add(wb.getSheetAt(i).getSheetName());
                    continue;
                }
            }
            if(wb.getSheetAt(i).getRow(0)!=null&&wb.getSheetAt(i).getRow(0).getCell(0).getStringCellValue().contains("路面横坡"))
            {
                if(wb.getSheetAt(i).getRow(1).getCell(2)==null
                        ||"".equals(wb.getSheetAt(i).getRow(1).getCell(2).getStringCellValue())){
                    delsheets.add(wb.getSheetAt(i).getSheetName());
                    continue;
                }
                if(wb.getSheetAt(i).getRow(6).getCell(0)==null||
                        wb.getSheetAt(i).getRow(6).getCell(0).getStringCellValue().equals("")){
                    delsheets.add(wb.getSheetAt(i).getSheetName());
                    continue;
                }
            }
            if(wb.getSheetAt(i).getRow(0)!=null&&wb.getSheetAt(i).getRow(0).getCell(0).getStringCellValue().contains("钻芯法"))//&&wb.getSheetAt(i).getSheetName().equals("连接线")
            {
                if(wb.getSheetAt(i).getRow(1).getCell(1)==null
                        ||"".equals(wb.getSheetAt(i).getRow(1).getCell(1).getStringCellValue())){
                    delsheets.add(wb.getSheetAt(i).getSheetName());
                    continue;
                }
            }
            if(wb.getSheetAt(i).getRow(0)!=null&&wb.getSheetAt(i).getRow(0).getCell(0).getStringCellValue().contains("渗水系数"))//&&wb.getSheetAt(i).getSheetName().equals("连接线")
            {
                if(wb.getSheetAt(i).getRow(6).getCell(0)==null
                        ||"".equals(wb.getSheetAt(i).getRow(6).getCell(0).getStringCellValue())){
                    delsheets.add(wb.getSheetAt(i).getSheetName());
                    continue;
                }
            }
            if(wb.isSheetHidden(i)){
                continue;
            }
            if(!printSheets.contains(wb.getSheetAt(i).getSheetName())){
                delsheets.add(wb.getSheetAt(i).getSheetName());
            }
        }
        for (int i = 0; i < delsheets.size(); i++) {
            wb.removeSheetAt(wb.getSheetIndex(delsheets.get(i)));
        }
    }

    /**
     * 获取砼强度鉴定表中的坚定结果
     * @param map
     * @return
     * @throws IOException
     */
    public static List<Map<String,Object>> gettqdjcjg(Map<String,Object> map) throws IOException {
        List<Map<String,Object>> mapList = new ArrayList<>();
        Map<String,Object> jgmap = new HashMap<>();
        DecimalFormat df = new DecimalFormat(".00");
        DecimalFormat decf = new DecimalFormat("0.##");
        //创建工作簿
        XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(map.get("f").toString()));
        //读取工作表
        XSSFSheet slSheet = xwb.getSheet(map.get("sheetname").toString());
        slSheet.getRow(2).getCell(34).setCellType(CellType.STRING);
        slSheet.getRow(2).getCell(35).setCellType(CellType.STRING);
        slSheet.getRow(2).getCell(36).setCellType(CellType.STRING);
        String bt = slSheet.getRow(0).getCell(0).getStringCellValue();//混凝土强度质量鉴定表（回弹法）
        //String bt1 = bt.getStringCellValue();
        String xmname = slSheet.getRow(1).getCell(2).getStringCellValue();//陕西高速
        //String xmname1 = xmname.getStringCellValue();
        String htdname = slSheet.getRow(1).getCell(29).getStringCellValue();//LJ-1
        //String htdname1 = htdname.getStringCellValue();
        String hd = slSheet.getRow(2).getCell(2).getStringCellValue();//涵洞
        //String hd1 = hd.getStringCellValue();
        if(slSheet != null){
            if(map.get("proname").toString().equals(xmname) && map.get("title").toString().equals(bt) && map.get("htd").toString().equals(htdname) && map.get("fbgc").toString().equals(hd)){
                double zds= Double.valueOf(slSheet.getRow(2).getCell(34).getStringCellValue());
                double hgds= Double.valueOf(slSheet.getRow(2).getCell(35).getStringCellValue());
                double hgl= Double.valueOf(slSheet.getRow(2).getCell(36).getStringCellValue());
                String zdsz = decf.format(zds);
                String hgdsz = decf.format(hgds);
                String hglz = df.format(hgl);
                jgmap.put("总点数",zdsz);
                jgmap.put("合格点数",hgdsz);
                jgmap.put("合格率",hglz);
                mapList.add(jgmap);
            }else {
                return null;
            }
        }
        return mapList;
    }


    /**
     * 获取断面尺寸鉴定表中的检测结果
     * @param map
     * @return
     * @throws IOException
     */

    public static List<Map<String,Object>> getdmcjjcjg(Map<String,Object> map) throws IOException {
        List<Map<String,Object>> mapList = new ArrayList<>();
        Map<String,Object> jgmap = new HashMap<>();
        DecimalFormat df = new DecimalFormat(".00");
        DecimalFormat decf = new DecimalFormat("0.##");
        //创建工作簿
        XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(map.get("f").toString()));
        //读取工作表
        XSSFSheet slSheet = xwb.getSheet(map.get("sheetname").toString());
        XSSFCell bt = slSheet.getRow(0).getCell(0);//混凝土强度质量鉴定表（回弹法）
        XSSFCell xmname = slSheet.getRow(1).getCell(1);//陕西高速
        XSSFCell htdname = slSheet.getRow(1).getCell(6);//LJ-1
        XSSFCell hd = slSheet.getRow(2).getCell(1);//涵洞
        if(slSheet != null){
            if(map.get("proname").toString().equals(xmname.toString()) && map.get("title").toString().equals(bt.toString()) && map.get("htd").toString().equals(htdname.toString()) && map.get("fbgc").toString().equals(hd.toString())){
                //获取到最后一行
                int lastRowNum = slSheet.getLastRowNum();
                slSheet.getRow(lastRowNum-1).getCell(1).setCellType(CellType.STRING);
                slSheet.getRow(lastRowNum-1).getCell(3).setCellType(CellType.STRING);
                slSheet.getRow(lastRowNum-1).getCell(6).setCellType(CellType.STRING);
                double zds= Double.valueOf(slSheet.getRow(lastRowNum-1).getCell(1).getStringCellValue());
                double hgds= Double.valueOf(slSheet.getRow(lastRowNum-1).getCell(3).getStringCellValue());
                double hgl= Double.valueOf(slSheet.getRow(lastRowNum-1).getCell(6).getStringCellValue());
                String zdsz = decf.format(zds);
                String hgdsz = decf.format(hgds);
                String hglz = df.format(hgl);
                jgmap.put("检测总点数",zdsz);
                jgmap.put("合格点数",hgdsz);
                jgmap.put("合格率",hglz);
                mapList.add(jgmap);
            }else {
                return null;
            }
        }
        return mapList;
    }

    /**
     * 取出数据中最新的检测时间
     * @param date1
     * @param date2
     * @return
     * @throws ParseException
     */
    public static String getLastDate(String date1, String date2) throws ParseException {
        DateFormat df = new SimpleDateFormat("yyyy.MM.dd");
        Date dt1 = df.parse(date1);
        Date dt2 = df.parse(date2);
        if (dt1.getTime() > dt2.getTime()) {
            return date1;
        } else if (dt1.getTime() < dt2.getTime()){
            return date2;
        }
        return date1;
    }


    public static XSSFCellStyle dBtoExcelUtils(XSSFWorkbook wb){
        XSSFCellStyle cellstyle = wb.createCellStyle();
        XSSFFont font=wb.createFont();
        font.setFontHeightInPoints((short)10);
        font.setFontName("宋体");
        cellstyle.setFont(font);
        cellstyle.setBorderBottom(BorderStyle.THIN); //下边框
        cellstyle.setBorderLeft(BorderStyle.THIN);//左边框
        cellstyle.setBorderTop(BorderStyle.THIN);//上边框
        cellstyle.setBorderRight(BorderStyle.THIN);//右边框
        cellstyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直
        cellstyle.setAlignment(HorizontalAlignment.CENTER);//水平
        XSSFDataFormat format= wb.createDataFormat();
        cellstyle.setDataFormat(format.getFormat("C##0"));
        return cellstyle;
    }
    public static XSSFCellStyle tsfCellStyle(XSSFWorkbook wb){
        XSSFCellStyle cellstyle = wb.createCellStyle();
        XSSFFont font=wb.createFont();
        font.setFontHeightInPoints((short)11);
        font.setFontName("宋体");
        cellstyle.setFont(font);
        cellstyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直
        cellstyle.setAlignment(HorizontalAlignment.CENTER);//水平
        XSSFDataFormat format= wb.createDataFormat();
        cellstyle.setDataFormat(format.getFormat("C##0"));
        return cellstyle;
    }

    public static String getAllowError1(String allowerror){
        if(allowerror.contains("±")){
            return allowerror.replaceAll("[^0-9]", "");
        }
        else{
            return allowerror.split(",")[0].replaceAll("[^0-9]", "");
        }
    }
    public static String getAllowError2(String allowerror){
        if(allowerror.contains("±")){
            return allowerror.replaceAll("[^0-9]", "");
        }
        else{
            return allowerror.split(",")[1].replaceAll("[^0-9]", "");
        }
    }

    public static String getAllowError3(String allowerror){
        return allowerror;
    }

    public static void updateFormula(XSSFWorkbook wb, XSSFSheet sheet) {
        for(int row = sheet.getFirstRowNum() ; row < sheet.getPhysicalNumberOfRows() ; row++){
            XSSFRow r = sheet.getRow(row);
            XSSFCell c = null;
            XSSFFormulaEvaluator eval = null;
            eval = new XSSFFormulaEvaluator((XSSFWorkbook) wb);
            for (int i = r.getFirstCellNum(); i < r.getLastCellNum(); i++) {
                c = r.getCell(i);
                try{
                    if (c.getCellType() == XSSFCell.CELL_TYPE_FORMULA)
                        eval.evaluateFormulaCell(c);
                    //c.setCellFormula(null);
                }
                catch(Exception e){
                    continue;
                }
            }
        }
    }
}
