package glgc.jjgys.model.projectvo.ljgc;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.alibaba.excel.metadata.BaseRowModel;
import com.fasterxml.jackson.annotation.JsonFormat;
import lombok.Data;

import java.util.Date;

/**
 * <p>
 * 
 * </p>
 *
 * @author wq
 * @since 2023-02-21
 */
@Data
@HeadStyle(fillForegroundColor = 13)
@HeadFontStyle(fontHeightInPoints = 10)
@HeadRowHeight(30)
public class JjgFbgcLjgcLjtsfysdHtVo extends BaseRowModel {


    @ColumnWidth(23)
    @ExcelProperty(value = "路基压实度_灰土规定值" ,index = 0)
    private String htgdz;

    @ColumnWidth(23)
    @ExcelProperty(value = "试验时间" ,index = 1)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date sysj;

    @ColumnWidth(23)
    @ExcelProperty(value = "桩号" ,index = 2)
    private String zh;

    @ColumnWidth(23)
    @ExcelProperty(value = "结构层次" ,index = 3)
    private String jgcc;

    @ColumnWidth(23)
    @ExcelProperty(value = "结构类型" ,index = 4)
    private String jglx;

    @ColumnWidth(23)
    @ExcelProperty(value = "最大干密度" ,index = 5)
    private String zdgmd;

    @ColumnWidth(23)
    @ExcelProperty(value = "标准砂密度" ,index = 6)
    private String bzsmd;

    @ColumnWidth(23)
    @ExcelProperty(value = "取样桩号及位置" ,index = 7)
    private String qyzhjwz;

    @ColumnWidth(23)
    @ExcelProperty(value = "试坑深度(cm)" ,index = 8)
    private String sksd;

    @ColumnWidth(23)
    @ExcelProperty(value = "锥体及基板和表面间砂质量（g）" ,index = 9)
    private String ztjjbhbmjszl;

    @ColumnWidth(23)
    @ExcelProperty(value = "灌砂前筒+砂质量（g）" ,index = 10)
    private String gsqtszl;

    @ColumnWidth(23)
    @ExcelProperty(value = "灌砂后筒+砂质量（g）" ,index = 11)
    private String gshtszl;

    @ColumnWidth(23)
    @ExcelProperty(value = "试样质量（g）" ,index = 12)
    private String syzl;

    @ColumnWidth(23)
    @ExcelProperty(value = "盒号" ,index = 13)
    private String hh;

    @ColumnWidth(23)
    @ExcelProperty(value = "盒质量（g）" ,index = 14)
    private String hzl;

    @ColumnWidth(23)
    @ExcelProperty(value = "盒+湿试样质量（g）" ,index = 15)
    private String hsshzl;

    @ColumnWidth(23)
    @ExcelProperty(value = "盒+干试样质量（g）" ,index = 16)
    private String hgsyzl;

    @ColumnWidth(23)
    @ExcelProperty(value = "序号" ,index = 17)
    private String xh;


}
