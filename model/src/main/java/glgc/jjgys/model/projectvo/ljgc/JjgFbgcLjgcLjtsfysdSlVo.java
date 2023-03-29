package glgc.jjgys.model.projectvo.ljgc;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.alibaba.excel.metadata.BaseRowModel;
import com.baomidou.mybatisplus.annotation.TableField;
import com.fasterxml.jackson.annotation.JsonFormat;
import io.swagger.annotations.ApiModelProperty;
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
public class JjgFbgcLjgcLjtsfysdSlVo extends BaseRowModel {


    @ColumnWidth(23)
    @ExcelProperty(value = "路基压实度_砂砾规定值" ,index = 0)
    private String slgdz;

    @ColumnWidth(23)
    @ExcelProperty(value = "试验时间" ,index = 1)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date sysj;

    @ColumnWidth(23)
    @ExcelProperty(value = "桩号" ,index = 2)
    private String zh;

    @ColumnWidth(23)
    @ExcelProperty(value = "a" ,index = 3)
    private String a;

    @ColumnWidth(23)
    @ExcelProperty(value = "b" ,index = 4)
    private String b;

    @ColumnWidth(23)
    @ExcelProperty(value = "c" ,index = 5)
    private String c;

    @ColumnWidth(23)
    @ExcelProperty(value = "取样桩号及位置" ,index = 7)
    private String qyzhjwz;

    @ColumnWidth(23)
    @ExcelProperty(value = "试坑深度(cm)" ,index = 8)
    private String sksd;


    @ColumnWidth(23)
    @ExcelProperty(value = "灌砂前筒+砂质量（g）" ,index = 10)
    private String gsqtszl;

    @ColumnWidth(23)
    @ExcelProperty(value = "灌砂后筒+砂质量（g）" ,index = 11)
    private String gshtszl;

    @ColumnWidth(23)
    @ExcelProperty(value = "锥体砂重" ,index = 12)
    private String ztsz;

    private String lsdmd;

    private String hhldszl;

    private String hh;

    private String hgzl;

    private String hzl;

    private String klzl;

    private String xh;



}
