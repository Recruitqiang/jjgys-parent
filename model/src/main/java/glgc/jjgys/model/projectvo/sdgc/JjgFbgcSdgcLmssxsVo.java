package glgc.jjgys.model.projectvo.sdgc;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.alibaba.excel.metadata.BaseRowModel;
import com.baomidou.mybatisplus.annotation.TableField;
import com.fasterxml.jackson.annotation.JsonFormat;
import lombok.Data;

import java.util.Date;

/**
 * <p>
 * 
 * </p>
 *
 * @author wq
 * @since 2023-05-04
 */
@Data
@HeadStyle(fillForegroundColor = 13)
@HeadFontStyle(fontHeightInPoints = 10)
@HeadRowHeight(30)
public class JjgFbgcSdgcLmssxsVo extends BaseRowModel {

    /**
     * 检测时间
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "检测时间" ,index = 0)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;

    /**
     * 隧道名称
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "隧道名称" ,index = 1)
    private String sdmc;

    /**
     * 路、匝道、隧道
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "路、匝道、隧道" ,index = 2)
    private String lzdsd;

    /**
     * 桩号
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "桩号" ,index = 3)
    private String zh;

    /**
     * 初读数(mL)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "初读数(mL)" ,index = 4)
    private String cds;

    /**
     * 第一分钟读数(mL)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "第一分钟读数(mL)" ,index = 5)
    private String ofzds;

    /**
     * 第二分钟读数(mL)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "第二分钟读数(mL)" ,index = 6)
    private String tfzds;

    /**
     * 水量(mL)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "水量(mL)" ,index = 7)
    private String sl;

    /**
     * 时间(s)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "时间(s)" ,index = 8)
    private String sj;

    /**
     * 渗水系数规定值
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "渗水系数规定值" ,index = 9)
    private String ssxsgdz;



}
