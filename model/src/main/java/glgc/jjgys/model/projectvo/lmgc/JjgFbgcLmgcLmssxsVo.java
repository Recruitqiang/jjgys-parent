package glgc.jjgys.model.projectvo.lmgc;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.alibaba.excel.metadata.BaseRowModel;
import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableField;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import com.fasterxml.jackson.annotation.JsonFormat;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;
import lombok.EqualsAndHashCode;

import java.io.Serializable;
import java.util.Date;

/**
 * <p>
 * 
 * </p>
 *
 * @author wq
 * @since 2023-04-19
 */
@Data
@HeadStyle(fillForegroundColor = 13)
@HeadFontStyle(fontHeightInPoints = 10)
@HeadRowHeight(30)
public class JjgFbgcLmgcLmssxsVo extends BaseRowModel {


    /**
     * 检测时间
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "检测时间" ,index = 0)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;

    /**
     * 路线类型
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "路、匝道、隧道" ,index = 1)
    private String lxlx;

    /**
     * 桩号
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "桩号" ,index = 2)
    private String zh;

    /**
     * 初读数(mL)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "初读数(mL)" ,index = 3)
    private String cds;

    /**
     * 第一分钟读数(mL)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "第一分钟读数(mL)" ,index = 4)
    private String ofzds;

    /**
     * 第二分钟读数(mL)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "第二分钟读数(mL)" ,index = 5)
    private String tfzds;

    /**
     * 水量(mL)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "水量(mL)" ,index = 6)
    private String sl;

    /**
     * 时间(s)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "时间(s)" ,index = 7)
    private String sj;

    /**
     * 渗水系数规定值
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "渗水系数规定值" ,index = 8)
    private String ssxsgdz;



}
