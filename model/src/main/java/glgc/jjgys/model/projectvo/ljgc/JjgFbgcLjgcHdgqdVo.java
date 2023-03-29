package glgc.jjgys.model.projectvo.ljgc;

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
import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;

import java.io.Serializable;
import java.util.Date;

/**
 * <p>
 * 
 * </p>
 *
 * @author wq
 * @since 2023-01-14
 */
@Data
@HeadStyle(fillForegroundColor = 13)
@HeadFontStyle(fontHeightInPoints = 10)
@HeadRowHeight(30)
public class JjgFbgcLjgcHdgqdVo extends BaseRowModel {


    @ColumnWidth(23)
    @ExcelProperty(value = "检测时间" ,index = 0)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;

    @ColumnWidth(23)
    @ExcelProperty(value = "桩号" ,index = 1)
    private String zh;

    @ColumnWidth(23)
    @ExcelProperty(value = "部位1" ,index = 2)
    private String bw1;

    @ColumnWidth(23)
    @ExcelProperty(value = "部位2" ,index = 3)
    private String bw2;

    @ColumnWidth(23)
    @ExcelProperty(value = "测定值1" ,index = 4)
    private String cdz1;

    @ColumnWidth(23)
    @ExcelProperty(value = "值2" ,index = 5)
    private String z2;

    @ColumnWidth(23)
    @ExcelProperty(value = "值3" ,index = 6)
    private String z3;

    @ColumnWidth(23)
    @ExcelProperty(value = "值4" ,index = 7)
    private String z4;

    @ColumnWidth(23)
    @ExcelProperty(value = "值5" ,index = 8)
    private String z5;

    @ColumnWidth(23)
    @ExcelProperty(value = "值6" ,index = 9)
    private String z6;

    @ColumnWidth(23)
    @ExcelProperty(value = "值7" ,index = 10)
    private String z7;

    @ColumnWidth(23)
    @ExcelProperty(value = "值8" ,index = 11)
    private String z8;

    @ColumnWidth(23)
    @ExcelProperty(value = "值9" ,index = 12)
    private String z9;

    @ColumnWidth(23)
    @ExcelProperty(value = "值10" ,index = 13)
    private String z10;

    @ColumnWidth(23)
    @ExcelProperty(value = "值11" ,index = 14)
    private String z11;

    @ColumnWidth(23)
    @ExcelProperty(value = "值12" ,index = 15)
    private String z12;

    @ColumnWidth(23)
    @ExcelProperty(value = "值13" ,index = 16)
    private String z13;

    @ColumnWidth(23)
    @ExcelProperty(value = "值14" ,index = 17)
    private String z14;

    @ColumnWidth(23)
    @ExcelProperty(value = "值15" ,index = 18)
    private String z15;

    @ColumnWidth(23)
    @ExcelProperty(value = "值16" ,index = 19)
    private String z16;

    @ColumnWidth(23)
    @ExcelProperty(value = "回弹角度" ,index = 20)
    private String htjd;

    @ColumnWidth(23)
    @ExcelProperty(value = "浇筑面" ,index = 21)
    private String jzm;

    @ColumnWidth(23)
    @ExcelProperty(value = "碳化深度" ,index = 22)
    private String thsd;

    @ColumnWidth(23)
    @ExcelProperty(value = "是否泵送" ,index = 23)
    private String sfbs;

    @ColumnWidth(23)
    @ExcelProperty(value = "设计强度" ,index = 24)
    private String sjqd;


}
