package glgc.jjgys.model.projectvo.jagc;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
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
 * @since 2023-03-01
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_jtaqss_jabxfhl")
public class JjgFbgcJtaqssJabxfhlVo extends BaseRowModel {

    @ColumnWidth(23)
    @ExcelProperty(value = "检测日期" ,index = 0)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;

    @ColumnWidth(23)
    @ExcelProperty(value = "位置及类型" ,index = 1)
    private String wzjlx;

    @ColumnWidth(23)
    @ExcelProperty(value = "基底厚度规定值(mm)" ,index = 2)
    private String jdhdgdz;

    @ColumnWidth(23)
    @ExcelProperty(value = "基底厚度实测值1(mm)" ,index = 3)
    private String jdhdscz1;

    @ColumnWidth(23)
    @ExcelProperty(value = "基底厚度实测值2(mm)" ,index = 4)
    private String jdhdscz2;

    @ColumnWidth(23)
    @ExcelProperty(value = "基底厚度实测值3(mm)" ,index = 5)
    private String jdhdscz3;

    @ColumnWidth(23)
    @ExcelProperty(value = "立柱壁厚规定值(mm)" ,index = 6)
    private String lzbhgdz;

    @ColumnWidth(23)
    @ExcelProperty(value = "立柱壁厚实测值1(mm)" ,index = 7)
    private String lzbhscz1;

    @ColumnWidth(23)
    @ExcelProperty(value = "立柱壁厚实测值2(mm)" ,index = 8)
    private String lzbhscz2;

    @ColumnWidth(23)
    @ExcelProperty(value = "立柱壁厚实测值3(mm)" ,index = 9)
    private String lzbhscz3;

    @ColumnWidth(23)
    @ExcelProperty(value = "中心高度规定值(mm)" ,index = 10)
    private String zxgdgdz;

    @ColumnWidth(23)
    @ExcelProperty(value = "中心高度允许偏差+(mm)" ,index = 11)
    private String zxgdyxpsz;

    @ColumnWidth(23)
    @ExcelProperty(value = "中心高度允许偏差-(mm)" ,index = 12)
    private String zxgdyxpsf;

    @ColumnWidth(23)
    @ExcelProperty(value = "中心高度实测值1(mm)" ,index = 13)
    private String zxgdscz1;

    @ColumnWidth(23)
    @ExcelProperty(value = "中心高度实测值2(mm)" ,index = 14)
    private String zxgdscz2;

    @ColumnWidth(23)
    @ExcelProperty(value = "中心高度实测值3(mm)" ,index = 15)
    private String zxgdscz3;

    @ColumnWidth(23)
    @ExcelProperty(value = "埋入深度规定值(mm)" ,index = 16)
    private String mrsdgdz;

    @ColumnWidth(23)
    @ExcelProperty(value = "埋入深度实测值(mm)" ,index = 17)
    private String mrsdscz;



}
