package glgc.jjgys.model.projectvo.qlgc;

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
import java.time.LocalDate;
import java.util.Date;

/**
 * <p>
 * 
 * </p>
 *
 * @author wq
 * @since 2023-03-22
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_qlgc_xb_szd")
public class JjgFbgcQlgcXbSzdVo extends BaseRowModel {

    /**
     * 检测时间
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "检测日期" ,index = 0)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;

    /**
     * 桥梁名称
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "桥梁名称" ,index = 1)
    private String qlmc;

    /**
     * 墩台号
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "墩台号" ,index = 2)
    private String dth;

    /**
     * 墩柱高度
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "墩柱高度" ,index = 3)
    private String dzgd;


    /**
     * 横向实测值
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "横向实测值" ,index = 4)
    private String hxscz;

    /**
     * 纵向实测值(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "纵向实测值(mm)" ,index = 5)
    private String zxscz;


}
