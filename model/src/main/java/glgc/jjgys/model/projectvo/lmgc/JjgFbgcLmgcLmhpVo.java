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
 * @since 2023-04-22
 */
@Data
@HeadStyle(fillForegroundColor = 13)
@HeadFontStyle(fontHeightInPoints = 10)
@HeadRowHeight(30)
public class JjgFbgcLmgcLmhpVo extends BaseRowModel {


    /**
     * 检测时间
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "检测时间" ,index = 0)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;

    /**
     * 路面类型
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "路面类型" ,index = 1)
    private String lmlx;

    /**
     * Z/Y
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "Z/Y" ,index = 2)
    private String zy;

    /**
     * 桩号
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "桩号" ,index = 3)
    private String zh;

    /**
     * 位置
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "位置" ,index = 4)
    private String wz;

    /**
     * 前视读数(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "前视读数(mm)" ,index = 5)
    private String qsds;

    /**
     * 后视读数(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "后视读数(mm)" ,index = 6)
    private String hsds;

    /**
     * 长(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "长(mm)" ,index = 7)
    private String length;

    /**
     * 设计值(%)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "设计值(%)" ,index = 8)
    private String sjz;

    /**
     * 允许偏差(%)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "允许偏差(%)" ,index = 9)
    private String yxps;

    /**
     * 路线类型
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "路线类型" ,index = 10)
    private String lxlx;



}
