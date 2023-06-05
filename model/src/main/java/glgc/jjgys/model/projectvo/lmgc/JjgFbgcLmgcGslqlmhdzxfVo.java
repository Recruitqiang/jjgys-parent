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
 * @since 2023-04-25
 */
@Data
@HeadStyle(fillForegroundColor = 13)
@HeadFontStyle(fontHeightInPoints = 10)
@HeadRowHeight(30)
public class JjgFbgcLmgcGslqlmhdzxfVo extends BaseRowModel {

    /**
     * 检测时间
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "检测时间" ,index = 0)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;

    /**
     * 路、桥、隧、匝道
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "路、桥、隧、匝道" ,index = 1)
    private String lx;

    /**
     * 桩号
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "桩号" ,index = 2)
    private String zh;

    /**
     * 上面层测值1(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "上面层测值1(mm)" ,index = 3)
    private String smccz1;

    /**
     * 上面层测值2(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "上面层测值2(mm)" ,index = 4)
    private String smccz2;

    /**
     * 上面层测值3(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "上面层测值3(mm)" ,index = 5)
    private String smccz3;

    /**
     * 上面层测值4(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "上面层测值4(mm)" ,index = 6)
    private String smccz4;

    /**
     * 上面层设计值(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "上面层设计值(mm)" ,index = 7)
    private String smcsjz;

    /**
     * 总厚度测值1(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "总厚度测值1(mm)" ,index = 8)
    private String zhdcz1;

    /**
     * 总厚度测值2(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "总厚度测值2(mm)" ,index = 9)
    private String zhdcz2;

    /**
     * 总厚度测值3(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "总厚度测值3(mm)" ,index = 10)
    private String zhdcz3;

    /**
     * 总厚度测值4(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "总厚度测值4(mm)" ,index = 11)
    private String zhdcz4;

    /**
     * 总厚度设计值(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "总厚度设计值(mm)" ,index = 12)
    private String zhdsjz;



}
