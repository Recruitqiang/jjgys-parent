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
import java.util.Date;

/**
 * <p>
 * 
 * </p>
 *
 * @author wq
 * @since 2023-03-20
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_qlgc_sb_bhchd")
public class JjgFbgcQlgcSbBhchdVo extends BaseRowModel {

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
     * 构件编号及检测部位
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "构件编号及检测部位" ,index = 2)
    private String gjbhjjcbw;

    /**
     * 钢筋直径(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "钢筋直径(mm)" ,index = 3)
    private String gjzj;

    /**
     * 设计值(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "设计值(mm)" ,index = 4)
    private String sjz;

    /**
     * 实测值1(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "实测值1(mm)" ,index = 5)
    private String scz1;

    @ColumnWidth(23)
    @ExcelProperty(value = "实测值2(mm)" ,index = 6)
    private String scz2;

    @ColumnWidth(23)
    @ExcelProperty(value = "实测值3(mm)" ,index = 7)
    private String scz3;

    @ColumnWidth(23)
    @ExcelProperty(value = "实测值4(mm)" ,index = 8)
    private String scz4;

    @ColumnWidth(23)
    @ExcelProperty(value = "实测值5(mm)" ,index = 9)
    private String scz5;

    @ColumnWidth(23)
    @ExcelProperty(value = "实测值6(mm)" ,index = 10)
    private String scz6;

    @ColumnWidth(23)
    @ExcelProperty(value = "修正值" ,index = 11)
    private String xzz;

    /**
     * 允许偏差+(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "允许偏差+(mm)" ,index = 12)
    private String yxwcz;

    /**
     * 允许偏差-(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "允许偏差-(mm)" ,index = 13)
    private String yxwcf;


}
