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
 * 桥面构造深度
 * </p>
 *
 * @author wq
 * @since 2023-03-20
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_qlgc_qmgzsd")
public class JjgFbgcQlgcQmgzsdVo extends BaseRowModel {


    /**
     * 检测时间
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "检测日期" ,index = 0)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;

    /**
     * 结构名称
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "结构名称" ,index = 1)
    private String jgmc;

    /**
     * ABM
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "ABM" ,index = 2)
    private String abm;

    /**
     * 桩号
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "桩号" ,index = 3)
    private String zh;

    /**
     * 车道
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "车道" ,index = 4)
    private String cd;

    /**
     * 设计最小值
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "设计最小值" ,index = 5)
    private String sjzxz;

    /**
     * 设计最大值
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "设计最大值" ,index = 6)
    private String sjzdz;

    /**
     * 测点1D1(㎜)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "测点1D1(㎜)" ,index = 7)
    private String cd1d1;

    /**
     * 测点1D2(㎜)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "测点1D2(㎜)" ,index = 8)
    private String cd1d2;

    /**
     * 测点2D1(㎜)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "测点2D1(㎜)" ,index = 9)
    private String cd2d1;

    /**
     * 测点2D2(㎜)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "测点2D2(㎜)" ,index = 10)
    private String cd2d2;

    /**
     * 测点3D1(㎜)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "测点3D1(㎜)" ,index = 11)
    private String cd3d1;

    /**
     * 测点3D2(㎜)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "测点3D2(㎜)" ,index = 12)
    private String cd3d2;


}
