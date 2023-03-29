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
 * @since 2023-03-20
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_qlgc_qmpzd")
public class JjgFbgcQlgcQmpzdVo extends BaseRowModel {


    /**
     * 检测时间
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "检测日期" ,index = 0)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;

    /**
     * 桥名
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "桥名" ,index = 1)
    private String qm;

    /**
     * 位置
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "位置" ,index = 2)
    private String wz;

    /**
     * 设计值
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "设计值" ,index = 3)
    private String sjz;

    /**
     * 实测值
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "实测值" ,index = 4)
    private String scz;


}
