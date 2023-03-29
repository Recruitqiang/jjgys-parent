package glgc.jjgys.model.projectvo.jagc;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.metadata.BaseRowModel;
import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableField;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
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
@TableName("jjg_fbgc_jtaqss_jathldmcc")
public class JjgFbgcJtaqssJathldmccVo extends BaseRowModel {

    @ColumnWidth(23)
    @ExcelProperty(value = "桩号" ,index = 0)
    private String zh;

    @ColumnWidth(23)
    @ExcelProperty(value = "部位" ,index = 1)
    private String bw;

    @ColumnWidth(23)
    @ExcelProperty(value = "类别" ,index = 2)
    private String lb;

    @ColumnWidth(23)
    @ExcelProperty(value = "设计值(mm)" ,index = 3)
    private String sjz;

    @ColumnWidth(23)
    @ExcelProperty(value = "实测值(mm)" ,index = 4)
    private String scz;

    @ColumnWidth(23)
    @ExcelProperty(value = "允许误差+(mm)" ,index = 5)
    private String yxwcz;

    @ColumnWidth(23)
    @ExcelProperty(value = "允许误差-(mm)" ,index = 6)
    private String yxwcf;

    @ColumnWidth(23)
    @ExcelProperty(value = "检测日期" ,index = 7)
    private Date jcrq;

}
