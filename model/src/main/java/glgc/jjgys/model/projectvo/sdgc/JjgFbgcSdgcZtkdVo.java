package glgc.jjgys.model.projectvo.sdgc;

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
 * @since 2023-03-26
 */
@Data
@HeadStyle(fillForegroundColor = 13)
@HeadFontStyle(fontHeightInPoints = 10)
@HeadRowHeight(30)
public class JjgFbgcSdgcZtkdVo extends BaseRowModel {

    /**
     * 检测时间
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "检测日期" ,index = 0)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;

    /**
     * 隧道名称
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "隧道名称" ,index = 1)
    private String sdmc;

    /**
     * 桩号
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "桩号" ,index = 2)
    private String zh;

    /**
     * 左边宽(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "左边宽(mm)" ,index = 3)
    private String zbk;

    /**
     * 右边宽(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "右边宽(mm)" ,index = 4)
    private String ybk;

    /**
     * 设计值(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "设计值(mm)" ,index = 5)
    private String sjz;


}
