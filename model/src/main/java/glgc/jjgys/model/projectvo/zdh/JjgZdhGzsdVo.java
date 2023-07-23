package glgc.jjgys.model.projectvo.zdh;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.alibaba.excel.metadata.BaseRowModel;
import com.baomidou.mybatisplus.annotation.TableField;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import lombok.Data;
import lombok.EqualsAndHashCode;

import java.io.Serializable;
import java.time.LocalDate;

/**
 * <p>
 * 
 * </p>
 *
 * @author wq
 * @since 2023-06-15
 */
@Data
@HeadStyle(fillForegroundColor = 13)
@HeadFontStyle(fontHeightInPoints = 10)
@HeadRowHeight(30)
public class JjgZdhGzsdVo extends BaseRowModel {

    /**
     * 起点桩号
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "起点桩号" ,index = 0)
    private String qdzh;

    /**
     * 终点桩号
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "终点桩号" ,index = 1)
    private String zdzh;

    /**
     * mtd
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "代表MTD" ,index = 2)
    private String mtd;


    /**
     * 类型标识
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "路线类型" ,index = 3)
    private String lxbs;

    /**
     * 匝道标识
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "匝道标识" ,index = 4)
    private String zdbs;



}
