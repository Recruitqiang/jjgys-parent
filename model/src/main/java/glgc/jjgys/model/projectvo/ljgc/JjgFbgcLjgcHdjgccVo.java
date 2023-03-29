package glgc.jjgys.model.projectvo.ljgc;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.alibaba.excel.metadata.BaseRowModel;
import com.fasterxml.jackson.annotation.JsonFormat;
import lombok.Data;

import java.util.Date;

/**
 * <p>
 * 
 * </p>
 *
 * @author wq
 * @since 2023-01-14
 */
@Data
@HeadStyle(fillForegroundColor = 13)
@HeadFontStyle(fontHeightInPoints = 10)
@HeadRowHeight(30)
public class JjgFbgcLjgcHdjgccVo extends BaseRowModel {

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
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;


}
