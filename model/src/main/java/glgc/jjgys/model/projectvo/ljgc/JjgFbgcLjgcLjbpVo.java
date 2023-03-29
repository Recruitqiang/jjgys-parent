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

@Data
@HeadStyle(fillForegroundColor = 13)
@HeadFontStyle(fontHeightInPoints = 10)
@HeadRowHeight(30)
public class JjgFbgcLjgcLjbpVo extends BaseRowModel {

    @ColumnWidth(23)
    @ExcelProperty(value = "桩号" ,index = 0)
    private String zh;

    @ColumnWidth(23)
    @ExcelProperty(value = "位置" ,index = 1)
    private String wz;

    @ColumnWidth(23)
    @ExcelProperty(value = "设计值" ,index = 2)
    private String sjz;

    @ColumnWidth(23)
    @ExcelProperty(value = "长" ,index = 3)
    private String length;

    @ColumnWidth(23)
    @ExcelProperty(value = "高" ,index = 4)
    private String high;

    @ColumnWidth(23)
    @ExcelProperty(value = "检测日期" ,index = 5)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;
}
