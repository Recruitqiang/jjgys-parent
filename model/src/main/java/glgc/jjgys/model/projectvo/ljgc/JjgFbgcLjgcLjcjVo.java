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
public class JjgFbgcLjgcLjcjVo extends BaseRowModel {

    @ColumnWidth(23)
    @ExcelProperty(value = "检测日期" ,index = 0)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;


    @ColumnWidth(23)
    @ExcelProperty(value = "桩号" ,index = 1)
    private String zh;


    @ColumnWidth(23)
    @ExcelProperty(value = "允许偏差" ,index = 2)
    private String yxps;

    @ColumnWidth(23)
    @ExcelProperty(value = "检查桩号" ,index = 3)
    private String jczh;

    @ColumnWidth(23)
    @ExcelProperty(value = "碾压读数1" ,index = 4)
    private String nyds1;

    @ColumnWidth(23)
    @ExcelProperty(value = "碾压读数2" ,index = 5)
    private String nyds2;

    @ColumnWidth(23)
    @ExcelProperty(value = "备注" ,index = 6)
    private String bz;

    @ColumnWidth(23)
    @ExcelProperty(value = "序号" ,index = 7)
    private String xh;

}
