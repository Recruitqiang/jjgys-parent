package glgc.jjgys.model.projectvo.lqs;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.alibaba.excel.metadata.BaseRowModel;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;

@Data
@HeadStyle(fillForegroundColor = 13)
@HeadFontStyle(fontHeightInPoints = 10)
@HeadRowHeight(30)
public class SdVo extends BaseRowModel {

    @ColumnWidth(23)
    @ExcelProperty(value = "隧道名称" ,index = 0)
    private String sdname;

    @ColumnWidth(23)
    @ExcelProperty(value = "合同段" ,index = 1)
    private String htd;

    @ColumnWidth(23)
    @ExcelProperty(value = "桩号" ,index = 2)
    private String zh;

    @ColumnWidth(23)
    @ExcelProperty(value = "隧道全长" ,index = 3)
    private String sdqc;

    @ColumnWidth(23)
    @ExcelProperty(value = "铺筑类型" ,index = 4)
    private String pzlx;

    @ColumnWidth(23)
    @ExcelProperty(value = "A/B/M(匝道桥填匝道标志)" ,index = 5)
    private String zdbz;

    @ColumnWidth(23)
    @ExcelProperty(value = "位置（填“主线”）" ,index = 6)
    private String wz;


}
