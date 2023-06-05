package glgc.jjgys.model.projectvo.lqs;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.alibaba.excel.metadata.BaseRowModel;
import com.baomidou.mybatisplus.annotation.TableField;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;

@Data
@HeadStyle(fillForegroundColor = 13)
@HeadFontStyle(fontHeightInPoints = 10)
@HeadRowHeight(30)
public class LjxVo extends BaseRowModel {

    @ColumnWidth(23)
    @ExcelProperty(value = "连接线名称" ,index = 0)
    private String ljxname;

    @ColumnWidth(23)
    @ExcelProperty(value = "合同段(填路基标名称)" ,index = 1)
    private String htd;

    @ColumnWidth(23)
    @ExcelProperty(value = "路幅" ,index = 2)
    private String lf;

    @ColumnWidth(23)
    @ExcelProperty(value = "桩号起" ,index = 3)
    private String zhq;

    @ColumnWidth(23)
    @ExcelProperty(value = "桩号止" ,index = 4)
    private String zhz;

    @ColumnWidth(23)
    @ExcelProperty(value = "标志" ,index = 5)
    private String bz;

    @ColumnWidth(23)
    @ExcelProperty(value = "连接线路面所属合同段" ,index = 6)
    private String ljxlmsshtd;
}
