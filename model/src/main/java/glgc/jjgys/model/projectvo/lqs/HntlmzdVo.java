package glgc.jjgys.model.projectvo.lqs;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.alibaba.excel.metadata.BaseRowModel;
import com.baomidou.mybatisplus.annotation.TableField;
import lombok.Data;

@Data
@HeadStyle(fillForegroundColor = 13)
@HeadFontStyle(fontHeightInPoints = 10)
@HeadRowHeight(30)
public class HntlmzdVo extends BaseRowModel {

    @ColumnWidth(23)
    @ExcelProperty(value = "混凝土路面名称" ,index = 0)
    private String hntlmname;

    @ColumnWidth(23)
    @ExcelProperty(value = "合同段" ,index = 1)
    private String htd;

    @ColumnWidth(23)
    @ExcelProperty(value = "路幅" ,index = 2)
    private String lf;

    @ColumnWidth(23)
    @ExcelProperty(value = "桩号起" ,index = 3)
    private String zhq;

    @ColumnWidth(23)
    @ExcelProperty(value = "桩号起止" ,index = 4)
    private String zhz;

    @ColumnWidth(23)
    @ExcelProperty(value = "路面全长" ,index = 5)
    private String lmqc;

    @ColumnWidth(23)
    @ExcelProperty(value = "铺筑类型" ,index = 6)
    private String pzlx;

    @ColumnWidth(23)
    @ExcelProperty(value = "A/B/M（匝道填字母，主线不填）" ,index = 7)
    private String zdlx;

    @ColumnWidth(23)
    @ExcelProperty(value = "位置（主线填合同段名称，匝道填互通名称）" ,index = 8)
    private String wz;

    @ColumnWidth(23)
    @ExcelProperty(value = "互通路面所属合同段" ,index = 9)
    private String htlmsshtd;
}
