package glgc.jjgys.model.projectvo.lqs;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.alibaba.excel.metadata.BaseRowModel;
import lombok.Data;

import javax.validation.constraints.Pattern;
import java.io.Serializable;


@Data
@HeadStyle(fillForegroundColor = 13)
@HeadFontStyle(fontHeightInPoints = 10)
@HeadRowHeight(30)
public class QlVo extends BaseRowModel {

    @ColumnWidth(23)
    @ExcelProperty(value = "桥梁名称" ,index = 0)
    private String qlname;

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
    @ExcelProperty(value = "桩号止" ,index = 4)
    private String zhz;

    @ColumnWidth(23)
    @ExcelProperty(value = "桥梁全长" ,index = 5)
    private String qlqc;

    @ColumnWidth(23)
    @ExcelProperty(value = "单孔跨径" ,index = 6)
    private String dkkj;

    @ColumnWidth(23)
    @ExcelProperty(value = "铺筑类型" ,index = 7)
    private String pzlx;

    @ColumnWidth(23)
    @ExcelProperty(value = "A/B/M(匝道桥填匝道标志)" ,index = 8)
    private String bz;

    @ColumnWidth(23)
    @ExcelProperty(value = "位置（主线填合同段名称，匝道填互通名称）" ,index = 9)
    private String wz;

}
