package glgc.jjgys.model.projectvo.sdgc;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.alibaba.excel.metadata.BaseRowModel;
import com.baomidou.mybatisplus.annotation.TableName;
import com.fasterxml.jackson.annotation.JsonFormat;
import lombok.Data;
import lombok.EqualsAndHashCode;

import java.util.Date;

/**
 * <p>
 * 
 * </p>
 *
 * @author wq
 * @since 2023-05-04
 */
@Data
@HeadStyle(fillForegroundColor = 13)
@HeadFontStyle(fontHeightInPoints = 10)
@HeadRowHeight(30)
public class JjgFbgcSdgcHntlmqdVo extends BaseRowModel {


    /**
     * 检测时间
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "检测时间" ,index = 0)
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
     * 试样平均直径(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "试样平均直径(mm)" ,index = 3)
    private String sypjzj;

    /**
     * 试样平均厚度(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "试样平均厚度(mm)" ,index = 4)
    private String sypjhd;

    /**
     * 极限荷载(N)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "极限荷载(N)" ,index = 5)
    private String jxhz;

    /**
     * 路面强度规定值
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "路面强度规定值" ,index = 6)
    private String lmqdgdz;



}
