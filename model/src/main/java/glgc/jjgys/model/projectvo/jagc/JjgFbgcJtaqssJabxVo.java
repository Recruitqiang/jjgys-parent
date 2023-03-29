package glgc.jjgys.model.projectvo.jagc;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
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
 * @since 2023-03-01
 */
@Data
public class JjgFbgcJtaqssJabxVo extends BaseRowModel {


    /**
     * 检测时间
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "检测日期" ,index = 0)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;

    /**
     * 标线类型
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "标线类型" ,index = 1)
    private String bxlx;

    /**
     * 位置(ZK0+001)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "位置(ZK0+001)" ,index = 2)
    private String wz;

    /**
     * 厚度规定值(mm)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "厚度规定值(mm)" ,index = 3)
    private String hdgdz;

    /**
     * 厚度允许偏差+
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "厚度允许偏差+" ,index = 4)
    private String hdyxpsz;

    /**
     * 厚度允许偏差-
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "厚度允许偏差-" ,index = 5)
    private String hdyxpsf;

    /**
     * 厚度实测值1
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "厚度实测值1" ,index = 6)
    private String hdscz1;

    /**
     * 厚度实测值2
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "厚度实测值2" ,index = 7)
    private String hdscz2;

    /**
     * 厚度实测值3
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "厚度实测值3" ,index = 8)
    private String hdscz3;

    /**
     * 厚度实测值4
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "厚度实测值4" ,index = 9)
    private String hdscz4;

    /**
     * 厚度实测值5
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "厚度实测值5" ,index = 10)
    private String hdscz5;

    /**
     * 白线逆反射系数规定值(mcd*m-2*lx-1)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "白线逆反射系数规定值(mcd*m-2*lx-1)" ,index = 11)
    private String bxnfsxsgdz;

    /**
     * 白线实测值1
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "白线实测值1" ,index = 12)
    private String bxscz1;

    /**
     * 白线实测值2
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "白线实测值2" ,index = 13)
    private String bxscz2;

    /**
     * 白线实测值3
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "白线实测值3" ,index = 14)
    private String bxscz3;

    /**
     * 白线实测值4
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "白线实测值4" ,index = 15)
    private String bxscz4;

    /**
     * 白线实测值5
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "白线实测值5" ,index = 16)
    private String bxscz5;

    /**
     * 黄线逆反射系数规定值
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "黄线逆反射系数规定值" ,index = 17)
    private String hxnfsxsgdz;

    /**
     * 黄线实测值1
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "黄线实测值1" ,index = 18)
    private String hxscz1;

    /**
     * 黄线实测值2
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "黄线实测值2" ,index = 19)
    private String hxscz2;

    /**
     * 黄线实测值3
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "黄线实测值3" ,index = 20)
    private String hxscz3;

    /**
     * 黄线实测值4
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "黄线实测值4" ,index = 21)
    private String hxscz4;

    /**
     * 黄线实测值5
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "黄线实测值5" ,index = 22)
    private String hxscz5;


}
