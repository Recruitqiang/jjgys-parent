package glgc.jjgys.model.projectvo.sdgc;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.alibaba.excel.metadata.BaseRowModel;
import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableField;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import com.fasterxml.jackson.annotation.JsonFormat;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;
import lombok.EqualsAndHashCode;

import java.io.Serializable;
import java.util.Date;

/**
 * <p>
 * 
 * </p>
 *
 * @author wq
 * @since 2023-03-27
 */
@Data
@HeadStyle(fillForegroundColor = 13)
@HeadFontStyle(fontHeightInPoints = 10)
@HeadRowHeight(30)
public class JjgFbgcSdgcCqhdVo extends BaseRowModel {

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

    @ColumnWidth(23)
    @ExcelProperty(value = "位置" ,index = 3)
    private String wz;

    /**
     * 设计厚度
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "设计厚度" ,index = 4)
    private String sjhd;

    /**
     * 左拱腰1
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "左拱腰1" ,index = 5)
    private String zgy1;

    /**
     * 左拱腰2
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "左拱腰2" ,index = 6)
    private String zgy2;

    /**
     * 左拱腰3
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "左拱腰3" ,index = 7)
    private String zgy3;

    /**
     * 右拱腰1
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "右拱腰1" ,index = 8)
    private String ygy1;

    /**
     * 右拱腰2
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "右拱腰2" ,index = 9)
    private String ygy2;

    /**
     * 右拱腰3
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "右拱腰3" ,index = 10)
    private String ygy3;

    /**
     * 拱顶
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "拱顶" ,index = 11)
    private String gd;



}
