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
 * @since 2023-05-07
 */
@Data
@HeadStyle(fillForegroundColor = 13)
@HeadFontStyle(fontHeightInPoints = 10)
@HeadRowHeight(30)
public class JjgFbgcSdgcSdlqlmysdVo extends BaseRowModel {

    /**
     * 路面类型
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "路面类型" ,index = 0)
    private String lmlx;

    /**
     * 检测时间
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "检测时间" ,index = 1)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;

    /**
     * 路，桥，隧
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "路，桥，隧" ,index = 2)
    private Date lqs;

    /**
     * 隧道名称
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "隧道名称" ,index = 3)
    private String sdmc;

    /**
     * 桩号
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "桩号" ,index = 4)
    private String zh;

    /**
     * 取样位置
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "取样位置" ,index = 5)
    private String qywz;

    /**
     * 层位
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "层位" ,index = 6)
    private String cw;

    /**
     * 干燥试件质量(g)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "干燥试件质量(g)" ,index = 7)
    private String gzsjzl;

    /**
     * 试件水中质量(g)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "试件水中质量(g)" ,index = 8)
    private String sjszzl;

    /**
     * 时间表干质量(g)
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "时间表干质量(g)" ,index = 9)
    private String sjbgzl;

    /**
     * 实验室标准密度
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "实验室标准密度" ,index = 10)
    private String sysbzmd;

    /**
     * 最大理论密度
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "最大理论密度" ,index = 11)
    private String zdllmd;

    /**
     * 实验室标准密度规定值
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "实验室标准密度规定值" ,index = 12)
    private String sysbzmdgdz;

    /**
     * 最大理论密度规定值
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "最大理论密度规定值" ,index = 13)
    private String zdllmdgdz;


}
