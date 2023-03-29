package glgc.jjgys.model.projectvo.ljgc;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
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
 * @since 2023-03-01
 */
@Data
public class JjgFbgcLjgcLjwcLcfVo extends BaseRowModel {

    @ColumnWidth(23)
    @ExcelProperty(value = "检测时间" ,index = 0)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;

    @ColumnWidth(23)
    @ExcelProperty(value = "桩号" ,index = 1)
    private String zh;

    @ColumnWidth(23)
    @ExcelProperty(value = "设计弯沉值(0.01mm)" ,index = 2)
    private String sjwcz;

    @ColumnWidth(23)
    @ExcelProperty(value = "结构层次" ,index = 3)
    private String jgcc;

    @ColumnWidth(23)
    @ExcelProperty(value = "温度影响系数" ,index = 4)
    private String wdyxxs;

    @ColumnWidth(23)
    @ExcelProperty(value = "季节影响系数" ,index = 5)
    private String jjyxxs;

    @ColumnWidth(23)
    @ExcelProperty(value = "结构类型" ,index = 6)
    private String jglx;

    @ColumnWidth(23)
    @ExcelProperty(value = "目标可靠指标" ,index = 7)
    private String mbkkzb;

    @ColumnWidth(23)
    @ExcelProperty(value = "湿度影响系数" ,index = 8)
    private String sdyxxs;

    @ColumnWidth(23)
    @ExcelProperty(value = "落锤重（T)" ,index = 9)
    private String lcz;

    @ColumnWidth(23)
    @ExcelProperty(value = "仪器名称" ,index = 10)
    private String yqmc;

    @ColumnWidth(23)
    @ExcelProperty(value = "抽检桩号" ,index = 11)
    private String cjzh;

    @ColumnWidth(23)
    @ExcelProperty(value = "车道" ,index = 12)
    private String cd;

    @ColumnWidth(23)
    @ExcelProperty(value = "实测弯沉值(0.01㎜)" ,index = 13)
    private String scwcz;

    @ColumnWidth(23)
    @ExcelProperty(value = "路表温度" ,index = 14)
    private String lbwd;

    @ColumnWidth(23)
    @ExcelProperty(value = "序号" ,index = 15)
    private String xh;

    @ColumnWidth(23)
    @ExcelProperty(value = "备注" ,index = 16)
    private String bz;


}
