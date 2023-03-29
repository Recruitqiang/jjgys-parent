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
 * @since 2023-02-27
 */
@Data
public class JjgFbgcLjgcLjwcVo extends BaseRowModel {

    @ColumnWidth(23)
    @ExcelProperty(value = "检测时间" ,index = 0)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;

    @ColumnWidth(23)
    @ExcelProperty(value = "桩号" ,index = 1)
    private String zh;

    @ColumnWidth(23)
    @ExcelProperty(value = "验收弯沉值(0.01mm)" ,index = 2)
    private String yswcz;

    @ColumnWidth(23)
    @ExcelProperty(value = "目标可靠指标" ,index = 3)
    private String mbkkzb;


    @ColumnWidth(23)
    @ExcelProperty(value = "温度影响系数" ,index = 4)
    @TableField("sdyxxs")
    private String sdyxxs;//10

    @ColumnWidth(23)
    @ExcelProperty(value = "季节影响系数" ,index = 5)
    private String jjyxxs;

    @ColumnWidth(23)
    @ExcelProperty(value = "结构层次" ,index = 6)
    private String jgcc;

    @ColumnWidth(23)
    @ExcelProperty(value = "结构类型" ,index = 7)
    private String jglx;

    @ColumnWidth(23)
    @ExcelProperty(value = "后轴重(T)" ,index = 8)
    private String hzz;

    @ColumnWidth(23)
    @ExcelProperty(value = "轮胎气压(MPa)" ,index = 9)
    private String ltqy;//15

    @ColumnWidth(23)
    @ExcelProperty(value = "抽检桩号" ,index = 10)
    private String cjzh;

    @ColumnWidth(23)
    @ExcelProperty(value = "车道" ,index = 11)
    private String cd;

    @ColumnWidth(23)
    @ExcelProperty(value = "左值(0.01mm)" ,index = 12)
    private String zz;

    @ColumnWidth(23)
    @ExcelProperty(value = "右值(0.01mm)" ,index = 13)
    private String yz;

    @ColumnWidth(23)
    @ExcelProperty(value = "路表温度" ,index = 14)
    private String lbwd;


    @ColumnWidth(23)
    @ExcelProperty(value = "序号" ,index = 15)
    private String xh;//22

    @ColumnWidth(23)
    @ExcelProperty(value = "备注" ,index = 16)
    private String bz;


}
