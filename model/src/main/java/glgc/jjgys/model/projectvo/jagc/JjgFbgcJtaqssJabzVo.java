package glgc.jjgys.model.projectvo.jagc;

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
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_jtaqss_jabz")
public class JjgFbgcJtaqssJabzVo extends BaseRowModel {


    /**
     * 检测时间
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "检测日期" ,index = 0)
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    private Date jcsj;

    /**
     * 位置
     */
    @ColumnWidth(23)
    @ExcelProperty(value = "位置" ,index = 1)
    private String wz;

    @ColumnWidth(23)
    @ExcelProperty(value = "立柱类型" ,index = 2)
    private String lzlx;

    @ColumnWidth(23)
    @ExcelProperty(value = "竖直度允许偏差(mm/m)" ,index = 3)
    private String szdyxps;

    @ColumnWidth(23)
    @ExcelProperty(value = "方向1实测值1(mm/m)" ,index = 4)
    private String fx1scz1;

    @ColumnWidth(23)
    @ExcelProperty(value = "方向1实测值2(mm/m)" ,index = 5)
    private String fx1scz2;

    @ColumnWidth(23)
    @ExcelProperty(value = "方向2实测值1(mm/m)" ,index = 6)
    private String fx2scz1;

    @ColumnWidth(23)
    @ExcelProperty(value = "方向2实测值2(mm/m)" ,index = 7)
    private String fx2scz2;

    @ColumnWidth(23)
    @ExcelProperty(value = "净空规定值(mm)" ,index = 8)
    private String jkgdz;

    @ColumnWidth(23)
    @ExcelProperty(value = "净空实测值(mm)" ,index = 9)
    private String jkscz;

    @ColumnWidth(23)
    @ExcelProperty(value = "厚度允许偏差(mm)" ,index = 10)
    private String hdyxps;

    @ColumnWidth(23)
    @ExcelProperty(value = "厚度测值1(mm)" ,index = 11)
    private String hdcz1;

    @ColumnWidth(23)
    @ExcelProperty(value = "厚度测值2(mm)" ,index = 12)
    private String hdcz2;

    @ColumnWidth(23)
    @ExcelProperty(value = "白色V类允许偏差(cd*lx-1*m-2)" ,index = 13)
    private String bsvlyxps;

    @ColumnWidth(23)
    @ExcelProperty(value = "白色V类实测值1" ,index = 14)
    private String bsvlscz1;

    @ColumnWidth(23)
    @ExcelProperty(value = "白色V类实测值2" ,index = 15)
    private String bsvlscz2;

    @ColumnWidth(23)
    @ExcelProperty(value = "白色IV类允许偏差" ,index = 16)
    private String bswlyxps;

    @ColumnWidth(23)
    @ExcelProperty(value = "白色IV类实测值1" ,index = 17)
    private String bswlscz1;

    @ColumnWidth(23)
    @ExcelProperty(value = "白色IV类实测值2" ,index = 18)
    private String bswlscz2;

    @ColumnWidth(23)
    @ExcelProperty(value = "绿色V类允许偏差(cd*lx-1*m-2)" ,index = 19)
    private String lsvlyxps;

    @ColumnWidth(23)
    @ExcelProperty(value = "绿色V类实测值1" ,index = 20)
    private String lsvlscz1;

    @ColumnWidth(23)
    @ExcelProperty(value = "绿色V类实测值2" ,index = 21)
    private String lsvlscz2;

    @ColumnWidth(23)
    @ExcelProperty(value = "绿色IV类允许偏差" ,index = 22)
    private String lswlyxps;

    @ColumnWidth(23)
    @ExcelProperty(value = "绿色IV类实测值1" ,index = 23)
    private String lswlscz1;

    @ColumnWidth(23)
    @ExcelProperty(value = "绿色IV类实测值2" ,index = 24)
    private String lswlscz2;

    @ColumnWidth(23)
    @ExcelProperty(value = "黄色V类允许偏差(cd*lx-1*m-2)" ,index = 25)
    private String hsvlyxps;

    @ColumnWidth(23)
    @ExcelProperty(value = "黄色V类实测值1" ,index = 26)
    private String hsvlscz1;

    @ColumnWidth(23)
    @ExcelProperty(value = "黄色V类实测值2" ,index = 27)
    private String hsvlscz2;

    @ColumnWidth(23)
    @ExcelProperty(value = "黄色IV类允许偏差" ,index = 28)
    private String hswlyxps;

    @ColumnWidth(23)
    @ExcelProperty(value = "黄色IV类实测值1" ,index = 29)
    private String hswlscz1;

    @ColumnWidth(23)
    @ExcelProperty(value = "黄色IV类实测值2" ,index = 30)
    private String hswlscz2;

    @ColumnWidth(23)
    @ExcelProperty(value = "蓝色V类允许偏差(cd*lx-1*m-2)" ,index = 31)
    private String lasvlyxps;

    @ColumnWidth(23)
    @ExcelProperty(value = "蓝色V类实测值1" ,index = 32)
    private String lasvlscz1;

    @ColumnWidth(23)
    @ExcelProperty(value = "蓝色V类实测值2" ,index = 33)
    private String lasvlscz2;

    @ColumnWidth(23)
    @ExcelProperty(value = "蓝色IV类允许偏差" ,index = 34)
    private String laswlyxps;

    @ColumnWidth(23)
    @ExcelProperty(value = "蓝色IV类实测值1" ,index = 35)
    private String laswlscz1;

    @ColumnWidth(23)
    @ExcelProperty(value = "蓝色IV类实测值2" ,index = 36)
    private String laswlscz2;

    @ColumnWidth(23)
    @ExcelProperty(value = "红色V类允许偏差(cd*lx-1*m-2)" ,index = 37)
    private String rsvlyxps;

    @ColumnWidth(23)
    @ExcelProperty(value = "红色V类实测值1" ,index = 38)
    private String rsvlscz1;

    @ColumnWidth(23)
    @ExcelProperty(value = "红色V类实测值2" ,index = 39)
    private String rsvlscz2;

    @ColumnWidth(23)
    @ExcelProperty(value = "红色IV类允许偏差" ,index = 40)
    private String rswlyxps;

    @ColumnWidth(23)
    @ExcelProperty(value = "红色IV类实测值1" ,index = 41)
    private String rswlscz1;

    @ColumnWidth(23)
    @ExcelProperty(value = "红色IV类实测值2" ,index = 42)
    private String rswlscz2;



}
