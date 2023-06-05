package glgc.jjgys.model.project;

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableName;
import java.time.LocalDate;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableField;
import java.io.Serializable;
import java.util.Date;

import io.swagger.annotations.ApiModelProperty;
import lombok.Data;
import lombok.EqualsAndHashCode;

/**
 * <p>
 * 
 * </p>
 *
 * @author wq
 * @since 2023-04-25
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_lmgc_gslqlmhdzxf")
public class JjgFbgcLmgcGslqlmhdzxf implements Serializable {

    private static final long serialVersionUID = 1L;

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    /**
     * 检测时间
     */
    @TableField("jcsj")
    private Date jcsj;

    /**
     * 路、桥、隧、匝道
     */
    @TableField("lx")
    private String lx;

    /**
     * 桩号
     */
    @TableField("zh")
    private String zh;

    /**
     * 上面层测值1(mm)
     */
    @TableField("smccz1")
    private String smccz1;

    /**
     * 上面层测值2(mm)
     */
    @TableField("smccz2")
    private String smccz2;

    /**
     * 上面层测值3(mm)
     */
    @TableField("smccz3")
    private String smccz3;

    /**
     * 上面层测值4(mm)
     */
    @TableField("smccz4")
    private String smccz4;

    /**
     * 上面层设计值(mm)
     */
    @TableField("smcsjz")
    private String smcsjz;

    /**
     * 总厚度测值1(mm)
     */
    @TableField("zhdcz1")
    private String zhdcz1;

    /**
     * 总厚度测值2(mm)
     */
    @TableField("zhdcz2")
    private String zhdcz2;

    /**
     * 总厚度测值3(mm)
     */
    @TableField("zhdcz3")
    private String zhdcz3;

    /**
     * 总厚度测值4(mm)
     */
    @TableField("zhdcz4")
    private String zhdcz4;

    /**
     * 总厚度设计值(mm)
     */
    @TableField("zhdsjz")
    private String zhdsjz;

    @TableField("proname")
    private String proname;

    @TableField("fbgc")
    private String fbgc;

    @TableField("htd")
    private String htd;

    @TableField("createTime")
    private Date createtime;


}
