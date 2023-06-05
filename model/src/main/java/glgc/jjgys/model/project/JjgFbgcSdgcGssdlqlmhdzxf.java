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
 * @since 2023-05-20
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_sdgc_gssdlqlmhdzxf")
public class JjgFbgcSdgcGssdlqlmhdzxf implements Serializable {

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
    @TableField("lqszd")
    private String lqszd;

    /**
     * 隧道名称
     */
    @TableField("sdmc")
    private String sdmc;

    /**
     * 桩号
     */
    @TableField("zh")
    private String zh;

    /**
     * 部位
     */
    @TableField("bw")
    private String bw;

    /**
     * 上面层测值1(mm)
     */
    @TableField("smcz1")
    private String smcz1;

    /**
     * 上面层测值2(mm)
     */
    @TableField("smcz2")
    private String smcz2;

    /**
     * 上面层测值3(mm)
     */
    @TableField("smcz3")
    private String smcz3;

    /**
     * 上面层测值4(mm)
     */
    @TableField("smcz4")
    private String smcz4;

    /**
     * 上面层设计值(mm)
     */
    @TableField("smcsjz")
    private String smcsjz;

    /**
     * 总厚度测值1(mm)
     */
    @TableField("zhdz1")
    private String zhdz1;

    /**
     * 总厚度测值2(mm)
     */
    @TableField("zhdz2")
    private String zhdz2;

    /**
     * 总厚度测值3(mm)
     */
    @TableField("zhdz3")
    private String zhdz3;

    /**
     * 总厚度测值4(mm)
     */
    @TableField("zhdz4")
    private String zhdz4;

    /**
     * 总厚度设计值(mm)
     */
    @TableField("zhdsjz")
    private String zhdsjz;

    @TableField("proname")
    private String proname;

    @TableField("htd")
    private String htd;

    @TableField("fbgc")
    private String fbgc;

    @TableField("createTime")
    private Date createtime;


}
