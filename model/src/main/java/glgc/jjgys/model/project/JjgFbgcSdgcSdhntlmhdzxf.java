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
@TableName("jjg_fbgc_sdgc_sdhntlmhdzxf")
public class JjgFbgcSdgcSdhntlmhdzxf implements Serializable {

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
     * 测值1(mm)
     */
    @TableField("cz1")
    private String cz1;

    /**
     * 测值2(mm)
     */
    @TableField("cz2")
    private String cz2;

    /**
     * 测值3(mm)
     */
    @TableField("cz3")
    private String cz3;

    /**
     * 测值4(mm)
     */
    @TableField("cz4")
    private String cz4;

    /**
     * 设计值(mm)
     */
    @TableField("sjz")
    private String sjz;

    /**
     * 允许偏差(mm)
     */
    @TableField("yxps")
    private String yxps;

    @TableField("proname")
    private String proname;

    @TableField("fbgc")
    private String fbgc;

    @TableField("htd")
    private String htd;

    @TableField("createTime")
    private Date createtime;


}
