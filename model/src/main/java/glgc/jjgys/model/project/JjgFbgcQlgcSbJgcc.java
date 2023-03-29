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
 * @since 2023-03-20
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_qlgc_sb_jgcc")
public class JjgFbgcQlgcSbJgcc implements Serializable {

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
     * 桥梁名称  4
     */
    @TableField("qlmc")
    private String qlmc;

    /**
     * 梁板号
     */
    @TableField("lbh")
    private String lbh;

    /**
     * 部位
     */
    @TableField("bw")
    private String bw;

    /**
     * 类别
     */
    @TableField("lb")
    private String lb;

    /**
     * 设计值(mm)
     */
    @TableField("sjz")
    private String sjz;

    /**
     * 实测值(mm)
     */
    @TableField("scz")
    private String scz;

    /**
     * 允许偏差+(mm)
     */
    @TableField("yxwcz")
    private String yxwcz;

    /**
     * 允许偏差-(mm)
     */
    @TableField("yxwcf")
    private String yxwcf;

    @TableField("createTime")
    private Date createtime;

    @TableField("htd")
    private String htd;

    @TableField("proname")
    private String proname;

    @TableField("fbgc")
    private String fbgc;


}
