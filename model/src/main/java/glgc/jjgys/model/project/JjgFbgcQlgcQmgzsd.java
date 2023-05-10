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
 * 桥面构造深度
 * </p>
 *
 * @author wq
 * @since 2023-03-20
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_qlgc_qmgzsd")
public class JjgFbgcQlgcQmgzsd implements Serializable {

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
     * 桥梁名称
     */
    @TableField("qlmc")
    private String qlmc;

    /**
     * 结构名称
     */
    @TableField("jgmc")
    private String jgmc;

    /**
     * ABM
     */
    @TableField("abm")
    private String abm;

    /**
     * ZY
     */
    @TableField("zy")
    private String zy;

    /**
     * 桩号
     */
    @TableField("zh")
    private String zh;

    /**
     * 车道
     */
    @TableField("cd")
    private String cd;

    /**
     * 设计最小值
     */
    @TableField("sjzxz")
    private String sjzxz;

    /**
     * 设计最大值
     */
    @TableField("sjzdz")
    private String sjzdz;

    /**
     * 测点1D1(㎜)
     */
    @TableField("cd1d1")
    private String cd1d1;

    /**
     * 测点1D2(㎜)
     */
    @TableField("cd1d2")
    private String cd1d2;

    /**
     * 测点2D1(㎜)
     */
    @TableField("cd2d1")
    private String cd2d1;

    /**
     * 测点2D2(㎜)
     */
    @TableField("cd2d2")
    private String cd2d2;

    /**
     * 测点3D1(㎜)
     */
    @TableField("cd3d1")
    private String cd3d1;

    /**
     * 测点3D2(㎜)
     */
    @TableField("cd3d2")
    private String cd3d2;

    @TableField("proname")
    private String proname;

    @TableField("htd")
    private String htd;

    @TableField("fbgc")
    private String fbgc;

    @TableField("createTime")
    private Date createtime;


}
