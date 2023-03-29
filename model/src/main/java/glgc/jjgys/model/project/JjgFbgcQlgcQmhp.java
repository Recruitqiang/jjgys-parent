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
@TableName("jjg_fbgc_qlgc_qmhp")
public class JjgFbgcQlgcQmhp implements Serializable {

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
     * 路面类型
     */
    @TableField("lmlx")
    private String lmlx;

    /**
     * 桩号
     */
    @TableField("zh")
    private String zh;

    /**
     * 位置
     */
    @TableField("wz")
    private String wz;

    /**
     * 前视读数(mm)
     */
    @TableField("qsds")
    private String qsds;

    /**
     * 后视读数(mm)
     */
    @TableField("hsds")
    private String hsds;

    /**
     * 长(mm)
     */
    @TableField("length")
    private String length;

    /**
     * 设计值(%)
     */
    @TableField("sjz")
    private String sjz;

    /**
     * 允许偏差(%)
     */
    @TableField("yxps")
    private String yxps;

    @TableField("proname")
    private String proname;

    @TableField("htd")
    private String htd;

    @TableField("fbgc")
    private String fbgc;

    @TableField("createTime")
    private Date createtime;


}
