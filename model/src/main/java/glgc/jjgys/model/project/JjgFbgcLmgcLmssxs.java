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
 * @since 2023-04-19
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_lmgc_lmssxs")
public class JjgFbgcLmgcLmssxs implements Serializable {

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
     * 路线类型
     */
    @TableField("lxlx")
    private String lxlx;

    /**
     * 桩号
     */
    @TableField("zh")
    private String zh;

    /**
     * 初读数(mL)
     */
    @TableField("cds")
    private String cds;

    /**
     * 第一分钟读数(mL)
     */
    @TableField("ofzds")
    private String ofzds;

    /**
     * 第二分钟读数(mL)
     */
    @TableField("tfzds")
    private String tfzds;

    /**
     * 水量(mL)
     */
    @TableField("sl")
    private String sl;

    /**
     * 时间(s)
     */
    @TableField("sj")
    private String sj;

    /**
     * 渗水系数规定值
     */
    @TableField("ssxsgdz")
    private String ssxsgdz;

    @TableField("proname")
    private String proname;

    @TableField("fbgc")
    private String fbgc;

    @TableField("htd")
    private String htd;

    @TableField("createTime")
    private Date createtime;


}
