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
 * @since 2023-03-08
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_ljgc_ljwc_lcf")
public class JjgFbgcLjgcLjwcLcf implements Serializable {

    private static final long serialVersionUID = 1L;

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    @TableField("jcsj")
    private Date jcsj;

    @TableField("zh")
    private String zh;

    @TableField("sjwcz")
    private String sjwcz;

    @TableField("jgcc")
    private String jgcc;

    @TableField("wdyxxs")
    private String wdyxxs;

    @TableField("jjyxxs")
    private String jjyxxs;

    @TableField("jglx")
    private String jglx;

    @TableField("mbkkzb")
    private String mbkkzb;

    @TableField("sdyxxs")
    private String sdyxxs;

    @TableField("lcz")
    private String lcz;

    @TableField("yqmc")
    private String yqmc;

    @TableField("cjzh")
    private String cjzh;

    @TableField("cd")
    private String cd;

    @TableField("scwcz")
    private String scwcz;

    @TableField("lbwd")
    private String lbwd;

    @TableField("xh")
    private String xh;

    @TableField("bz")
    private String bz;

    @TableField("proname")
    private String proname;

    @TableField("fbgc")
    private String fbgc;

    @TableField("htd")
    private String htd;

    @TableField("createTime")
    private Date createtime;


}
