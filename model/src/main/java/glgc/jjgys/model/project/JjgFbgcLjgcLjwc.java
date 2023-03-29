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
 * @since 2023-02-27
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_ljgc_ljwc")
public class JjgFbgcLjgcLjwc implements Serializable {

    private static final long serialVersionUID = 1L;

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    @TableField("jcsj")
    private Date jcsj;

    @TableField("zh")
    private String zh;

    @TableField("yswcz")
    private String yswcz;

    @TableField("mbkkzb")
    private String mbkkzb;

    @TableField("sdyxxs")
    private String sdyxxs;

    @TableField("jjyxxs")
    private String jjyxxs;

    @TableField("jgcc")
    private String jgcc;

    @TableField("jglx")
    private String jglx;

    @TableField("hzz")
    private String hzz;

    @TableField("ltqy")
    private String ltqy;

    @TableField("cjzh")
    private String cjzh;

    @TableField("cd")
    private String cd;

    @TableField("zz")
    private String zz;

    @TableField("yz")
    private String yz;

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
