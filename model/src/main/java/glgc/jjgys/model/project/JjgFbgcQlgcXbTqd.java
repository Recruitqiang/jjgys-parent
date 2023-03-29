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
 * @since 2023-03-22
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_qlgc_xb_tqd")
public class JjgFbgcQlgcXbTqd implements Serializable {

    private static final long serialVersionUID = 1L;

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    @TableField("jcsj")
    private Date jcsj;

    @TableField("qlmc")
    private String qlmc;

    @TableField("bw1")
    private String bw1;

    @TableField("bw2")
    private String bw2;

    @TableField("cdz1")
    private String cdz1;

    @TableField("z2")
    private String z2;

    @TableField("z3")
    private String z3;

    @TableField("z4")
    private String z4;

    @TableField("z5")
    private String z5;

    @TableField("z6")
    private String z6;

    @TableField("z7")
    private String z7;

    @TableField("z8")
    private String z8;

    @TableField("z9")
    private String z9;

    @TableField("z10")
    private String z10;

    @TableField("z11")
    private String z11;

    @TableField("z12")
    private String z12;

    @TableField("z13")
    private String z13;

    @TableField("z14")
    private String z14;

    @TableField("z15")
    private String z15;

    @TableField("z16")
    private String z16;

    @TableField("htjd")
    private String htjd;

    @TableField("jzm")
    private String jzm;

    @TableField("thsd")
    private String thsd;

    @TableField("sfbs")
    private String sfbs;

    @TableField("sjqd")
    private String sjqd;

    @TableField("createTime")
    private Date createtime;

    @TableField("htd")
    private String htd;

    @TableField("proname")
    private String proname;

    @TableField("fbgc")
    private String fbgc;


}
