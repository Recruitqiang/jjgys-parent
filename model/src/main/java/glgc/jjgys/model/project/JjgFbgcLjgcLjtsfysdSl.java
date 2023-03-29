package glgc.jjgys.model.project;

import java.time.LocalDate;

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableField;
import java.io.Serializable;
import java.util.Date;

import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;
import lombok.EqualsAndHashCode;

/**
 * <p>
 * 
 * </p>
 *
 * @author wq
 * @since 2023-02-21
 */
@Data
@ApiModel(description = "路基土石方压实度_砂砾")
@TableName("jjg_fbgc_ljgc_ljtsfysd_sl")
public class JjgFbgcLjgcLjtsfysdSl implements Serializable {

    private static final long serialVersionUID = 1L;

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    @ApiModelProperty(value = "路基压实度_砂砾规定值")
    @TableField("slgdz")
    private String slgdz;

    @ApiModelProperty(value = "试验时间")
    @TableField("sysj")
    private Date sysj;

    @ApiModelProperty(value = "桩号")
    @TableField("zh")
    private String zh;

    @ApiModelProperty(value = "a")
    @TableField("a")
    private String a;

    @ApiModelProperty(value = "b")
    @TableField("b")
    private String b;

    @ApiModelProperty(value = "c")
    @TableField("c")
    private String c;

    @ApiModelProperty(value = "取样桩号及位置")
    @TableField("qyzhjwz")
    private String qyzhjwz;

    @ApiModelProperty(value = "试坑深度(cm)")
    @TableField("sksd")
    private String sksd;

    @ApiModelProperty(value = "灌砂前筒+砂质量（g）")
    @TableField("gsqtszl")
    private String gsqtszl;

    @ApiModelProperty(value = "灌砂后筒+砂质量（g）")
    @TableField("gshtszl")
    private String gshtszl;

    @ApiModelProperty(value = "锥体砂重")
    @TableField("ztsz")
    private String ztsz;

    @ApiModelProperty(value = "量砂的密度")
    @TableField("lsdmd")
    private String lsdmd;

    @ApiModelProperty(value = "混合料的湿质量")
    @TableField("hhldszl")
    private String hhldszl;

    @ApiModelProperty(value = "盒号")
    @TableField("hh")
    private String hh;

    @ApiModelProperty(value = "盒+干质量")
    @TableField("hgzl")
    private String hgzl;

    @ApiModelProperty(value = "盒质量")
    @TableField("hzl")
    private String hzl;

    @ApiModelProperty(value = "5-38mm颗粒质量(g)")
    @TableField("klzl")
    private String klzl;

    @ApiModelProperty(value = "xh")
    @TableField("xh")
    private String xh;

    @ApiModelProperty(value = "合同段")
    @TableField("htd")
    private String htd;

    @ApiModelProperty(value = "项目名称")
    @TableField("proname")
    private String proname;

    @ApiModelProperty(value = "分部工程")
    @TableField("fbgc")
    private String fbgc;

    @TableField("createTime")
    private Date createtime;


}
