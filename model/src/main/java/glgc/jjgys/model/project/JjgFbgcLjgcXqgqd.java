package glgc.jjgys.model.project;

import java.time.LocalDate;
import java.time.LocalDateTime;

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableField;
import java.io.Serializable;
import java.util.Date;

import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import com.fasterxml.jackson.annotation.JsonFormat;
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
 * @since 2023-02-13
 */
@Data
@ApiModel(description = "小桥砼强度")
@TableName("jjg_fbgc_ljgc_xqgqd")
public class JjgFbgcLjgcXqgqd implements Serializable {

    private static final long serialVersionUID = 1L;

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    @ApiModelProperty(value = "检测时间")
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    @TableField("jcsj")
    private Date jcsj;

    @ApiModelProperty(value = "桩号")
    @TableField("zh")
    private String zh;

    @ApiModelProperty(value = "部位1")
    @TableField("bw1")
    private String bw1;

    @ApiModelProperty(value = "部位2")
    @TableField("bw2")
    private String bw2;

    @ApiModelProperty(value = "测定值1")
    @TableField("cdz1")
    private String cdz1;

    @ApiModelProperty(value = "值2")
    @TableField("z2")
    private String z2;

    @ApiModelProperty(value = "值3")
    @TableField("z3")
    private String z3;

    @ApiModelProperty(value = "值4")
    @TableField("z4")
    private String z4;

    @ApiModelProperty(value = "值5")
    @TableField("z5")
    private String z5;

    @ApiModelProperty(value = "值6")
    @TableField("z6")
    private String z6;

    @ApiModelProperty(value = "值7")
    @TableField("z7")
    private String z7;

    @ApiModelProperty(value = "值8")
    @TableField("z8")
    private String z8;

    @ApiModelProperty(value = "值9")
    @TableField("z9")
    private String z9;

    @ApiModelProperty(value = "值10")
    @TableField("z10")
    private String z10;

    @ApiModelProperty(value = "值11")
    @TableField("z11")
    private String z11;

    @ApiModelProperty(value = "值12")
    @TableField("z12")
    private String z12;

    @ApiModelProperty(value = "值13")
    @TableField("z13")
    private String z13;

    @ApiModelProperty(value = "值14")
    @TableField("z14")
    private String z14;

    @ApiModelProperty(value = "值15")
    @TableField("z15")
    private String z15;

    @ApiModelProperty(value = "值16")
    @TableField("z16")
    private String z16;

    @ApiModelProperty(value = "回弹角度")
    @TableField("htjd")
    private String htjd;

    @ApiModelProperty(value = "浇筑面")
    @TableField("jzm")
    private String jzm;

    @ApiModelProperty(value = "碳化深度")
    @TableField("thsd")
    private String thsd;

    @ApiModelProperty(value = "设计强度")
    @TableField("sjqd")
    private String sjqd;

    @ApiModelProperty(value = "创建时间")
    @TableField("createTime")
    private Date createtime;

    @ApiModelProperty(value = "合同段名称")
    @TableField("htd")
    private String htd;

    @ApiModelProperty(value = "项目名称")
    @TableField("proname")
    private String proname;

    @ApiModelProperty(value = "分部工程")
    @TableField("fbgc")
    private String fbgc;

    @ApiModelProperty(value = "是否泵送")
    @TableField("sfbs")
    private String sfbs;


}
