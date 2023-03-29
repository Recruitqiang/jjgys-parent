package glgc.jjgys.model.project;

import java.time.LocalDate;

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
 * @since 2023-02-21
 */
@Data
@ApiModel(description = "路基土石方压实度_灰土")
@TableName("jjg_fbgc_ljgc_ljtsfysd_ht")
public class JjgFbgcLjgcLjtsfysdHt implements Serializable {

    private static final long serialVersionUID = 1L;

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    @ApiModelProperty(value = "路基压实度_灰土规定值")
    @TableField("htgdz")
    private String htgdz;

    @ApiModelProperty(value = "试验时间")
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    @TableField("sysj")
    private Date sysj;//4

    @ApiModelProperty(value = "桩号")
    @TableField("zh")
    private String zh;

    @ApiModelProperty(value = "结构层次")
    @TableField("jgcc")
    private String jgcc;//8

    @ApiModelProperty(value = "结构类型")
    @TableField("jglx")
    private String jglx;

    @ApiModelProperty(value = "最大干密度")
    @TableField("zdgmd")
    private String zdgmd;

    @ApiModelProperty(value = "标准砂密度")
    @TableField("bzsmd")
    private String bzsmd;

    @ApiModelProperty(value = "取样桩号及位置")
    @TableField("qyzhjwz")
    private String qyzhjwz;

    @ApiModelProperty(value = "试坑深度(cm)")
    @TableField("sksd")
    private String sksd;

    @ApiModelProperty(value = "锥体及基板和表面间砂质量（g）")
    @TableField("ztjjbhbmjszl")
    private String ztjjbhbmjszl;

    @ApiModelProperty(value = "灌砂前筒+砂质量（g）")
    @TableField("gsqtszl")
    private String gsqtszl;

    @ApiModelProperty(value = "灌砂后筒+砂质量（g）")
    @TableField("gshtszl")
    private String gshtszl;

    @ApiModelProperty(value = "试样质量（g）")
    @TableField("syzl")
    private String syzl;

    @ApiModelProperty(value = "盒号")
    @TableField("hh")
    private String hh;

    @ApiModelProperty(value = "盒质量（g）")
    @TableField("hzl")
    private String hzl;

    @ApiModelProperty(value = "盒+湿试样质量（g）")
    @TableField("hsshzl")
    private String hsshzl;

    @ApiModelProperty(value = "盒+干试样质量（g）")
    @TableField("hgsyzl")
    private String hgsyzl;

    @ApiModelProperty(value = "序号")
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
