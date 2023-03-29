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
@ApiModel(description = "路基边坡")
@TableName("jjg_fbgc_ljgc_ljbp")
public class JjgFbgcLjgcLjbp implements Serializable {

    private static final long serialVersionUID = 1L;

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    @ApiModelProperty(value = "桩号")
    @TableField("zh")
    private String zh;

    @ApiModelProperty(value = "位置")
    @TableField("wz")
    private String wz;

    @ApiModelProperty(value = "设计值")
    @TableField("sjz")
    private String sjz;

    @ApiModelProperty(value = "长")
    @TableField("length")
    private String length;

    @ApiModelProperty(value = "高")
    @TableField("high")
    private String high;

    @ApiModelProperty(value = "检测时间")
    @JsonFormat(pattern="yyyy-MM-dd",timezone="GMT+8")
    @TableField("jcsj")
    private Date jcsj;

    @ApiModelProperty(value = "创建时间")
    @TableField("createTime")
    private Date createtime;

    @ApiModelProperty(value = "合同段")
    @TableField("htd")
    private String htd;

    @ApiModelProperty(value = "项目名称")
    @TableField("proname")
    private String proname;

    @ApiModelProperty(value = "分部工程")
    @TableField("fbgc")
    private String fbgc;

}
