package glgc.jjgys.model.project;

import java.time.LocalDate;
import java.io.Serializable;
import java.util.Date;

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableField;
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
@ApiModel(description = "路基沉降")
@TableName("jjg_fbgc_ljgc_ljcj")
public class JjgFbgcLjgcLjcj implements Serializable {

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


    @ApiModelProperty(value = "允许偏差")
    @TableField("yxps")
    private String yxps;

    @ApiModelProperty(value = "检查桩号")
    @TableField("jczh")
    private String jczh;

    @ApiModelProperty(value = "碾压读数1")
    @TableField("nyds1")
    private String nyds1;

    @ApiModelProperty(value = "碾压读数2")
    @TableField("nyds2")
    private String nyds2;

    @ApiModelProperty(value = "合同段")
    @TableField("htd")
    private String htd;

    @ApiModelProperty(value = "项目名称")
    @TableField("proname")
    private String proname;

    @ApiModelProperty(value = "分部工程")
    @TableField("fbgc")
    private String fbgc;

    @ApiModelProperty(value = "备注")
    @TableField("bz")
    private String bz;

    @ApiModelProperty(value = "序号")
    @TableField("xh")
    private String xh;

    @ApiModelProperty(value = "创建时间")
    @TableField("createTime")
    private Date createtime;


}
