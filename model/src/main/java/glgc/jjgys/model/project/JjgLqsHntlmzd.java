package glgc.jjgys.model.project;

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableField;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;

import java.util.Date;

@Data
@ApiModel(description = "混凝土路面及匝道")
@TableName("jjg_lqs_hntlmzd")
public class JjgLqsHntlmzd {

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    @ApiModelProperty(value = "混凝土路面名称")
    @TableField("name")
    private String hntlmname;

    @ApiModelProperty(value = "合同段")
    @TableField("htd")
    private String htd;

    @ApiModelProperty(value = "桩号")
    @TableField("zh")
    private String zh;

    @ApiModelProperty(value = "路面全长")
    @TableField("lmqc")
    private String lmqc;

    @ApiModelProperty(value = "铺筑类型")
    @TableField("pzlx")
    private String pzlx;

    @ApiModelProperty(value = "标志")
    @TableField("bz")
    private String zdlx;

    @ApiModelProperty(value = "位置")
    @TableField("wz")
    private String wz;

    @ApiModelProperty(value = "互通路面所属合同段")
    @TableField("sshtd")
    private String htlmsshtd;

    @ApiModelProperty(value = "项目名称")
    @TableField("proname")
    private String proname;

    @TableField("create_time")
    private Date createTime;



}
