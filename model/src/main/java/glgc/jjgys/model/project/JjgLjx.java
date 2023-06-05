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
@ApiModel(description = "连接线清单")
@TableName("jjg_lqs_ljx")
public class JjgLjx {

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    @ApiModelProperty(value = "连接线名称")
    @TableField("name")
    private String ljxname;

    @ApiModelProperty(value = "合同段")
    @TableField("htd")
    private String htd;

    @ApiModelProperty(value = "路幅")
    @TableField("lf")
    private String lf;

    @ApiModelProperty(value = "桩号起")
    @TableField("zhq")
    private String zhq;

    @ApiModelProperty(value = "桩号止")
    @TableField("zhz")
    private String zhz;

    @ApiModelProperty(value = "标志")
    @TableField("bz")
    private String bz;

    @ApiModelProperty(value = "连接线路面所属合同段")
    @TableField("sshtd")
    private String ljxlmsshtd;

    @ApiModelProperty(value = "项目名称")
    @TableField("proname")
    private String proname;

    @TableField("create_time")
    private Date createTime;
}
