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
@ApiModel(description = "连接线隧道")
@TableName("jjg_lqs_ljxsd")
public class JjgLjxSd {

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    @ApiModelProperty(value = "连接线隧道名称")
    @TableField("name")
    private String ljxsdname;

    @ApiModelProperty(value = "合同段")
    @TableField("htd")
    private String htd;

    @ApiModelProperty(value = "桩号")
    @TableField("zh")
    private String zh;

    @ApiModelProperty(value = "铺筑类型")
    @TableField("pzlx")
    private String pzlx;

    @ApiModelProperty(value = "所属连接线名称")
    @TableField("ssljxmc")
    private String ssljxmc;

    @ApiModelProperty(value = "项目名称")
    @TableField("proname")
    private String proname;

    @TableField("create_time")
    private Date createTime;

}
