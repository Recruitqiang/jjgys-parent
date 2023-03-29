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
@ApiModel(description = "收费站清单")
@TableName("jjg_lqs_sfz")
public class JjgSfz {

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    @ApiModelProperty(value = "匝道收费站名称")
    @TableField("name")
    private String zdsfzname;

    @ApiModelProperty(value = "合同段")
    @TableField("htd")
    private String htd;

    @ApiModelProperty(value = "桩号")
    @TableField("zh")
    private String zh;

    @ApiModelProperty(value = "铺筑类型")
    @TableField("pzlx")
    private String pzlx;

    @ApiModelProperty(value = "所属匝道")
    @TableField("sszd")
    private String sszd;

    @ApiModelProperty(value = "所属互通名称")
    @TableField("sshtmc")
    private String sshtmc;

    @ApiModelProperty(value = "项目名称")
    @TableField("proname")
    private String proname;

    @TableField("create_time")
    private Date createTime;
}
