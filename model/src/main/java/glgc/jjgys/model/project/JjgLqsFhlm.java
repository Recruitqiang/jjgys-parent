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
@ApiModel(description = "复合路面清单")
@TableName("jjg_lqs_fhlm")
public class JjgLqsFhlm {

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    @ApiModelProperty(value = "复合路面名称")
    @TableField("name")
    private String fhlmname;

    @ApiModelProperty(value = "合同段")
    @TableField("htd")
    private String htd;

    @ApiModelProperty(value = "桩号")
    @TableField("zh")
    private String zh;

    @ApiModelProperty(value = "路面全长")
    @TableField("lmqc")
    private String lmqc;

    @ApiModelProperty(value = "项目名称")
    @TableField("proname")
    private String proname;

    @TableField("create_time")
    private Date createTime;




}
