package glgc.jjgys.model.system;

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableField;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import glgc.jjgys.model.base.BaseEntity;
import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;

import java.util.Date;


@Data
@ApiModel(description = "Project")
@TableName("jjg_projectinfo")
public class Project{

    private static final long serialVersionUID = 1L;

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    @ApiModelProperty(value = "项目名称")
    @TableField("proname")
    private String proName;

    @ApiModelProperty(value = "公路等级")
    @TableField("grade")
    private String grade;

    @TableField("create_time")
    private Date createTime;
}
