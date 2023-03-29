package glgc.jjgys.model.vo;

import io.swagger.annotations.ApiModelProperty;
import lombok.Data;


@Data
public class ProjectQueryVo {
    @ApiModelProperty(value = "项目名称")
    private String proName;

    @ApiModelProperty(value = "公路等级")
    private String grade;

    @ApiModelProperty(value = "创建时间")
    private String createTime;
}
