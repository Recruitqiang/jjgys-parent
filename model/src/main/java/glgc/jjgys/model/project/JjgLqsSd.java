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
@ApiModel(description = "隧道清单")
@TableName("jjg_lqs_sd")
public class JjgLqsSd {

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    @ApiModelProperty(value = "隧道名称")
    @TableField("name")
    private String sdname;

    @ApiModelProperty(value = "合同段")
    @TableField("htd")
    private String htd;

    @ApiModelProperty(value = "路幅")
    @TableField("lf")
    private String lf;

    @ApiModelProperty(value = "桩号起")
    @TableField("zhq")
    private Double zhq;

    @ApiModelProperty(value = "桩号止")
    @TableField("zhz")
    private Double zhz;

    @ApiModelProperty(value = "隧道全长")
    @TableField("sdqc")
    private String sdqc;

    @ApiModelProperty(value = "铺筑类型")
    @TableField("pzlx")
    private String pzlx;

    @ApiModelProperty(value = "标志")
    @TableField("zdbz")
    private String zdbz;

    @ApiModelProperty(value = "位置")
    @TableField("wz")
    private String wz;

    @ApiModelProperty(value = "项目名称")
    @TableField("proname")
    private String proname;

    @TableField("create_time")
    private Date createTime;



}
