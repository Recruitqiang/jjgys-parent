package glgc.jjgys.model.project;

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableField;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;

import java.io.Serializable;
import java.util.Date;

@Data
@ApiModel(description = "合同段")
@TableName("jjg_htdinfo")
public class JjgHtd {

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    @ApiModelProperty(value = "合同段名称")
    @TableField("name")
    private String name;

    @ApiModelProperty(value = "施工单位")
    @TableField("sgdw")
    private String sgdw;

    @ApiModelProperty(value = "监理单位")
    @TableField("jldw")
    private String jldw;

    @ApiModelProperty(value = "ZY")
    @TableField("zy")
    private String zy;

    @ApiModelProperty(value = "工程部位(起)")
    @TableField("zhq")
    private String zhq;

    @ApiModelProperty(value = "工程部位(止)")
    @TableField("zhz")
    private String zhz;

    @ApiModelProperty(value = "合同段类型")
    @TableField("lx")
    private String lx;

    @ApiModelProperty(value = "项目名称")
    @TableField("proname")
    private String proname;

    @TableField("create_time")
    private Date createTime;
}
