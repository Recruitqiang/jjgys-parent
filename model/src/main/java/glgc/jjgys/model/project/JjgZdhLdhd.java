package glgc.jjgys.model.project;

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableName;
import java.time.LocalDate;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableField;
import java.io.Serializable;
import java.util.Date;

import io.swagger.annotations.ApiModelProperty;
import lombok.Data;
import lombok.EqualsAndHashCode;

/**
 * <p>
 * 
 * </p>
 *
 * @author wq
 * @since 2023-07-01
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_zdh_ldhd")
public class JjgZdhLdhd implements Serializable {

    private static final long serialVersionUID = 1L;

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    /**
     * 桩号
     */
    @TableField("zh")
    private Double zh;

    @TableField("ld")
    private String ld;

    /**
     * 车道
     */
    @TableField("cd")
    private String cd;

    /**
     * 类型标识
     */
    @TableField("lxbs")
    private String lxbs;

    /**
     * 匝道标识
     */
    @TableField("zdbs")
    private String zdbs;

    @TableField("proname")
    private String proname;

    @TableField("htd")
    private String htd;

    @TableField("createTime")
    private Date createtime;


}
