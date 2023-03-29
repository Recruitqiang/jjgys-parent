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
 * @since 2023-03-22
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_qlgc_xb_bhchd")
public class JjgFbgcQlgcXbBhchd implements Serializable {

    private static final long serialVersionUID = 1L;

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    /**
     * 检测时间
     */
    @TableField("jcsj")
    private Date jcsj;

    /**
     * 桥梁名称
     */
    @TableField("qlmc")
    private String qlmc;

    /**
     * 构件编号及检测部位
     */
    @TableField("gjbhjjcbw")
    private String gjbhjjcbw;

    /**
     * 钢筋直径(mm)
     */
    @TableField("gjzj")
    private String gjzj;

    /**
     * 设计值(mm)
     */
    @TableField("sjz")
    private String sjz;

    /**
     * 实测值1(mm)
     */
    @TableField("scz1")
    private String scz1;

    @TableField("scz2")
    private String scz2;

    @TableField("scz3")
    private String scz3;

    @TableField("scz4")
    private String scz4;

    @TableField("scz5")
    private String scz5;

    @TableField("scz6")
    private String scz6;

    @TableField("xzz")
    private String xzz;

    /**
     * 允许偏差-(mm)
     */
    @TableField("yxwcf")
    private String yxwcf;

    /**
     * 允许偏差+(mm)
     */
    @TableField("yxwcz")
    private String yxwcz;

    @TableField("createTime")
    private Date createtime;

    @TableField("htd")
    private String htd;

    @TableField("proname")
    private String proname;

    @TableField("fbgc")
    private String fbgc;


}
