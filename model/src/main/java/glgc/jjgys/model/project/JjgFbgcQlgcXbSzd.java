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
@TableName("jjg_fbgc_qlgc_xb_szd")
public class JjgFbgcQlgcXbSzd implements Serializable {

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
     * 墩台号   5
     */
    @TableField("dth")
    private String dth;

    /**
     * 墩柱高度
     */
    @TableField("dzgd")
    private String dzgd;



    /**
     * 横向实测值
     */
    @TableField("hxscz")
    private String hxscz;

    /**
     * 纵向实测值(mm)
     */
    @TableField("zxscz")
    private String zxscz;

    @TableField("createTime")
    private Date createtime;

    @TableField("htd")
    private String htd;

    @TableField("proname")
    private String proname;

    @TableField("fbgc")
    private String fbgc;


}
