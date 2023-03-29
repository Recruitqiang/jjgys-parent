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
 * @since 2023-03-27
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_sdgc_cqhd")
public class JjgFbgcSdgcCqhd implements Serializable {

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
     * 隧道名称
     */
    @TableField("sdmc")
    private String sdmc;

    /**
     * 桩号
     */
    @TableField("zh")
    private String zh;

    /**
     * 位置
     */
    @TableField("wz")
    private String wz;

    /**
     * 设计厚度
     */
    @TableField("sjhd")
    private String sjhd;

    /**
     * 左拱腰1
     */
    @TableField("zgy1")
    private String zgy1;

    /**
     * 左拱腰2
     */
    @TableField("zgy2")
    private String zgy2;

    /**
     * 左拱腰3
     */
    @TableField("zgy3")
    private String zgy3;

    /**
     * 右拱腰1
     */
    @TableField("ygy1")
    private String ygy1;

    /**
     * 右拱腰2
     */
    @TableField("ygy2")
    private String ygy2;

    /**
     * 右拱腰3
     */
    @TableField("ygy3")
    private String ygy3;

    /**
     * 拱顶
     */
    @TableField("gd")
    private String gd;

    @TableField("proname")
    private String proname;

    @TableField("htd")
    private String htd;

    @TableField("fbgc")
    private String fbgc;

    @TableField("createTime")
    private Date createtime;


}
