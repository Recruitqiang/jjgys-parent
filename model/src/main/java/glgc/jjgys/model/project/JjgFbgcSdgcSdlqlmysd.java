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
 * @since 2023-05-07
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_sdgc_sdlqlmysd")
public class JjgFbgcSdgcSdlqlmysd implements Serializable {

    private static final long serialVersionUID = 1L;

    @TableId(type = IdType.ASSIGN_UUID)
    @ApiModelProperty(value = "id")
    @TableField("id")
    private String id;

    /**
     * 路面类型
     */
    @TableField("lmlx")
    private String lmlx;

    /**
     * 检测时间
     */
    @TableField("jcsj")
    private Date jcsj;

    /**
     * 路，桥，隧
     */
    @TableField("lqs")
    private String lqs;

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
     * 取样位置
     */
    @TableField("qywz")
    private String qywz;

    /**
     * 层位
     */
    @TableField("cw")
    private String cw;

    /**
     * 干燥试件质量(g)
     */
    @TableField("gzsjzl")
    private String gzsjzl;

    /**
     * 试件水中质量(g)
     */
    @TableField("sjszzl")
    private String sjszzl;

    /**
     * 时间表干质量(g)
     */
    @TableField("sjbgzl")
    private String sjbgzl;

    /**
     * 实验室标准密度
     */
    @TableField("sysbzmd")
    private String sysbzmd;

    /**
     * 最大理论密度
     */
    @TableField("zdllmd")
    private String zdllmd;

    /**
     * 实验室标准密度规定值
     */
    @TableField("sysbzmdgdz")
    private String sysbzmdgdz;

    /**
     * 最大理论密度规定值
     */
    @TableField("zdllmdgdz")
    private String zdllmdgdz;

    @TableField("proname")
    private String proname;

    @TableField("htd")
    private String htd;

    @TableField("fbgc")
    private String fbgc;

    @TableField("createTime")
    private Date createtime;


}
