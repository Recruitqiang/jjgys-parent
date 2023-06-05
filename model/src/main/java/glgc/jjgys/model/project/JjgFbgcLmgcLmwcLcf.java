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
 * @since 2023-04-18
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_lmgc_lmwc_lcf")
public class JjgFbgcLmgcLmwcLcf implements Serializable {

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
     * 工程部位
     */
    @TableField("gcbw")
    private String gcbw;

    /**
     * 验收弯沉值(0.01mm)
     */
    @TableField("yswcz")
    private String yswcz;

    /**
     * 季节影响系数
     */
    @TableField("jjyxxs")
    private String jjyxxs;

    /**
     * 结构层次
     */
    @TableField("jgcc")
    private String jgcc;

    /**
     * 结构类型
     */
    @TableField("jglx")
    private String jglx;

    /**
     * 目标可靠指标
     */
    @TableField("mbkkzb")
    private String mbkkzb;

    /**
     * 湿度影响系数
     */
    @TableField("sdyxxs")
    private String sdyxxs;

    /**
     * 材料层厚度Ha
     */
    @TableField("clchd")
    private String clchd;

    /**
     * 路基顶面回弹模量E0
     */
    @TableField("ljdmhtml")
    private String ljdmhtml;

    /**
     * 落锤重(T)
     */
    @TableField("lcz")
    private String lcz;

    /**
     * 桩号
     */
    @TableField("zh")
    private String zh;

    /**
     * 车道
     */
    @TableField("cd")
    private String cd;

    /**
     * 实测弯沉值(0.01㎜)
     */
    @TableField("scwcz")
    private String scwcz;

    /**
     * 路表温度
     */
    @TableField("lbwd")
    private String lbwd;

    /**
     * 路面类型标注
     */
    @TableField("lmbzlx")
    private String lmbzlx;

    /**
     * 沥青层总厚度
     */
    @TableField("lqczhd")
    private String lqczhd;

    /**
     * 沥青表面温度
     */
    @TableField("lqbmwd")
    private String lqbmwd;

    /**
     * 测前5h平均温度
     */
    @TableField("cqpjwd")
    private String cqpjwd;

    /**
     * 序号
     */
    @TableField("xh")
    private String xh;

    /**
     * 备注
     */
    @TableField("bz")
    private String bz;

    @TableField("proname")
    private String proname;

    @TableField("fbgc")
    private String fbgc;

    @TableField("htd")
    private String htd;

    @TableField("createTime")
    private Date createtime;


}
