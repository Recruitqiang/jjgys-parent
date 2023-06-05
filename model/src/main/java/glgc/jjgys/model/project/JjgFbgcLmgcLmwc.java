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
 * @since 2023-04-16
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_lmgc_lmwc")
public class JjgFbgcLmgcLmwc implements Serializable {

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
     * 平衡湿度路基顶面回弹模量E0
     */
    @TableField("phsdljdmhtml")
    private String phsdljdmhtml;

    /**
     * 后轴重(T)
     */
    @TableField("hzz")
    private String hzz;

    /**
     * 轮胎气压(MPa)
     */
    @TableField("ltqy")
    private String ltqy;

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
     * 左值(0.01mm)
     */
    @TableField("zz")
    private String zz;

    /**
     * 右值(0.01mm)
     */
    @TableField("yz")
    private String yz;

    /**
     * 路表温度
     */
    @TableField("lbwd")
    private String lbwd;

    /**
     * 备注
     */
    @TableField("bz")
    private String bz;

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

    @TableField("xh")
    private String xh;

    @TableField("proname")
    private String proname;

    @TableField("fbgc")
    private String fbgc;

    @TableField("htd")
    private String htd;

    @TableField("createTime")
    private Date createtime;


}
