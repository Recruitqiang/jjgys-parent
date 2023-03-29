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
 * @since 2023-03-01
 */
@Data
@EqualsAndHashCode(callSuper = false)
@TableName("jjg_fbgc_jtaqss_jabz")
public class JjgFbgcJtaqssJabz implements Serializable {

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
     * 位置
     */
    @TableField("wz")
    private String wz;

    @TableField("lzlx")
    private String lzlx;

    @TableField("szdyxps")
    private String szdyxps;

    @TableField("fx1scz1")
    private String fx1scz1;

    @TableField("fx1scz2")
    private String fx1scz2;

    @TableField("fx2scz1")
    private String fx2scz1;

    @TableField("fx2scz2")
    private String fx2scz2;

    @TableField("jkgdz")
    private String jkgdz;

    @TableField("jkscz")
    private String jkscz;

    @TableField("hdyxps")
    private String hdyxps;

    @TableField("hdcz1")
    private String hdcz1;

    @TableField("hdcz2")
    private String hdcz2;

    @TableField("bsvlyxps")
    private String bsvlyxps;

    @TableField("bsvlscz1")
    private String bsvlscz1;

    @TableField("bsvlscz2")
    private String bsvlscz2;

    @TableField("bswlyxps")
    private String bswlyxps;

    @TableField("bswlscz1")
    private String bswlscz1;

    @TableField("bswlscz2")
    private String bswlscz2;

    @TableField("lsvlyxps")
    private String lsvlyxps;

    @TableField("lsvlscz1")
    private String lsvlscz1;

    @TableField("lsvlscz2")
    private String lsvlscz2;

    @TableField("lswlyxps")
    private String lswlyxps;

    @TableField("lswlscz1")
    private String lswlscz1;

    @TableField("lswlscz2")
    private String lswlscz2;

    @TableField("hsvlyxps")
    private String hsvlyxps;

    @TableField("hsvlscz1")
    private String hsvlscz1;

    @TableField("hsvlscz2")
    private String hsvlscz2;

    @TableField("hswlyxps")
    private String hswlyxps;

    @TableField("hswlscz1")
    private String hswlscz1;

    @TableField("hswlscz2")
    private String hswlscz2;

    @TableField("lasvlyxps")
    private String lasvlyxps;

    @TableField("lasvlscz1")
    private String lasvlscz1;

    @TableField("lasvlscz2")
    private String lasvlscz2;

    @TableField("laswlyxps")
    private String laswlyxps;

    @TableField("laswlscz1")
    private String laswlscz1;

    @TableField("laswlscz2")
    private String laswlscz2;

    @TableField("rsvlyxps")
    private String rsvlyxps;

    @TableField("rsvlscz1")
    private String rsvlscz1;

    @TableField("rsvlscz2")
    private String rsvlscz2;

    @TableField("rswlyxps")
    private String rswlyxps;

    @TableField("rswlscz1")
    private String rswlscz1;

    @TableField("rswlscz2")
    private String rswlscz2;

    @TableField("proname")
    private String proname;

    @TableField("htd")
    private String htd;

    @TableField("fbgc")
    private String fbgc;

    @TableField("createTime")
    private Date createtime;


}
