package glgc.jjgys.system.mapper;

import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.model.vo.SysUserQueryVo;
import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;
import org.springframework.stereotype.Repository;

import java.util.List;

/**
 * <p>
 * 用户表 Mapper 接口
 * </p>
 */
@Repository
@Mapper
public interface SysUserMapper extends BaseMapper<SysUser> {

    String verifyPwd(@Param("userName") String userName);

    IPage<SysUser> selectPage(Page<SysUser> pageParam, @Param("vo") SysUserQueryVo sysUserQueryVo);

    List<String> selectUserName();

    boolean updatepwd(String username, String encrypt);
}
