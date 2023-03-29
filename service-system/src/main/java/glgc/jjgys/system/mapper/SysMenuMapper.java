package glgc.jjgys.system.mapper;

import glgc.jjgys.model.system.SysMenu;
import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;
import org.springframework.stereotype.Repository;

import java.util.List;

/**
 * <p>
 * 菜单表 Mapper 接口
 * </p>
 */
@Repository
@Mapper
public interface SysMenuMapper extends BaseMapper<SysMenu> {

    //根据userid查找菜单权限数据
    List<SysMenu> findMenuListUserId(@Param("userId") String userId);
}
