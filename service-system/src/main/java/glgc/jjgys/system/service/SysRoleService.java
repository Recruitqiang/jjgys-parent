package glgc.jjgys.system.service;

import glgc.jjgys.model.system.SysRole;
import glgc.jjgys.model.vo.AssginRoleVo;
import glgc.jjgys.model.vo.SysRoleQueryVo;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import com.baomidou.mybatisplus.extension.service.IService;

import java.util.Map;

public interface SysRoleService extends IService<SysRole> {
    //条件分页查询
    IPage<SysRole> selectPage(Page<SysRole> pageParam, SysRoleQueryVo sysRoleQueryVo);

    //获取用户的角色数据
    Map<String, Object> getRolesByUserId(String userId);

    //用户分配角色
    void doAssign(AssginRoleVo assginRoleVo);
}
