package glgc.jjgys.system.service.impl;

import glgc.jjgys.common.utils.MD5;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.model.vo.RouterVo;
import glgc.jjgys.model.vo.SysUserQueryVo;
import glgc.jjgys.system.mapper.SysUserMapper;
import glgc.jjgys.system.service.SysMenuService;
import glgc.jjgys.system.service.SysUserService;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * <p>
 * 用户表 服务实现类
 * </p>
 */
@Service
public class SysUserServiceImpl extends ServiceImpl<SysUserMapper, SysUser> implements SysUserService {

    @Autowired
    private SysMenuService sysMenuService;
    @Autowired
    private SysUserMapper sysUserMapper;

    //用户列表
    @Override
    public IPage<SysUser> selectPage(Page<SysUser> pageParam, SysUserQueryVo sysUserQueryVo) {
        return baseMapper.selectPage(pageParam,sysUserQueryVo);
    }


    //username查询
    @Override
    public SysUser getUserInfoByUserName(String username) {
        QueryWrapper<SysUser> wrapper = new QueryWrapper<>();
        wrapper.eq("username",username);
        return baseMapper.selectOne(wrapper);
    }

    //根据用户名称获取用户信息（基本信息 和 菜单权限 和 按钮权限数据）
    @Override
    public Map<String, Object> getUserInfo(String username) {
        //根据username查询用户基本信息
        SysUser sysUser = this.getUserInfoByUserName(username);
        //根据userid查询菜单权限值
        List<RouterVo> routerVolist = sysMenuService.getUserMenuList(sysUser.getId().toString());
        //根据userid查询按钮权限值
        List<String> permsList = sysMenuService.getUserButtonList(sysUser.getId().toString());

        Map<String,Object> result = new HashMap<>();
        result.put("name",username);
        result.put("avatar","https://wpimg.wallstcn.com/f778738c-e4f8-4870-b634-56703b4acafe.gif");
        result.put("roles","[\"admin\"]");
        //菜单权限数据
        result.put("routers",routerVolist);
        //按钮权限数据
        result.put("buttons",permsList);
        return result;
    }

    @Override
    public String selectpwd(String username) {
        String userpassword = sysUserMapper.verifyPwd(username);
        return userpassword;
    }

    @Override
    public boolean updatepwd(String username, String newpass) {
        String encrypt = MD5.encrypt(newpass);
        boolean res = sysUserMapper.updatepwd(username,encrypt);
        if(res){
            return true;
        }
        return false;
    }

    @Override
    public boolean checkSysUserPwd(String oldpw, String username) {
        //将用户输入的旧密码加密和数据库中的密码做对比
        String encrypt = MD5.encrypt(oldpw);
        String s = selectpwd(username);
        if (encrypt.equals(s)){
            return true;
        }
        return false;
    }
}
