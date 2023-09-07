package glgc.jjgys.system.service.impl;

import glgc.jjgys.model.project.JjgHtd;
import glgc.jjgys.model.system.SysMenu;
import glgc.jjgys.model.system.SysRoleMenu;
import glgc.jjgys.model.vo.AssginMenuVo;
import glgc.jjgys.model.vo.RouterVo;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.SysMenuMapper;
import glgc.jjgys.system.mapper.SysRoleMenuMapper;
import glgc.jjgys.system.service.SysMenuService;
import glgc.jjgys.system.utils.MenuHelper;
import glgc.jjgys.system.utils.RouterHelper;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;

/**
 * <p>
 * 菜单表 服务实现类
 * </p>
 */
@Service
public class SysMenuServiceImpl extends ServiceImpl<SysMenuMapper, SysMenu> implements SysMenuService {

    @Autowired
    private SysRoleMenuMapper sysRoleMenuMapper;

    @Autowired
    private SysMenuMapper sysMenuMapper;

    //菜单列表（树形）
    @Override
    public List<SysMenu> findNodes() {
        //获取所有菜单
        List<SysMenu> sysMenuList = baseMapper.selectList(null);
        //所有菜单数据转换要求数据格式
        List<SysMenu> resultList = MenuHelper.bulidTree(sysMenuList);
        return resultList;
    }

    //删除菜单
    @Override
    public void removeMenuById(String id) {
        //查询当前删除菜单下面是否子菜单
        //根据id = parentid
        QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("parent_id",id);
        Integer count = baseMapper.selectCount(wrapper);
        if(count > 0) {//有子菜单
            throw new JjgysException(201,"请先删除子菜单");
        }
        //调用删除
        baseMapper.deleteById(id);
    }

    //根据角色分配菜单
    @Override
    public List<SysMenu> findMenuByRoleId(String roleId) {
        //获取所有菜单 status=1
//        QueryWrapper<SysMenu> wrapperMenu = new QueryWrapper<>();
//        wrapperMenu.eq("status",1);
        List<SysMenu> menuList = baseMapper.selectList(null);

        //根据角色id查询 角色分配过的菜单列表
        QueryWrapper<SysRoleMenu> wrapperRoleMenu = new QueryWrapper<>();
        wrapperRoleMenu.eq("role_id",roleId);
        List<SysRoleMenu> roleMenus = sysRoleMenuMapper.selectList(wrapperRoleMenu);

        //从第二步查询列表中，获取角色分配所有菜单id
        List<String> roleMenuIds = new ArrayList<>();
        for (SysRoleMenu sysRoleMenu:roleMenus) {
            String menuId = sysRoleMenu.getMenuId();
            roleMenuIds.add(menuId);
        }

        //数据处理：isSelect 如果菜单选中 true，否则false
        // 拿着分配菜单id 和 所有菜单比对，有相同的，让isSelect值true
        for (SysMenu sysMenu:menuList) {
            if(roleMenuIds.contains(sysMenu.getId())) {
                sysMenu.setSelect(true);
            } else {
                sysMenu.setSelect(false);
            }
        }

        //转换成树形结构为了最终显示 MenuHelper方法实现
        List<SysMenu> sysMenus = MenuHelper.bulidTree(menuList);
        return sysMenus;
    }

    //给角色分配菜单权限
    @Override
    public void doAssign(AssginMenuVo assginMenuVo) {
        //根据角色id删除菜单权限
        QueryWrapper<SysRoleMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("role_id",assginMenuVo.getRoleId());
        sysRoleMenuMapper.delete(wrapper);

        //遍历菜单id列表，一个一个进行添加
        List<String> menuIdList = assginMenuVo.getMenuIdList();
        for (String menuId:menuIdList) {
            SysRoleMenu sysRoleMenu = new SysRoleMenu();
            sysRoleMenu.setMenuId(menuId);
            sysRoleMenu.setRoleId(assginMenuVo.getRoleId());
            sysRoleMenuMapper.insert(sysRoleMenu);
        }
    }

    //根据userid查询菜单权限值
    @Override
    public List<RouterVo> getUserMenuList(String userId) {
        //admin是超级管理员，操作所有内容
        List<SysMenu> sysMenuList = null;
        //判断userid值是1代表超级管理员，查询所有权限数据
        if("1".equals(userId)) {
            QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
            wrapper.orderByAsc("sort_value");
            sysMenuList = baseMapper.selectList(wrapper);
        } else {
            //如果userid不是1，其他类型用户，查询这个用户权限
            sysMenuList = baseMapper.findMenuListUserId(userId);
        }

        //构建是树形结构
        List<SysMenu> sysMenuTreeList = MenuHelper.bulidTree(sysMenuList);

        //转换成前端路由要求格式数据
        List<RouterVo> routerVoList = RouterHelper.buildRouters(sysMenuTreeList);
        return routerVoList;
    }

    //根据userid查询按钮权限值
    @Override
    public List<String> getUserButtonList(String userId) {
        List<SysMenu> sysMenuList = null;
        //判断是否管理员
        if("1".equals(userId)) {
            sysMenuList =
                    baseMapper.selectList(new QueryWrapper<SysMenu>().eq("status",1));
        } else {
            sysMenuList = baseMapper.findMenuListUserId(userId);
        }
        //sysMenuList遍历
        List<String> permissionList = new ArrayList<>();
        for (SysMenu sysMenu:sysMenuList) {
            // type=2
            if(sysMenu.getType()==2) {
                String perms = sysMenu.getPerms();
                permissionList.add(perms);
            }
        }
        return permissionList;
    }

    @Override
    public List<SysMenu> selectname(String proname) {
        QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("name", proname);
        List<SysMenu> htdList = sysMenuMapper.selectList(wrapper);
        return htdList;
    }

    @Override
    public List<SysMenu> selectscname(Long pronameid) {
        QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("name", "实测数据");
        wrapper.eq("parent_id", pronameid);
        List<SysMenu> htdList = sysMenuMapper.selectList(wrapper);
        return htdList;
    }

    @Override
    public boolean delecthtd(Long scid, String htd) {
        QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("name", htd);
        wrapper.eq("parent_id", scid);
        sysMenuMapper.delete(wrapper);
        return true;
    }

    @Override
    public List<SysMenu> selecthtd(Long scid, String htd) {
        QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("parent_id", scid);
        List<SysMenu> htdList = sysMenuMapper.selectList(wrapper);
        return htdList;

    }

    @Override
    public boolean delectfbgc(Long fbgcid) {
        QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("parent_id", fbgcid);
        List<SysMenu> list = sysMenuMapper.selectList(wrapper);
        for (SysMenu sysMenu : list) {
            Long id = sysMenu.getId();
            sysMenuMapper.deleteById(id);
        }
        return true;
    }

    @Override
    public SysMenu selectcdinfo(String proName) {
        QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("name", proName);
        wrapper.eq("parent_id", "1600779312636739586");
        SysMenu sysMenu = sysMenuMapper.selectOne(wrapper);
        return sysMenu;
    }

    @Override
    public SysMenu getscChildrenMenu(Long proid) {
        QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("name", "实测数据");
        wrapper.eq("parent_id", proid);
        SysMenu sysMenu = sysMenuMapper.selectOne(wrapper);
        return sysMenu;

    }

    @Override
    public List<SysMenu> getAllHtd(Long scid) {
        QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("parent_id", scid);
        List<SysMenu> menuList = sysMenuMapper.selectList(wrapper);
        return menuList;
    }

    @Override
    public void removeFbgc(List<SysMenu> htdlist) {
        for (SysMenu sysMenu : htdlist) {
            Long htdid = sysMenu.getId();
            QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
            wrapper.eq("parent_id", htdid);
            List<SysMenu> menuList = sysMenuMapper.selectList(wrapper);//所有分部工程的信息
            removefbgcInfo(menuList);
        }
        //
        for (SysMenu sysMenu : htdlist) {
            sysMenuMapper.deleteById(sysMenu.getId());
        }
    }

    @Override
    public void delectChildrenMenu(Long proid) {
        QueryWrapper<SysMenu> wrapper = new QueryWrapper<>();
        wrapper.eq("parent_id", proid);
        List<SysMenu> menuList = sysMenuMapper.selectList(wrapper);
        for (SysMenu sysMenu : menuList) {
            sysMenuMapper.deleteById(sysMenu.getId());
        }

    }

    private void removefbgcInfo(List<SysMenu> menuList) {
        for (SysMenu sysMenu : menuList) {
            sysMenuMapper.deleteById(sysMenu.getId());
        }
    }
}
