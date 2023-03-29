package glgc.jjgys.system.service.impl;

import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.model.project.JjgFbgcLjgcHdgqd;
import glgc.jjgys.model.system.Project;
import glgc.jjgys.model.system.SysMenu;
import glgc.jjgys.system.mapper.ProjectMapper;
import glgc.jjgys.system.service.ProjectService;
import glgc.jjgys.system.service.SysMenuService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;
import java.util.Random;
import java.util.UUID;

@Service
public class ProjectServiceImpl extends ServiceImpl<ProjectMapper, Project> implements ProjectService {

    @Autowired
    private ProjectMapper projectMapper;
    @Autowired
    private SysMenuService sysMenuService;

    @Override
    public void addOtherInfo(String proName) {
        SysMenu sysMenu = new SysMenu();
        //long uniqId = uniqId();
        //获取当前菜单的所属上级，先写死吧
        sysMenu.setParentId(Long.parseLong("1600779312636739586"));//交工管理的id
        sysMenu.setType(0);
        sysMenu.setPath("gaosuName");
        sysMenu.setComponent("ParentView");
        sysMenu.setName(proName);
        sysMenu.setIcon("el-icon-tickets");
        sysMenu.setSortValue(2);
        sysMenuService.save(sysMenu);

        QueryWrapper<SysMenu> wrapper=new QueryWrapper<>();
        wrapper.like("name",proName);
        List<SysMenu> list = sysMenuService.list(wrapper);
        Long gsid = list.get(0).getId();

        //设置下级菜单
        SysMenu htdSysMenu = new SysMenu();
        htdSysMenu.setParentId(gsid);
        htdSysMenu.setType(1);
        htdSysMenu.setPath("htdInfo");
        htdSysMenu.setComponent("sysproject/projectInfo/htdInfo");
        htdSysMenu.setName("合同段信息");
        htdSysMenu.setSortValue(1);

        SysMenu lqsSysMenu = new SysMenu();
        lqsSysMenu.setParentId(gsid);
        lqsSysMenu.setType(1);
        lqsSysMenu.setPath("lqsInfo");
        lqsSysMenu.setComponent("sysproject/projectInfo/lqsInfo");
        lqsSysMenu.setName("路桥隧信息");
        lqsSysMenu.setSortValue(2);

       /* SysMenu scSysMenu = new SysMenu();
        scSysMenu.setParentId(uniqId);
        scSysMenu.setType(1);
        scSysMenu.setPath("scData");
        scSysMenu.setComponent("sysproject/projectInfo/scData");
        scSysMenu.setName("实测数据");
        scSysMenu.setSortValue(3);*/


        sysMenuService.save(htdSysMenu);
        sysMenuService.save(lqsSysMenu);
        //sysMenuService.save(scSysMenu);
    }

    @Override
    public Integer getlevel(String proname) {
        QueryWrapper<Project> wrapper=new QueryWrapper<>();
        wrapper.like("proname",proname);
        Project project = projectMapper.selectOne(wrapper);
        String grade = project.getGrade();
        Integer integer = Integer.valueOf(grade);
        return integer;
    }

   /* private long uniqId() {
        Random random = new Random();
        String nanoRandom = System.nanoTime() + "" + random.nextInt(99999);
        int hash = Math.abs(UUID.randomUUID().hashCode());
        int needAdd = 19 - String.valueOf(hash).length() + 1;
        *//*Long aLong = Long.valueOf(hash+""+nanoRandom.substring(needAdd));
        return aLong;*//*
        return Long.parseLong(hash+""+nanoRandom.substring(needAdd));
        //return Long.valueOf(hash+""+nanoRandom.substring(needAdd));
    }*/

}
