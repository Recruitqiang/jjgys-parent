package glgc.jjgys.system.service.impl;

import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.model.project.JjgHtd;
import glgc.jjgys.model.system.SysMenu;
import glgc.jjgys.system.mapper.JjgHtdMapper;
import glgc.jjgys.system.mapper.SysMenuMapper;
import glgc.jjgys.system.service.JjgHtdService;
import glgc.jjgys.system.service.SysMenuService;
import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;
import java.util.Random;
import java.util.UUID;

@Service
public class JjgHtdServiceImpl extends ServiceImpl<JjgHtdMapper, JjgHtd> implements JjgHtdService {

    @Autowired
    private SysMenuService sysMenuService;

    @Autowired
    private SysMenuMapper sysMenuMapper;

    @Autowired
    private JjgHtdMapper jjgHtdMapper;

    @Override
    public boolean addhtd(JjgHtd jjgHtd) {
        //通过项目名称获取到菜单路由id
        QueryWrapper<SysMenu> wrapper=new QueryWrapper<>();
        System.out.println(jjgHtd.getProname());
        wrapper.like("name",jjgHtd.getProname());
        SysMenu sysMenu = sysMenuMapper.selectOne(wrapper);
        Long projectId = sysMenu.getId();
        //创建实测数据菜单     要判断一下当前是否有实测数据菜单  那就是说查询当前高速下，是否有实测数据这条数据
        String scid = jjgHtdMapper.selectscDataId(jjgHtd.getProname());//实测数据菜单的id

        if ("".equals(scid) || scid == null){
            SysMenu scSysMenu = new SysMenu();
            scSysMenu.setParentId(projectId);
            scSysMenu.setName("实测数据");
            scSysMenu.setType(0);
            scSysMenu.setPath("scData");
            scSysMenu.setComponent("sysproject/projectInfo/scData");
            scSysMenu.setSortValue(3);
            sysMenuService.save(scSysMenu);
        }

       String scid2 =  jjgHtdMapper.selectscDataId(jjgHtd.getProname());
        /*String scid = jjgHtdMapper.selectscDataId(jjgHtd.getProname());
        long l = Long.parseLong(scid);*/
        //创建合同段菜单   LJ-1
        SysMenu htdSysMenu = new SysMenu();
        long htdID=0;
        try {
            htdID = Long.parseLong(uniqId());
        }catch (Exception e){
            e.printStackTrace();

        }
        htdSysMenu.setId(htdID);
        //父id是实测数据菜单的id
        htdSysMenu.setParentId(Long.parseLong(scid2));
        htdSysMenu.setName(jjgHtd.getName());
        htdSysMenu.setType(0);
        htdSysMenu.setPath("htdData");
        htdSysMenu.setComponent("ParentView");
        htdSysMenu.setSortValue(1);
        sysMenuService.save(htdSysMenu);

        //根据合同段的类型，生成分部工程的菜单
        List<String> lx = handle(jjgHtd.getLx());
        System.out.println(lx+"1111111");
        for (String s : lx) {
            SysMenu fbgcSysMenu = new SysMenu();
            fbgcSysMenu.setParentId(htdID);
            fbgcSysMenu.setType(1);
            fbgcSysMenu.setSortValue(1);
            if(s.equals("路基工程")){
                fbgcSysMenu.setName("路基工程");
                fbgcSysMenu.setPath("ljlist");
                fbgcSysMenu.setComponent("sysproject/fbgx/ljlist");
            }else if(s.equals("路面工程")){
                fbgcSysMenu.setName("路面工程");
                fbgcSysMenu.setPath("lmlist");
                fbgcSysMenu.setComponent("sysproject/fbgx/lmlist");
            }else if(s.equals("交安工程")){
                fbgcSysMenu.setName("交安工程");
                fbgcSysMenu.setPath("jalist");
                fbgcSysMenu.setComponent("sysproject/fbgx/jalist");
            }else if(s.equals("桥梁工程")){
                fbgcSysMenu.setName("桥梁工程");
                fbgcSysMenu.setPath("qllist");
                fbgcSysMenu.setComponent("sysproject/fbgx/qllist");
            }else if(s.equals("隧道工程")){
                fbgcSysMenu.setName("隧道工程");
                fbgcSysMenu.setPath("sdlist");
                fbgcSysMenu.setComponent("sysproject/fbgx/sdlist");
            }
            sysMenuService.save(fbgcSysMenu);
        }
        //添加合同段信息到表中，并处理了前端传过来的合同段类型
        String join = StringUtils.join(lx, ",");
        jjgHtd.setLx(join);
        jjgHtdMapper.insert(jjgHtd);
        return true;
    }
    /**
     * 处理多合同段类型前端传到后端的对象，转换成字符串存
     * @param a
     * @return
     */
    public static List<String> handle(String a){
        String a1 = StringUtils.strip(a.toString(), "[]");
        String[] split = a1.split(",");
        List<String> list = new ArrayList<>();
        for (String s : split) {
            String str= s.replace("\"", "");
            list.add(str);
            System.out.println(str);
        }
        return list;
    }

    public String uniqId() {
        Random random = new Random();
        String nanoRandom = System.nanoTime() + "" + random.nextInt(99999);
        int hash = Math.abs(UUID.randomUUID().hashCode());
        int needAdd = 19 - String.valueOf(hash).length() + 1;
        /*Long aLong = Long.valueOf(hash+""+nanoRandom.substring(needAdd));
        return aLong;*/
        return hash+""+nanoRandom.substring(needAdd);
    }


}
