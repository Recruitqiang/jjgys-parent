package glgc.jjgys.system.service.impl;

import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.model.project.*;
import glgc.jjgys.model.system.Project;
import glgc.jjgys.model.system.SysMenu;
import glgc.jjgys.system.mapper.*;
import glgc.jjgys.system.service.*;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;
import java.util.Random;
import java.util.UUID;

@Slf4j
@Service
public class ProjectServiceImpl extends ServiceImpl<ProjectMapper, Project> implements ProjectService {

    @Autowired
    private ProjectMapper projectMapper;
    @Autowired
    private SysMenuService sysMenuService;

    @Autowired
    private JjgHtdService jjgHtdService;

    @Override
    public void addOtherInfo(String proName) {
        SysMenu sysMenu = new SysMenu();
        //long uniqId = uniqId();
        //获取当前菜单的所属上级，先写死吧
        sysMenu.setParentId(Long.parseLong("1600779312636739586"));//交工管理的id
        sysMenu.setType(0);
        sysMenu.setPath("gaosuName"+ UUID.randomUUID());
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


        SysMenu xmjdSysMenu = new SysMenu();
        xmjdSysMenu.setParentId(gsid);
        xmjdSysMenu.setType(1);
        xmjdSysMenu.setPath("xmjd");
        xmjdSysMenu.setComponent("sysproject/projectInfo/xmjd");
        xmjdSysMenu.setName("项目进度");
        xmjdSysMenu.setSortValue(4);

        SysMenu dpSysMenu = new SysMenu();
        dpSysMenu.setParentId(gsid);
        dpSysMenu.setType(1);
        dpSysMenu.setPath("dp");
        dpSysMenu.setComponent("sysproject/projectInfo/dp");
        dpSysMenu.setName("大屏");
        dpSysMenu.setSortValue(5);


        sysMenuService.save(htdSysMenu);
        sysMenuService.save(lqsSysMenu);
        sysMenuService.save(xmjdSysMenu);
        sysMenuService.save(dpSysMenu);

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

    @Autowired
    private JjgLqsQlService jjgLqsQlService;
    @Autowired
    private JjgLqsSdService jjgLqsSdService;
    @Autowired
    private JjgLqsFhlmService jjgLqsFhlmService;
    @Autowired
    private JjgLqsHntlmzdService jjgLqsHntlmzdService;
    @Autowired
    private JjgLqsSfzService jjgLqsSfzService;
    @Autowired
    private JjgFbgcLjgcHdgqdService jjgFbgcLjgcHdgqdService;

    @Autowired
    private JjgFbgcLjgcHdjgccService jjgFbgcLjgcHdjgccService;

    @Autowired
    private JjgFbgcLjgcLjbpService jjgFbgcLjgcLjbpService;

    @Autowired
    private JjgFbgcLjgcLjcjService jjgFbgcLjgcLjcjService;

    @Autowired
    private JjgFbgcLjgcLjtsfysdHtService jjgFbgcLjgcLjtsfysdHtService;

    @Autowired
    private JjgFbgcLjgcLjtsfysdSlService jjgFbgcLjgcLjtsfysdSlService;

    @Autowired
    private JjgFbgcLjgcLjwcService jjgFbgcLjgcLjwcService;

    @Autowired
    private JjgFbgcLjgcLjwcLcfService jjgFbgcLjgcLjwcLcfService;

    @Autowired
    private JjgFbgcLjgcPsdmccService jjgFbgcLjgcPsdmccService;

    @Autowired
    private JjgFbgcLjgcPspqhdService jjgFbgcLjgcPspqhdService;

    @Autowired
    private JjgFbgcLjgcXqgqdService jjgFbgcLjgcXqgqdService;

    @Autowired
    private JjgFbgcLjgcXqjgccService jjgFbgcLjgcXqjgccService;

    @Autowired
    private JjgFbgcLjgcZddmccService jjgFbgcLjgcZddmccService;

    @Autowired
    private JjgFbgcLjgcZdgqdService jjgFbgcLjgcZdgqdService;

    @Autowired
    private JjgFbgcLmgcGslqlmhdzxfService jjgFbgcLmgcGslqlmhdzxfService;

    @Autowired
    private JjgFbgcLmgcHntlmhdzxfService jjgFbgcLmgcHntlmhdzxfService;

    @Autowired
    private JjgFbgcLmgcHntlmqdService jjgFbgcLmgcHntlmqdService;

    @Autowired
    private JjgFbgcLmgcLmgzsdsgpsfService jjgFbgcLmgcLmgzsdsgpsfService;

    @Autowired
    private JjgFbgcLmgcLmhpService jjgFbgcLmgcLmhpService;

    @Autowired
    private JjgFbgcLmgcLmssxsService jjgFbgcLmgcLmssxsService;

    @Autowired
    private JjgFbgcLmgcLmwcService jjgFbgcLmgcLmwcService;

    @Autowired
    private JjgFbgcLmgcLmwcLcfService jjgFbgcLmgcLmwcLcfService;

    @Autowired
    private JjgFbgcLmgcLqlmysdService jjgFbgcLmgcLqlmysdService;

    @Autowired
    private JjgFbgcLmgcTlmxlbgcService jjgFbgcLmgcTlmxlbgcService;

    @Autowired
    private JjgFbgcJtaqssJabxService jjgFbgcJtaqssJabxService;

    @Autowired
    private JjgFbgcJtaqssJabxfhlService jjgFbgcJtaqssJabxfhlService;

    @Autowired
    private JjgFbgcJtaqssJabzService jjgFbgcJtaqssJabzService;

    @Autowired
    private JjgFbgcJtaqssJathldmccService jjgFbgcJtaqssJathldmccService;

    @Autowired
    private JjgFbgcJtaqssJathlqdService jjgFbgcJtaqssJathlqdService;

    @Autowired
    private JjgFbgcQlgcXbTqdService jjgFbgcQlgcXbTqdService;

    @Autowired
    private JjgFbgcQlgcXbJgccService jjgFbgcQlgcXbJgccService;

    @Autowired
    private JjgFbgcQlgcXbBhchdService jjgFbgcQlgcXbBhchdService;

    @Autowired
    private JjgFbgcQlgcXbSzdService jjgFbgcQlgcXbSzdService;

    @Autowired
    private JjgFbgcQlgcSbTqdService jjgFbgcQlgcSbTqdService;

    @Autowired
    private JjgFbgcQlgcSbJgccService jjgFbgcQlgcSbJgccService;

    @Autowired
    private JjgFbgcQlgcSbBhchdService jjgFbgcQlgcSbBhchdService;

    @Autowired
    private JjgFbgcSdgcCqtqdService jjgFbgcSdgcCqtqdService;

    @Autowired
    private JjgFbgcSdgcDmpzdService jjgFbgcSdgcDmpzdService;

    @Autowired
    private JjgFbgcSdgcZtkdService jjgFbgcSdgcZtkdService;

    @Autowired
    private JjgFbgcQlgcQmpzdService jjgFbgcQlgcQmpzdService;

    @Autowired
    private JjgFbgcQlgcQmhpService jjgFbgcQlgcQmhpService;

    @Autowired
    private JjgFbgcQlgcQmgzsdService jjgFbgcQlgcQmgzsdService;

    @Autowired
    private JjgFbgcSdgcCqhdService jjgFbgcSdgcCqhdService;

    @Autowired
    private JjgFbgcSdgcSdlqlmysdService jjgFbgcSdgcSdlqlmysdService;

    @Autowired
    private JjgFbgcSdgcLmssxsService jjgFbgcSdgcLmssxsService;

    @Autowired
    private JjgFbgcSdgcHntlmqdService jjgFbgcSdgcHntlmqdService;

    @Autowired
    private JjgFbgcSdgcTlmxlbgcService jjgFbgcSdgcTlmxlbgcService;

    @Autowired
    private JjgFbgcSdgcSdhpService jjgFbgcSdgcSdhpService;

    @Autowired
    private JjgFbgcSdgcSdhntlmhdzxfService jjgFbgcSdgcSdhntlmhdzxfService;

    @Autowired
    private JjgFbgcSdgcGssdlqlmhdzxfService jjgFbgcSdgcGssdlqlmhdzxfService;

    @Autowired
    private JjgFbgcSdgcLmgzsdsgpsfService jjgFbgcSdgcLmgzsdsgpsfService;


    @Override
    public void deleteOtherInfo(List<String> idList) {//idList是项目的id
        //先删菜单信息
        for (String id : idList) {
            //根据id查询项目信息
            QueryWrapper<Project> wrapper=new QueryWrapper<>();
            wrapper.eq("id",id);
            Project project = projectMapper.selectOne(wrapper);
            String proName = project.getProName();
            SysMenu sysMenu = sysMenuService.selectcdinfo(proName);
            Long proid = sysMenu.getId();
            //实测数据，合同段信息，路桥隧信息

            //先删实测数据下面的  name=实测数据，parent_id=proid
            SysMenu scChildrenMenu = sysMenuService.getscChildrenMenu(proid);
            Long scid = scChildrenMenu.getId();
            //合同段的菜单数据
            List<SysMenu> htdlist = sysMenuService.getAllHtd(scid);
            log.info("删除{}的分部工程和合同段菜单",project.getProName());
            sysMenuService.removeFbgc(htdlist);//至此，分部工程和合同段的数据删除了
            //然后删除 实测数据，合同段信息，路桥隧信息
            log.info("删除{}的实测数据，合同段信息和路桥隧信息菜单",project.getProName());
            sysMenuService.delectChildrenMenu(proid);
            log.info("删除{}的菜单",project.getProName());
            sysMenuService.removeById(proid);

            log.info("删除{}的在合同段表中的数据",proName);
            QueryWrapper<JjgHtd> htdwrapper=new QueryWrapper<>();
            htdwrapper.eq("proname",proName);
            jjgHtdService.remove(htdwrapper);

            log.info("删除{}的在路桥隧表中的数据",proName);
            QueryWrapper<JjgLqsQl> qlwrapper=new QueryWrapper<>();
            qlwrapper.eq("proname",proName);
            jjgLqsQlService.remove(qlwrapper);
            QueryWrapper<JjgLqsSd> sdwrapper=new QueryWrapper<>();
            sdwrapper.eq("proname",proName);
            jjgLqsSdService.remove(sdwrapper);
            QueryWrapper<JjgLqsFhlm> fhlmwrapper=new QueryWrapper<>();
            fhlmwrapper.eq("proname",proName);
            jjgLqsFhlmService.remove(fhlmwrapper);
            QueryWrapper<JjgLqsHntlmzd> hntlmzdwrapper=new QueryWrapper<>();
            hntlmzdwrapper.eq("proname",proName);
            jjgLqsHntlmzdService.remove(hntlmzdwrapper);
            QueryWrapper<JjgSfz> sfzwrapper=new QueryWrapper<>();
            sfzwrapper.eq("proname",proName);
            jjgLqsSfzService.remove(sfzwrapper);
            log.info("删除{}在项目表中的数据",proName);
            projectMapper.deleteById(id);

            log.info("删除{}在各个表中的实测数据",proName);
            QueryWrapper<JjgFbgcLmgcGslqlmhdzxf> gslqlmhdzxfWrapper=new QueryWrapper<>();
            gslqlmhdzxfWrapper.eq("proname",proName);
            jjgFbgcLmgcGslqlmhdzxfService.remove(gslqlmhdzxfWrapper);

            QueryWrapper<JjgFbgcLmgcHntlmhdzxf> hntlmhdzxfWrapper=new QueryWrapper<>();
            hntlmhdzxfWrapper.eq("proname",proName);
            jjgFbgcLmgcHntlmhdzxfService.remove(hntlmhdzxfWrapper);

            QueryWrapper<JjgFbgcLmgcHntlmqd> hntlmqdWrapper=new QueryWrapper<>();
            hntlmqdWrapper.eq("proname",proName);
            jjgFbgcLmgcHntlmqdService.remove(hntlmqdWrapper);

            QueryWrapper<JjgFbgcLmgcLmgzsdsgpsf> lmgzsdsgpsfWrapper=new QueryWrapper<>();
            lmgzsdsgpsfWrapper.eq("proname",proName);
            jjgFbgcLmgcLmgzsdsgpsfService.remove(lmgzsdsgpsfWrapper);

            QueryWrapper<JjgFbgcLmgcLmhp> lmhpWrapper=new QueryWrapper<>();
            lmhpWrapper.eq("proname",proName);
            jjgFbgcLmgcLmhpService.remove(lmhpWrapper);

            QueryWrapper<JjgFbgcLmgcLmssxs> lmssxsWrapper=new QueryWrapper<>();
            lmssxsWrapper.eq("proname",proName);
            jjgFbgcLmgcLmssxsService.remove(lmssxsWrapper);

            QueryWrapper<JjgFbgcLmgcLmwc> lmwcWrapper=new QueryWrapper<>();
            lmwcWrapper.eq("proname",proName);
            jjgFbgcLmgcLmwcService.remove(lmwcWrapper);

            QueryWrapper<JjgFbgcLmgcLmwcLcf> lmwclcfWrapper=new QueryWrapper<>();
            lmwclcfWrapper.eq("proname",proName);
            jjgFbgcLmgcLmwcLcfService.remove(lmwclcfWrapper);

            QueryWrapper<JjgFbgcLmgcLqlmysd> lqlmysdWrapper=new QueryWrapper<>();
            lqlmysdWrapper.eq("proname",proName);
            jjgFbgcLmgcLqlmysdService.remove(lqlmysdWrapper);

            QueryWrapper<JjgFbgcLmgcTlmxlbgc> tlmxlbgcWrapper=new QueryWrapper<>();
            tlmxlbgcWrapper.eq("proname",proName);
            jjgFbgcLmgcTlmxlbgcService.remove(tlmxlbgcWrapper);

            QueryWrapper<JjgFbgcLjgcHdgqd> hdgqdWrapper=new QueryWrapper<>();
            hdgqdWrapper.eq("proname",proName);
            jjgFbgcLjgcHdgqdService.remove(hdgqdWrapper);

            QueryWrapper<JjgFbgcLjgcHdjgcc> hdjgccWrapper=new QueryWrapper<>();
            hdjgccWrapper.eq("proname",proName);
            jjgFbgcLjgcHdjgccService.remove(hdjgccWrapper);

            QueryWrapper<JjgFbgcLjgcLjbp> ljbpWrapper=new QueryWrapper<>();
            ljbpWrapper.eq("proname",proName);
            jjgFbgcLjgcLjbpService.remove(ljbpWrapper);

            QueryWrapper<JjgFbgcLjgcLjcj> ljcjWrapper=new QueryWrapper<>();
            ljcjWrapper.eq("proname",proName);
            jjgFbgcLjgcLjcjService.remove(ljcjWrapper);

            QueryWrapper<JjgFbgcLjgcLjtsfysdHt> ljtsfysdHtWrapper=new QueryWrapper<>();
            ljtsfysdHtWrapper.eq("proname",proName);
            jjgFbgcLjgcLjtsfysdHtService.remove(ljtsfysdHtWrapper);

            QueryWrapper<JjgFbgcLjgcLjtsfysdSl> ljtsfysdslWrapper=new QueryWrapper<>();
            ljtsfysdslWrapper.eq("proname",proName);
            jjgFbgcLjgcLjtsfysdSlService.remove(ljtsfysdslWrapper);

            QueryWrapper<JjgFbgcLjgcLjwc> ljwcWrapper=new QueryWrapper<>();
            ljwcWrapper.eq("proname",proName);
            jjgFbgcLjgcLjwcService.remove(ljwcWrapper);

            QueryWrapper<JjgFbgcLjgcLjwcLcf> ljwclcfWrapper=new QueryWrapper<>();
            ljwclcfWrapper.eq("proname",proName);
            jjgFbgcLjgcLjwcLcfService.remove(ljwclcfWrapper);

            QueryWrapper<JjgFbgcLjgcPsdmcc> psdmccWrapper=new QueryWrapper<>();
            psdmccWrapper.eq("proname",proName);
            jjgFbgcLjgcPsdmccService.remove(psdmccWrapper);

            QueryWrapper<JjgFbgcLjgcPspqhd> pspqhdWrapper=new QueryWrapper<>();
            pspqhdWrapper.eq("proname",proName);
            jjgFbgcLjgcPspqhdService.remove(pspqhdWrapper);

            QueryWrapper<JjgFbgcLjgcXqgqd> xqgqdWrapper=new QueryWrapper<>();
            xqgqdWrapper.eq("proname",proName);
            jjgFbgcLjgcXqgqdService.remove(xqgqdWrapper);

            QueryWrapper<JjgFbgcLjgcXqjgcc> xqjgccWrapper=new QueryWrapper<>();
            xqjgccWrapper.eq("proname",proName);
            jjgFbgcLjgcXqjgccService.remove(xqjgccWrapper);

            QueryWrapper<JjgFbgcLjgcZddmcc> zddmccWrapper=new QueryWrapper<>();
            zddmccWrapper.eq("proname",proName);
            jjgFbgcLjgcZddmccService.remove(zddmccWrapper);

            QueryWrapper<JjgFbgcLjgcZdgqd> zdgqdWrapper=new QueryWrapper<>();
            zdgqdWrapper.eq("proname",proName);
            jjgFbgcLjgcZdgqdService.remove(zdgqdWrapper);

            QueryWrapper<JjgFbgcJtaqssJabx> jabxWrapper=new QueryWrapper<>();
            jabxWrapper.eq("proname",proName);
            jjgFbgcJtaqssJabxService.remove(jabxWrapper);

            QueryWrapper<JjgFbgcJtaqssJabxfhl> jabxfhlWrapper=new QueryWrapper<>();
            jabxfhlWrapper.eq("proname",proName);
            jjgFbgcJtaqssJabxfhlService.remove(jabxfhlWrapper);

            QueryWrapper<JjgFbgcJtaqssJabz> jabzWrapper=new QueryWrapper<>();
            jabzWrapper.eq("proname",proName);
            jjgFbgcJtaqssJabzService.remove(jabzWrapper);

            QueryWrapper<JjgFbgcJtaqssJathldmcc> jathldmccWrapper=new QueryWrapper<>();
            jathldmccWrapper.eq("proname",proName);
            jjgFbgcJtaqssJathldmccService.remove(jathldmccWrapper);

            QueryWrapper<JjgFbgcJtaqssJathlqd> jathlqdWrapper=new QueryWrapper<>();
            jathlqdWrapper.eq("proname",proName);
            jjgFbgcJtaqssJathlqdService.remove(jathlqdWrapper);

            QueryWrapper<JjgFbgcQlgcXbTqd> xbtqdWrapper=new QueryWrapper<>();
            xbtqdWrapper.eq("proname",proName);
            jjgFbgcQlgcXbTqdService.remove(xbtqdWrapper);

            QueryWrapper<JjgFbgcQlgcXbJgcc> xbjgccWrapper=new QueryWrapper<>();
            xbjgccWrapper.eq("proname",proName);
            jjgFbgcQlgcXbJgccService.remove(xbjgccWrapper);

            QueryWrapper<JjgFbgcQlgcXbBhchd> xbBhchdWrapper=new QueryWrapper<>();
            xbBhchdWrapper.eq("proname",proName);
            jjgFbgcQlgcXbBhchdService.remove(xbBhchdWrapper);

            QueryWrapper<JjgFbgcQlgcXbSzd> xbSzdWrapper=new QueryWrapper<>();
            xbSzdWrapper.eq("proname",proName);
            jjgFbgcQlgcXbSzdService.remove(xbSzdWrapper);

            QueryWrapper<JjgFbgcQlgcSbTqd> sbTqdWrapper=new QueryWrapper<>();
            sbTqdWrapper.eq("proname",proName);
            jjgFbgcQlgcSbTqdService.remove(sbTqdWrapper);

            QueryWrapper<JjgFbgcQlgcSbJgcc> sbJgccWrapper=new QueryWrapper<>();
            sbJgccWrapper.eq("proname",proName);
            jjgFbgcQlgcSbJgccService.remove(sbJgccWrapper);

            QueryWrapper<JjgFbgcQlgcSbBhchd> sbBhchdWrapper=new QueryWrapper<>();
            sbBhchdWrapper.eq("proname",proName);
            jjgFbgcQlgcSbBhchdService.remove(sbBhchdWrapper);

            QueryWrapper<JjgFbgcQlgcQmpzd> qmpzdWrapper=new QueryWrapper<>();
            qmpzdWrapper.eq("proname",proName);
            jjgFbgcQlgcQmpzdService.remove(qmpzdWrapper);

            QueryWrapper<JjgFbgcQlgcQmhp> qmhpWrapper=new QueryWrapper<>();
            qmhpWrapper.eq("proname",proName);
            jjgFbgcQlgcQmhpService.remove(qmhpWrapper);

            QueryWrapper<JjgFbgcQlgcQmgzsd> qmgzsdWrapper=new QueryWrapper<>();
            qmgzsdWrapper.eq("proname",proName);
            jjgFbgcQlgcQmgzsdService.remove(qmgzsdWrapper);

            QueryWrapper<JjgFbgcSdgcCqtqd> cqtqdWrapper=new QueryWrapper<>();
            cqtqdWrapper.eq("proname",proName);
            jjgFbgcSdgcCqtqdService.remove(cqtqdWrapper);

            QueryWrapper<JjgFbgcSdgcDmpzd> dmpzdWrapper=new QueryWrapper<>();
            dmpzdWrapper.eq("proname",proName);
            jjgFbgcSdgcDmpzdService.remove(dmpzdWrapper);

            QueryWrapper<JjgFbgcSdgcZtkd> ztkdQueryWrapper=new QueryWrapper<>();
            ztkdQueryWrapper.eq("proname",proName);
            jjgFbgcSdgcZtkdService.remove(ztkdQueryWrapper);

            QueryWrapper<JjgFbgcSdgcCqhd> cqhdQueryWrapper=new QueryWrapper<>();
            cqhdQueryWrapper.eq("proname",proName);
            jjgFbgcSdgcCqhdService.remove(cqhdQueryWrapper);

            QueryWrapper<JjgFbgcSdgcSdlqlmysd> sdlqlmysdQueryWrapper=new QueryWrapper<>();
            sdlqlmysdQueryWrapper.eq("proname",proName);
            jjgFbgcSdgcSdlqlmysdService.remove(sdlqlmysdQueryWrapper);

            QueryWrapper<JjgFbgcSdgcLmssxs> lmssxsQueryWrapper=new QueryWrapper<>();
            lmssxsQueryWrapper.eq("proname",proName);
            jjgFbgcSdgcLmssxsService.remove(lmssxsQueryWrapper);

            QueryWrapper<JjgFbgcSdgcHntlmqd> hntlmqdQueryWrapper=new QueryWrapper<>();
            hntlmqdQueryWrapper.eq("proname",proName);
            jjgFbgcSdgcHntlmqdService.remove(hntlmqdQueryWrapper);

            QueryWrapper<JjgFbgcSdgcTlmxlbgc> tlmxlbgcQueryWrapper=new QueryWrapper<>();
            tlmxlbgcQueryWrapper.eq("proname",proName);
            jjgFbgcSdgcTlmxlbgcService.remove(tlmxlbgcQueryWrapper);

            QueryWrapper<JjgFbgcSdgcSdhp> sdhpQueryWrapper=new QueryWrapper<>();
            sdhpQueryWrapper.eq("proname",proName);
            jjgFbgcSdgcSdhpService.remove(sdhpQueryWrapper);

            QueryWrapper<JjgFbgcSdgcSdhntlmhdzxf> sdhntlmhdzxfQueryWrapper=new QueryWrapper<>();
            sdhntlmhdzxfQueryWrapper.eq("proname",proName);
            jjgFbgcSdgcSdhntlmhdzxfService.remove(sdhntlmhdzxfQueryWrapper);

            QueryWrapper<JjgFbgcSdgcGssdlqlmhdzxf> gssdlqlmhdzxfQueryWrapper=new QueryWrapper<>();
            gssdlqlmhdzxfQueryWrapper.eq("proname",proName);
            jjgFbgcSdgcGssdlqlmhdzxfService.remove(gssdlqlmhdzxfQueryWrapper);

            QueryWrapper<JjgFbgcSdgcLmgzsdsgpsf> lmgzsdsgpsfQueryWrapper=new QueryWrapper<>();
            lmgzsdsgpsfQueryWrapper.eq("proname",proName);
            jjgFbgcSdgcLmgzsdsgpsfService.remove(lmgzsdsgpsfQueryWrapper);

        }
    }

}
