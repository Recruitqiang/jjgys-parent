package glgc.jjgys.system.service.impl;

import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.model.project.*;
import glgc.jjgys.model.system.SysMenu;
import glgc.jjgys.system.mapper.JjgHtdMapper;
import glgc.jjgys.system.mapper.SysMenuMapper;
import glgc.jjgys.system.service.*;
import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.math.BigInteger;
import java.text.DecimalFormat;
import java.util.*;

@Service
public class JjgHtdServiceImpl extends ServiceImpl<JjgHtdMapper, JjgHtd> implements JjgHtdService {

    @Autowired
    private SysMenuService sysMenuService;

    @Autowired
    private SysMenuMapper sysMenuMapper;

    @Autowired
    private JjgHtdMapper jjgHtdMapper;

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
        long htdID = uniqId();
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

    @Override
    public JjgHtd selectlx(String proname, String htd) {
        QueryWrapper<JjgHtd> wrapperhtd = new QueryWrapper<>();
        wrapperhtd.like("proname", proname);
        wrapperhtd.like("name", htd);
        JjgHtd jjgHtd = jjgHtdMapper.selectOne(wrapperhtd);
        return jjgHtd;
    }

    @Override
    public JjgHtd selectInfo(String proname, String htd) {
        QueryWrapper<JjgHtd> wrapperhtd = new QueryWrapper<>();
        wrapperhtd.like("proname", proname);
        wrapperhtd.like("name", htd);
        JjgHtd jjgHtd = jjgHtdMapper.selectOne(wrapperhtd);
        return jjgHtd;
    }

    @Override
    public String getAllzh(String proname) {
        QueryWrapper<JjgHtd> wrapperhtd = new QueryWrapper<>();
        wrapperhtd.like("proname", proname);
        List<JjgHtd> htdList = jjgHtdMapper.selectList(wrapperhtd);
        double minZhq = 0;
        double maxZhz = 0;
        for (JjgHtd jjgHtd : htdList) {
            Double value1 = Double.valueOf(jjgHtd.getZhq());
            Double value2 = Double.valueOf(jjgHtd.getZhz());
            if (value1 < minZhq) {
                minZhq = value1;
            }
            if (value2 > maxZhz) {
                maxZhz = value2;
            }
        }
        String getgcbw = getgcbw(minZhq, maxZhz);
        return getgcbw;

    }

    @Override
    public List<JjgHtd> gethtd(String proname) {
        /*QueryWrapper<JjgHtd> wrapperhtd = new QueryWrapper<>();
        wrapperhtd.select(proname);
        wrapperhtd.like("proname", proname);
        List<JjgHtd> htdList = jjgHtdMapper.selectList(wrapperhtd);*/
        List<JjgHtd> htdList = jjgHtdMapper.selecthtd(proname);
        return htdList;
    }

    @Override
    public boolean removeData(List<String> idList) {
        for (String s : idList) {
            QueryWrapper<JjgHtd> wrapper = new QueryWrapper<>();
            wrapper.eq("id", s);
            List<JjgHtd> htdList = jjgHtdMapper.selectList(wrapper);
            for (JjgHtd jjgHtd : htdList) {
                String proname = jjgHtd.getProname();
                String htd = jjgHtd.getName();
                //jjg_htdinfo已删，
                /**
                 * 先删分部工程,
                 * 根据项目名称在sys_menu查询，根据这个项目的id,找到子菜单的name是“实测数据”的菜单，然后将“实测数据”下的子菜单关于这个合同段信息删除
                 */
                //String proname = htdList.get(0).getProname();
                //String htd = htdList.get(0).getName();
                //删除sys_menu中的菜单数据
                List<SysMenu> menuProname = sysMenuService.selectname(proname);//根据项目名称拿到id
                Long pronameid = menuProname.get(0).getId();//陕西高速id
                List<SysMenu> menusc = sysMenuService.selectscname(pronameid);//根据项目id，去查关于parent_id的实测数据菜单
                Long scid = menusc.get(0).getId();//实测数据id
                List<SysMenu> fbgc = sysMenuService.selecthtd(scid,htd);
                for (SysMenu sysMenu : fbgc) {
                    Long id = sysMenu.getId();
                    sysMenuService.delectfbgc(id);
                }
                sysMenuService.delecthtd(scid,htd);
                //根据分部工程删表中的数据
                String lx = jjgHtd.getLx();
                delectTable(proname,htd,lx);
            }
        }
        return true;

    }

    private void delectTable(String proname, String htd, String lx) {
        String[] split = lx.split(",");
        for (String s : split) {
            if (s.equals("路面工程")){
                QueryWrapper<JjgFbgcLmgcGslqlmhdzxf> gslqlmhdzxfWrapper=new QueryWrapper<>();
                gslqlmhdzxfWrapper.eq("proname",proname);
                gslqlmhdzxfWrapper.eq("htd",htd);
                jjgFbgcLmgcGslqlmhdzxfService.remove(gslqlmhdzxfWrapper);

                QueryWrapper<JjgFbgcLmgcHntlmhdzxf> hntlmhdzxfWrapper=new QueryWrapper<>();
                hntlmhdzxfWrapper.eq("proname",proname);
                hntlmhdzxfWrapper.eq("htd",htd);
                jjgFbgcLmgcHntlmhdzxfService.remove(hntlmhdzxfWrapper);

                QueryWrapper<JjgFbgcLmgcHntlmqd> hntlmqdWrapper=new QueryWrapper<>();
                hntlmqdWrapper.eq("proname",proname);
                hntlmqdWrapper.eq("htd",htd);
                jjgFbgcLmgcHntlmqdService.remove(hntlmqdWrapper);

                QueryWrapper<JjgFbgcLmgcLmgzsdsgpsf> lmgzsdsgpsfWrapper=new QueryWrapper<>();
                lmgzsdsgpsfWrapper.eq("proname",proname);
                lmgzsdsgpsfWrapper.eq("htd",htd);
                jjgFbgcLmgcLmgzsdsgpsfService.remove(lmgzsdsgpsfWrapper);

                QueryWrapper<JjgFbgcLmgcLmhp> lmhpWrapper=new QueryWrapper<>();
                lmhpWrapper.eq("proname",proname);
                lmhpWrapper.eq("htd",htd);
                jjgFbgcLmgcLmhpService.remove(lmhpWrapper);

                QueryWrapper<JjgFbgcLmgcLmssxs> lmssxsWrapper=new QueryWrapper<>();
                lmssxsWrapper.eq("proname",proname);
                lmssxsWrapper.eq("htd",htd);
                jjgFbgcLmgcLmssxsService.remove(lmssxsWrapper);

                QueryWrapper<JjgFbgcLmgcLmwc> lmwcWrapper=new QueryWrapper<>();
                lmwcWrapper.eq("proname",proname);
                lmwcWrapper.eq("htd",htd);
                jjgFbgcLmgcLmwcService.remove(lmwcWrapper);

                QueryWrapper<JjgFbgcLmgcLmwcLcf> lmwclcfWrapper=new QueryWrapper<>();
                lmwclcfWrapper.eq("proname",proname);
                lmwclcfWrapper.eq("htd",htd);
                jjgFbgcLmgcLmwcLcfService.remove(lmwclcfWrapper);

                QueryWrapper<JjgFbgcLmgcLqlmysd> lqlmysdWrapper=new QueryWrapper<>();
                lqlmysdWrapper.eq("proname",proname);
                lqlmysdWrapper.eq("htd",htd);
                jjgFbgcLmgcLqlmysdService.remove(lqlmysdWrapper);

                QueryWrapper<JjgFbgcLmgcTlmxlbgc> tlmxlbgcWrapper=new QueryWrapper<>();
                tlmxlbgcWrapper.eq("proname",proname);
                tlmxlbgcWrapper.eq("htd",htd);
                jjgFbgcLmgcTlmxlbgcService.remove(tlmxlbgcWrapper);

            }else if (s.equals("路基工程")){
                QueryWrapper<JjgFbgcLjgcHdgqd> hdgqdWrapper=new QueryWrapper<>();
                hdgqdWrapper.eq("proname",proname);
                hdgqdWrapper.eq("htd",htd);
                jjgFbgcLjgcHdgqdService.remove(hdgqdWrapper);

                QueryWrapper<JjgFbgcLjgcHdjgcc> hdjgccWrapper=new QueryWrapper<>();
                hdjgccWrapper.eq("proname",proname);
                hdjgccWrapper.eq("htd",htd);
                jjgFbgcLjgcHdjgccService.remove(hdjgccWrapper);

                QueryWrapper<JjgFbgcLjgcLjbp> ljbpWrapper=new QueryWrapper<>();
                ljbpWrapper.eq("proname",proname);
                ljbpWrapper.eq("htd",htd);
                jjgFbgcLjgcLjbpService.remove(ljbpWrapper);

                QueryWrapper<JjgFbgcLjgcLjcj> ljcjWrapper=new QueryWrapper<>();
                ljcjWrapper.eq("proname",proname);
                ljcjWrapper.eq("htd",htd);
                jjgFbgcLjgcLjcjService.remove(ljcjWrapper);

                QueryWrapper<JjgFbgcLjgcLjtsfysdHt> ljtsfysdHtWrapper=new QueryWrapper<>();
                ljtsfysdHtWrapper.eq("proname",proname);
                ljtsfysdHtWrapper.eq("htd",htd);
                jjgFbgcLjgcLjtsfysdHtService.remove(ljtsfysdHtWrapper);

                QueryWrapper<JjgFbgcLjgcLjtsfysdSl> ljtsfysdslWrapper=new QueryWrapper<>();
                ljtsfysdslWrapper.eq("proname",proname);
                ljtsfysdslWrapper.eq("htd",htd);
                jjgFbgcLjgcLjtsfysdSlService.remove(ljtsfysdslWrapper);

                QueryWrapper<JjgFbgcLjgcLjwc> ljwcWrapper=new QueryWrapper<>();
                ljwcWrapper.eq("proname",proname);
                ljwcWrapper.eq("htd",htd);
                jjgFbgcLjgcLjwcService.remove(ljwcWrapper);

                QueryWrapper<JjgFbgcLjgcLjwcLcf> ljwclcfWrapper=new QueryWrapper<>();
                ljwclcfWrapper.eq("proname",proname);
                ljwclcfWrapper.eq("htd",htd);
                jjgFbgcLjgcLjwcLcfService.remove(ljwclcfWrapper);

                QueryWrapper<JjgFbgcLjgcPsdmcc> psdmccWrapper=new QueryWrapper<>();
                psdmccWrapper.eq("proname",proname);
                psdmccWrapper.eq("htd",htd);
                jjgFbgcLjgcPsdmccService.remove(psdmccWrapper);

                QueryWrapper<JjgFbgcLjgcPspqhd> pspqhdWrapper=new QueryWrapper<>();
                pspqhdWrapper.eq("proname",proname);
                pspqhdWrapper.eq("htd",htd);
                jjgFbgcLjgcPspqhdService.remove(pspqhdWrapper);

                QueryWrapper<JjgFbgcLjgcXqgqd> xqgqdWrapper=new QueryWrapper<>();
                xqgqdWrapper.eq("proname",proname);
                xqgqdWrapper.eq("htd",htd);
                jjgFbgcLjgcXqgqdService.remove(xqgqdWrapper);

                QueryWrapper<JjgFbgcLjgcXqjgcc> xqjgccWrapper=new QueryWrapper<>();
                xqjgccWrapper.eq("proname",proname);
                xqjgccWrapper.eq("htd",htd);
                jjgFbgcLjgcXqjgccService.remove(xqjgccWrapper);

                QueryWrapper<JjgFbgcLjgcZddmcc> zddmccWrapper=new QueryWrapper<>();
                zddmccWrapper.eq("proname",proname);
                zddmccWrapper.eq("htd",htd);
                jjgFbgcLjgcZddmccService.remove(zddmccWrapper);

                QueryWrapper<JjgFbgcLjgcZdgqd> zdgqdWrapper=new QueryWrapper<>();
                zdgqdWrapper.eq("proname",proname);
                zdgqdWrapper.eq("htd",htd);
                jjgFbgcLjgcZdgqdService.remove(zdgqdWrapper);

            }else if (s.equals("交安工程")){
                QueryWrapper<JjgFbgcJtaqssJabx> jabxWrapper=new QueryWrapper<>();
                jabxWrapper.eq("proname",proname);
                jabxWrapper.eq("htd",htd);
                jjgFbgcJtaqssJabxService.remove(jabxWrapper);

                QueryWrapper<JjgFbgcJtaqssJabxfhl> jabxfhlWrapper=new QueryWrapper<>();
                jabxfhlWrapper.eq("proname",proname);
                jabxfhlWrapper.eq("htd",htd);
                jjgFbgcJtaqssJabxfhlService.remove(jabxfhlWrapper);

                QueryWrapper<JjgFbgcJtaqssJabz> jabzWrapper=new QueryWrapper<>();
                jabzWrapper.eq("proname",proname);
                jabzWrapper.eq("htd",htd);
                jjgFbgcJtaqssJabzService.remove(jabzWrapper);

                QueryWrapper<JjgFbgcJtaqssJathldmcc> jathldmccWrapper=new QueryWrapper<>();
                jathldmccWrapper.eq("proname",proname);
                jathldmccWrapper.eq("htd",htd);
                jjgFbgcJtaqssJathldmccService.remove(jathldmccWrapper);

                QueryWrapper<JjgFbgcJtaqssJathlqd> jathlqdWrapper=new QueryWrapper<>();
                jathlqdWrapper.eq("proname",proname);
                jathlqdWrapper.eq("htd",htd);
                jjgFbgcJtaqssJathlqdService.remove(jathlqdWrapper);

            }else if (s.equals("桥梁工程")){
                QueryWrapper<JjgFbgcQlgcXbTqd> xbtqdWrapper=new QueryWrapper<>();
                xbtqdWrapper.eq("proname",proname);
                xbtqdWrapper.eq("htd",htd);
                jjgFbgcQlgcXbTqdService.remove(xbtqdWrapper);

                QueryWrapper<JjgFbgcQlgcXbJgcc> xbjgccWrapper=new QueryWrapper<>();
                xbjgccWrapper.eq("proname",proname);
                xbjgccWrapper.eq("htd",htd);
                jjgFbgcQlgcXbJgccService.remove(xbjgccWrapper);

                QueryWrapper<JjgFbgcQlgcXbBhchd> xbBhchdWrapper=new QueryWrapper<>();
                xbBhchdWrapper.eq("proname",proname);
                xbBhchdWrapper.eq("htd",htd);
                jjgFbgcQlgcXbBhchdService.remove(xbBhchdWrapper);

                QueryWrapper<JjgFbgcQlgcXbSzd> xbSzdWrapper=new QueryWrapper<>();
                xbSzdWrapper.eq("proname",proname);
                xbSzdWrapper.eq("htd",htd);
                jjgFbgcQlgcXbSzdService.remove(xbSzdWrapper);

                QueryWrapper<JjgFbgcQlgcSbTqd> sbTqdWrapper=new QueryWrapper<>();
                sbTqdWrapper.eq("proname",proname);
                sbTqdWrapper.eq("htd",htd);
                jjgFbgcQlgcSbTqdService.remove(sbTqdWrapper);

                QueryWrapper<JjgFbgcQlgcSbJgcc> sbJgccWrapper=new QueryWrapper<>();
                sbJgccWrapper.eq("proname",proname);
                sbJgccWrapper.eq("htd",htd);
                jjgFbgcQlgcSbJgccService.remove(sbJgccWrapper);

                QueryWrapper<JjgFbgcQlgcSbBhchd> sbBhchdWrapper=new QueryWrapper<>();
                sbBhchdWrapper.eq("proname",proname);
                sbBhchdWrapper.eq("htd",htd);
                jjgFbgcQlgcSbBhchdService.remove(sbBhchdWrapper);

                QueryWrapper<JjgFbgcQlgcQmpzd> qmpzdWrapper=new QueryWrapper<>();
                qmpzdWrapper.eq("proname",proname);
                qmpzdWrapper.eq("htd",htd);
                jjgFbgcQlgcQmpzdService.remove(qmpzdWrapper);

                QueryWrapper<JjgFbgcQlgcQmhp> qmhpWrapper=new QueryWrapper<>();
                qmhpWrapper.eq("proname",proname);
                qmhpWrapper.eq("htd",htd);
                jjgFbgcQlgcQmhpService.remove(qmhpWrapper);

                QueryWrapper<JjgFbgcQlgcQmgzsd> qmgzsdWrapper=new QueryWrapper<>();
                qmgzsdWrapper.eq("proname",proname);
                qmgzsdWrapper.eq("htd",htd);
                jjgFbgcQlgcQmgzsdService.remove(qmgzsdWrapper);


            }else if (s.equals("隧道工程")){
                QueryWrapper<JjgFbgcSdgcCqtqd> cqtqdWrapper=new QueryWrapper<>();
                cqtqdWrapper.eq("proname",proname);
                cqtqdWrapper.eq("htd",htd);
                jjgFbgcSdgcCqtqdService.remove(cqtqdWrapper);

                QueryWrapper<JjgFbgcSdgcDmpzd> dmpzdWrapper=new QueryWrapper<>();
                dmpzdWrapper.eq("proname",proname);
                dmpzdWrapper.eq("htd",htd);
                jjgFbgcSdgcDmpzdService.remove(dmpzdWrapper);

                QueryWrapper<JjgFbgcSdgcZtkd> ztkdQueryWrapper=new QueryWrapper<>();
                ztkdQueryWrapper.eq("proname",proname);
                ztkdQueryWrapper.eq("htd",htd);
                jjgFbgcSdgcZtkdService.remove(ztkdQueryWrapper);

                QueryWrapper<JjgFbgcSdgcCqhd> cqhdQueryWrapper=new QueryWrapper<>();
                cqhdQueryWrapper.eq("proname",proname);
                cqhdQueryWrapper.eq("htd",htd);
                jjgFbgcSdgcCqhdService.remove(cqhdQueryWrapper);

                QueryWrapper<JjgFbgcSdgcSdlqlmysd> sdlqlmysdQueryWrapper=new QueryWrapper<>();
                sdlqlmysdQueryWrapper.eq("proname",proname);
                sdlqlmysdQueryWrapper.eq("htd",htd);
                jjgFbgcSdgcSdlqlmysdService.remove(sdlqlmysdQueryWrapper);

                QueryWrapper<JjgFbgcSdgcLmssxs> lmssxsQueryWrapper=new QueryWrapper<>();
                lmssxsQueryWrapper.eq("proname",proname);
                lmssxsQueryWrapper.eq("htd",htd);
                jjgFbgcSdgcLmssxsService.remove(lmssxsQueryWrapper);

                QueryWrapper<JjgFbgcSdgcHntlmqd> hntlmqdQueryWrapper=new QueryWrapper<>();
                hntlmqdQueryWrapper.eq("proname",proname);
                hntlmqdQueryWrapper.eq("htd",htd);
                jjgFbgcSdgcHntlmqdService.remove(hntlmqdQueryWrapper);

                QueryWrapper<JjgFbgcSdgcTlmxlbgc> tlmxlbgcQueryWrapper=new QueryWrapper<>();
                tlmxlbgcQueryWrapper.eq("proname",proname);
                tlmxlbgcQueryWrapper.eq("htd",htd);
                jjgFbgcSdgcTlmxlbgcService.remove(tlmxlbgcQueryWrapper);

                QueryWrapper<JjgFbgcSdgcSdhp> sdhpQueryWrapper=new QueryWrapper<>();
                sdhpQueryWrapper.eq("proname",proname);
                sdhpQueryWrapper.eq("htd",htd);
                jjgFbgcSdgcSdhpService.remove(sdhpQueryWrapper);

                QueryWrapper<JjgFbgcSdgcSdhntlmhdzxf> sdhntlmhdzxfQueryWrapper=new QueryWrapper<>();
                sdhntlmhdzxfQueryWrapper.eq("proname",proname);
                sdhntlmhdzxfQueryWrapper.eq("htd",htd);
                jjgFbgcSdgcSdhntlmhdzxfService.remove(sdhntlmhdzxfQueryWrapper);

                QueryWrapper<JjgFbgcSdgcGssdlqlmhdzxf> gssdlqlmhdzxfQueryWrapper=new QueryWrapper<>();
                gssdlqlmhdzxfQueryWrapper.eq("proname",proname);
                gssdlqlmhdzxfQueryWrapper.eq("htd",htd);
                jjgFbgcSdgcGssdlqlmhdzxfService.remove(gssdlqlmhdzxfQueryWrapper);

                QueryWrapper<JjgFbgcSdgcLmgzsdsgpsf> lmgzsdsgpsfQueryWrapper=new QueryWrapper<>();
                lmgzsdsgpsfQueryWrapper.eq("proname",proname);
                lmgzsdsgpsfQueryWrapper.eq("htd",htd);
                jjgFbgcSdgcLmgzsdsgpsfService.remove(lmgzsdsgpsfQueryWrapper);


            }
        }

    }


    private String getgcbw(double a, double b) {
        DecimalFormat decimalFormat = new DecimalFormat("#.###");

        int aa = (int) a / 1000;
        int bb = (int) b / 1000;
        double cc = a % 1000;
        double dd = b % 1000;
        String result = "K"+aa+"+"+decimalFormat.format(cc)+"--"+"K"+bb+"+"+decimalFormat.format(dd);

        return result;
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


    public long uniqId() {
        Random random = new Random();
        String nanoRandom = System.nanoTime() + "" + random.nextInt(99999);
        int hash = Math.abs(UUID.randomUUID().hashCode());
        int needAdd = 19 - String.valueOf(hash).length() + 1;

        BigInteger bigInteger = new BigInteger(hash + "" + nanoRandom.substring(needAdd));
        BigInteger longMaxValue = BigInteger.valueOf(Long.MAX_VALUE);
        BigInteger longMinValue = BigInteger.valueOf(Long.MIN_VALUE);

        // Check if the generated BigInteger is higher than Long.MAX_VALUE
        if (bigInteger.compareTo(longMaxValue) >= 0) {
            bigInteger = bigInteger.mod(longMaxValue);
        }
        // Check if the generated BigInteger is lower than Long.MIN_VALUE
        else if (bigInteger.compareTo(longMinValue) <= 0) {
            bigInteger = bigInteger.mod(longMaxValue).add(longMaxValue);
        }

        long aLong = bigInteger.longValue();

        return aLong;
    }


}
