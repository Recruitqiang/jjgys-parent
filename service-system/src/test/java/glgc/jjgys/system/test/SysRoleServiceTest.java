package glgc.jjgys.system.test;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.WriteTable;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabx;
import glgc.jjgys.model.project.JjgFbgcLjgcHdgqd;
import glgc.jjgys.model.project.JjgFbgcLjgcHdjgcc;
import glgc.jjgys.model.project.JjgFbgcLjgcXqgqd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysRole;
import glgc.jjgys.system.mapper.JjgFbgcJtaqssJabxMapper;
import glgc.jjgys.system.service.*;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import glgc.jjgys.system.service.impl.JjgFbgcLjgcZdgqdServiceImpl;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@SpringBootTest
public class SysRoleServiceTest {

    //注入service
    @Autowired
    private SysRoleService sysRoleService;

    @Autowired
    private JjgFbgcLjgcHdgqdService jjgFbgcLjgcHdgqdService;
    @Autowired
    private JjgFbgcLjgcHdjgccService jjgFbgcLjgcHdjgccService;

    @Autowired
    private JjgFbgcLjgcXqgqdService jjgFbgcLjgcXqgqdService;

    @Autowired
    private JjgFbgcLjgcXqjgccService jjgFbgcLjgcXqjgccService;

    @Autowired
    private JjgFbgcLjgcZdgqdService jjgFbgcLjgcZdgqdService;

    @Autowired
    private JjgFbgcLjgcZddmccService jjgFbgcLjgcZddmccService;

    @Autowired
    private JjgFbgcLjgcPspqhdService jjgFbgcLjgcPspqhdService;

    @Autowired
    private JjgFbgcLjgcPsdmccService jjgFbgcLjgcPsdmccService;

    @Autowired
    private JjgFbgcLjgcLjbpService jjgFbgcLjgcLjbpService;

    @Autowired
    private JjgFbgcLjgcLjtsfysdHtService jjgFbgcLjgcLjtsfysdHtService;
    @Autowired
    private JjgFbgcLjgcLjcjService jjgFbgcLjgcLjcjService;
    @Autowired
    private JjgFbgcLjgcLjwcService jjgFbgcLjgcLjwcService;

    @Autowired
    private JjgFbgcLjgcLjwcLcfService jjgFbgcLjgcLjwcLcfService;

    @Autowired
    private JjgFbgcJtaqssJabxService jjgFbgcJtaqssJabxService;

    @Autowired
    private JjgFbgcJtaqssJabxMapper jjgFbgcJtaqssJabxMapper;

    @Autowired
    private JjgFbgcJtaqssJabzService jjgFbgcJtaqssJabzService;

    @Autowired
    private JjgFbgcJtaqssJathlqdService jjgFbgcJtaqssJathlqdService;
    @Autowired
    private JjgFbgcJtaqssJathldmccService jjgFbgcJtaqssJathldmccService;
    @Autowired
    private JjgFbgcJtaqssJabxfhlService jjgFbgcJtaqssJabxfhlService;

    @Autowired
    private JjgFbgcQlgcSbTqdService jjgFbgcQlgcSbTqdService;

    @Autowired
    private JjgFbgcQlgcSbJgccService jjgFbgcQlgcSbJgccService;

    @Autowired
    private JjgFbgcQlgcSbBhchdService jjgFbgcQlgcSbBhchdService;

    @Autowired
    private JjgFbgcQlgcXbJgccService jjgFbgcQlgcXbJgccService;

    @Autowired
    private JjgFbgcQlgcXbBhchdService jjgFbgcQlgcXbBhchdService;

    @Autowired
    private JjgFbgcQlgcXbTqdService jjgFbgcQlgcXbTqdService;

    @Autowired
    private JjgFbgcQlgcXbSzdService jjgFbgcQlgcXbSzdService;

    @Autowired
    private JjgFbgcQlgcQmhpService jjgFbgcQlgcQmhpService;

    @Autowired
    private JjgFbgcSdgcZtkdService jjgFbgcSdgcZtkdService;

    @Autowired
    private JjgFbgcSdgcCqtqdService jjgFbgcSdgcCqtqdService;

    @Autowired
    private JjgFbgcSdgcDmpzdService jjgFbgcSdgcDmpzdService;

    @Autowired
    private JjgFbgcSdgcSdhpService jjgFbgcSdgcSdhpService;

    @Autowired
    private JjgFbgcSdgcCqhdService jjgFbgcSdgcCqhdService;

    @Autowired
    private JjgFbgcLmgcLqlmysdService jjgFbgcLmgcLqlmysdService;

    @Autowired
    private JjgFbgcLmgcLmwcService jjgFbgcLmgcLmwcService;

    @Autowired
    private JjgFbgcLmgcLmwcLcfService jjgFbgcLmgcLmwcLcfService;

    @Autowired
    private JjgFbgcLmgcLmssxsService jjgFbgcLmgcLmssxsService;

    @Autowired
    private JjgFbgcLmgcHntlmqdService jjgFbgcLmgcHntlmqdService;

    @Autowired
    private JjgFbgcLmgcTlmxlbgcService jjgFbgcLmgcTlmxlbgcService;

    @Autowired
    private JjgFbgcLmgcLmhpService jjgFbgcLmgcLmhpService;

    @Autowired
    private JjgFbgcLmgcLmgzsdsgpsfService jjgFbgcLmgcLmgzsdsgpsfService;

    @Autowired
    private JjgFbgcQlgcQmpzdService jjgFbgcQlgcQmpzdService;

    @Autowired
    private JjgFbgcSdgcLmssxsService jjgFbgcSdgcLmssxsService;

    @Autowired
    private JjgFbgcSdgcLmgzsdsgpsfService jjgFbgcSdgcLmgzsdsgpsfService;

    @Autowired
    private JjgFbgcQlgcQmgzsdService jjgFbgcQlgcQmgzsdService;

    @Autowired
    private JjgFbgcSdgcHntlmqdService jjgFbgcSdgcHntlmqdService;

    @Autowired
    private JjgFbgcSdgcTlmxlbgcService jjgFbgcSdgcTlmxlbgcService;

    @Autowired
    private JjgFbgcSdgcSdlqlmysdService jjgFbgcSdgcSdlqlmysdService;

    @Autowired
    private JjgFbgcSdgcGssdlqlmhdzxfService jjgFbgcSdgcGssdlqlmhdzxfService;

    @Autowired
    private JjgFbgcLmgcHntlmhdzxfService jjgFbgcLmgcHntlmhdzxfService;
    @Autowired
    private JjgFbgcLmgcGslqlmhdzxfService jjgFbgcLmgcGslqlmhdzxfService;

    @Autowired
    private JjgFbgcSdgcSdhntlmhdzxfService jjgFbgcSdgcSdhntlmhdzxfService;

    @Autowired
    private JjgZdhMcxsService jjgZdhMcxsService;

    @Autowired
    private JjgZdhGzsdService jjgZdhGzsdService;

    @Autowired
    private JjgZdhPzdService jjgZdhPzdService;

    @Autowired
    private JjgZdhLdhdService jjgZdhLdhdService;

    @Autowired
    private JjgFbgcGenerateTablelService jjgFbgcGenerateTablelService;

    @Autowired
    private JjgFbgcGenerateWordService jjgFbgcGenerateWordService;

    @Autowired
    private JjgZdhCzService jjgZdhCzService;

    @Autowired
    private JjgLookProjectPlanService jjgLookProjectPlanService;



    @Test
    public void hdgqdsc() throws Exception {
        /*String s = "YK3+960";
        double v = Double.parseDouble(s.replaceAll("[^0-9]", ""));
        System.out.println(v);*/
        CommonInfoVo commonInfoVo = new CommonInfoVo();
        commonInfoVo.setProname("陕西高速");
        commonInfoVo.setHtd("LJ-1");
        //commonInfoVo.setSjz("0.1");
        /*commonInfoVo.setSjz("22");
        /*commonInfoVo.setSjz("22");
        commonInfoVo.setQlsjz("10");
        commonInfoVo.setSdsjz("10");
        commonInfoVo.setFhlmsjz("10");*/
        //commonInfoVo.setFbgc("路面工程");

        //jjgLookProjectPlanService.lookplan(commonInfoVo);


        /*QueryWrapper<JjgFbgcJtaqssJabx> wrapper=new QueryWrapper<>();
        wrapper.like("proname","陕西高速");
        wrapper.like("htd","LJ-1");
        wrapper.like("fbgc","标线");
        wrapper.orderByDesc("wz","hdscz1");
        List<JjgFbgcJtaqssJabx> data = jjgFbgcJtaqssJabxMapper.selectList(wrapper);
        jjgFbgcJtaqssJabxService.bxnfsxs(data);*/
        //jjgZdhGzsdService.generateJdb(commonInfoVo);

        //jjgFbgcGenerateTablelService.generateBGZBG("陕西高速");
        jjgFbgcGenerateWordService.generateword("陕西高速");


        //List<Map<String, Object>> maps = jjgFbgcSdgcGssdlqlmhdzxfService.lookJdbjg(commonInfoVo);
        //jjgFbgcSdgcZtkdService.generateJdb(commonInfoVo);
        //jjgFbgcLjgcLjcjService.generateJdb(commonInfoVo);
        //List<Map<String, Object>> maps = jjgFbgcLjgcLjtsfysdHtService.lookJdbjg(commonInfoVo);
        //System.out.println(maps);
    }


}
