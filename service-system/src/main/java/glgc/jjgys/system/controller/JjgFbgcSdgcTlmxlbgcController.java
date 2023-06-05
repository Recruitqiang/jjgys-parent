package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcLmgcTlmxlbgc;
import glgc.jjgys.model.project.JjgFbgcSdgcTlmxlbgc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.system.service.JjgFbgcSdgcTlmxlbgcService;
import glgc.jjgys.system.service.OperLogService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.context.request.RequestAttributes;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * <p>
 *  前端控制器
 * </p>
 *
 * @author wq
 * @since 2023-05-04
 */
@RestController
@RequestMapping("/jjg/fbgc/sdgc/tlmxlbgc")
public class JjgFbgcSdgcTlmxlbgcController {

    @Autowired
    private JjgFbgcSdgcTlmxlbgcService jjgFbgcSdgcTlmxlbgcService;

    @Autowired
    private OperLogService operLogService;


    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletRequest request,HttpServletResponse response, String proname, String htd,String fbgc) throws IOException {
        List<Map<String,Object>> sdmclist = jjgFbgcSdgcTlmxlbgcService.selectsdmc(proname,htd,fbgc);
        List list = new ArrayList<>();
        for (int i=0;i<sdmclist.size();i++){
            list.add(sdmclist.get(i).get("sdmc"));
        }
        String zipName = "48隧道混凝土路面相邻板高差";
        JjgFbgcCommonUtils.batchDownloadFile(request,response,zipName,list,filespath+File.separator+proname+File.separator+htd);
    }


    @ApiOperation("生成隧道相邻板高差鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcSdgcTlmxlbgcService.generateJdb(commonInfoVo);

    }
    @ApiOperation("查看隧道相邻板高差鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcSdgcTlmxlbgcService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("隧道相邻板高差模板文件导出")
    @GetMapping("exportsdxlbgs")
    public void exportsdxlbgs(HttpServletResponse response){
        jjgFbgcSdgcTlmxlbgcService.exportsdxlbgs(response);
    }


    @ApiOperation(value = "隧道相邻板高差数据文件导入")
    @PostMapping("importsdxlbgs")
    public Result importsdxlbgs(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcSdgcTlmxlbgcService.importsdxlbgs(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcSdgcTlmxlbgc jjgFbgcSdgcTlmxlbgc){
        //创建page对象
        Page<JjgFbgcSdgcTlmxlbgc> pageParam=new Page<>(current,limit);
        if (jjgFbgcSdgcTlmxlbgc != null){
            QueryWrapper<JjgFbgcSdgcTlmxlbgc> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcSdgcTlmxlbgc.getProname());
            wrapper.like("htd",jjgFbgcSdgcTlmxlbgc.getHtd());
            wrapper.like("fbgc",jjgFbgcSdgcTlmxlbgc.getFbgc());
            Date jcsj = jjgFbgcSdgcTlmxlbgc.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcSdgcTlmxlbgc> pageModel = jjgFbgcSdgcTlmxlbgcService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除隧道相邻板高差数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcSdgcTlmxlbgcService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }
    }

    @ApiOperation("根据id查询")
    @GetMapping("getTlmxlbgc/{id}")
    public Result getTlmxlbgc(@PathVariable String id) {
        JjgFbgcSdgcTlmxlbgc user = jjgFbgcSdgcTlmxlbgcService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改相邻板高差数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcSdgcTlmxlbgc user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgFbgcSdgcTlmxlbgcService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setFbgc(user.getFbgc());
            sysOperLog.setTitle("隧道相邻板高差数据");
            sysOperLog.setBusinessType("修改");
            sysOperLog.setOperName(JwtHelper.getUsername(request.getHeader("token")));
            sysOperLog.setOperIp(IpUtil.getIpAddress(request));
            sysOperLog.setOperTime(new Date());
            operLogService.saveSysLog(sysOperLog);
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

}

