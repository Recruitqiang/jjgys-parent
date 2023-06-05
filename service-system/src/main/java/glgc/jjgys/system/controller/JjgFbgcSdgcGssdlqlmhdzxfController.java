package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcSdgcGssdlqlmhdzxf;
import glgc.jjgys.model.project.JjgFbgcSdgcSdlqlmysd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.system.service.JjgFbgcSdgcGssdlqlmhdzxfService;
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
 * @since 2023-05-20
 */
@RestController
@RequestMapping("/jjg/fbgc/sdgc/gssdlqlmhdzxf")
public class JjgFbgcSdgcGssdlqlmhdzxfController {

    @Autowired
    private JjgFbgcSdgcGssdlqlmhdzxfService jjgFbgcSdgcGssdlqlmhdzxfService;

    @Autowired
    private OperLogService operLogService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletRequest request,HttpServletResponse response, String proname, String htd,String fbgc) throws IOException {
        List<Map<String,Object>> sdmclist = jjgFbgcSdgcGssdlqlmhdzxfService.selectsdmc(proname,htd,fbgc);
        List list = new ArrayList<>();
        for (int i=0;i<sdmclist.size();i++){
            list.add(sdmclist.get(i).get("sdmc"));
        }
        String zipName = "53隧道沥青路面厚度-钻芯法";
        JjgFbgcCommonUtils.batchDownloadFile(request,response,zipName,list,filespath+File.separator+proname+File.separator+htd);
    }

    @ApiOperation("生成高速隧道沥青路面厚度钻芯法鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcSdgcGssdlqlmhdzxfService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看高速隧道沥青路面厚度钻芯法鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcSdgcGssdlqlmhdzxfService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("高速隧道沥青路面厚度钻芯法模板文件导出")
    @GetMapping("exportsdzxf")
    public void exportsdzxf(HttpServletResponse response) {
        jjgFbgcSdgcGssdlqlmhdzxfService.exportsdzxf(response);
    }

    @ApiOperation(value = "高速隧道沥青路面厚度钻芯法数据文件导入")
    @PostMapping("importsdzxf")
    public Result importsdzxf(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcSdgcGssdlqlmhdzxfService.importsdzxf(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPageHt/{current}/{limit}")
    public Result findQueryPageHt(@PathVariable long current,
                                  @PathVariable long limit,
                                  @RequestBody JjgFbgcSdgcGssdlqlmhdzxf jjgFbgcSdgcGssdlqlmhdzxf){
        //创建page对象
        Page<JjgFbgcSdgcGssdlqlmhdzxf> pageParam=new Page<>(current,limit);
        if (jjgFbgcSdgcGssdlqlmhdzxf != null){
            QueryWrapper<JjgFbgcSdgcGssdlqlmhdzxf> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcSdgcGssdlqlmhdzxf.getProname());
            wrapper.like("htd",jjgFbgcSdgcGssdlqlmhdzxf.getHtd());
            wrapper.like("fbgc",jjgFbgcSdgcGssdlqlmhdzxf.getFbgc());
            Date sysj = jjgFbgcSdgcGssdlqlmhdzxf.getJcsj();
            if (!StringUtils.isEmpty(sysj)){
                wrapper.like("sysj",sysj);
            }
            //调用方法分页查询
            IPage<JjgFbgcSdgcGssdlqlmhdzxf> pageModel = jjgFbgcSdgcGssdlqlmhdzxfService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");

    }

    @ApiOperation("批量删除隧道沥青路面厚度钻芯法数据")
    @DeleteMapping("removeBeatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcSdgcGssdlqlmhdzxfService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getysd/{id}")
    public Result getysd(@PathVariable String id) {
        JjgFbgcSdgcGssdlqlmhdzxf user = jjgFbgcSdgcGssdlqlmhdzxfService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改隧道沥青路面厚度钻芯法数据")
    @PostMapping("updateysd")
    public Result updateysd(@RequestBody JjgFbgcSdgcGssdlqlmhdzxf user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgFbgcSdgcGssdlqlmhdzxfService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setFbgc(user.getFbgc());
            sysOperLog.setTitle("隧道沥青路面厚度钻芯法数据");
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

