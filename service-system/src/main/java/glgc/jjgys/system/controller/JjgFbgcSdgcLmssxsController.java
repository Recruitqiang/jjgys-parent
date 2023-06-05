package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcLmgcLmssxs;
import glgc.jjgys.model.project.JjgFbgcSdgcLmssxs;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.system.service.JjgFbgcSdgcLmssxsService;
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
@RequestMapping("/system/jjg-fbgc-sdgc-lmssxs")
public class JjgFbgcSdgcLmssxsController {

    @Autowired
    private JjgFbgcSdgcLmssxsService jjgFbgcSdgcLmssxsService;

    @Autowired
    private OperLogService operLogService;


    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletRequest request,HttpServletResponse response, String proname, String htd,String fbgc) throws IOException {
        List<Map<String,Object>> sdmclist = jjgFbgcSdgcLmssxsService.selectsdmc(proname,htd,fbgc);
        List list = new ArrayList<>();
        for (int i=0;i<sdmclist.size();i++){
            list.add(sdmclist.get(i).get("sdmc"));
        }
        String zipName = "46隧道沥青路面渗水系数";
        JjgFbgcCommonUtils.batchDownloadFile(request,response,zipName,list,filespath+File.separator+proname+File.separator+htd);
    }


    @ApiOperation("生成隧道沥青路面渗水系数鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcSdgcLmssxsService.generateJdb(commonInfoVo);

    }
    @ApiOperation("查看隧道沥青路面渗水系数鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcSdgcLmssxsService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("隧道沥青路面渗水系数模板文件导出")
    @GetMapping("exportSdlmssxs")
    public void exportSdlmssxs(HttpServletResponse response){
        jjgFbgcSdgcLmssxsService.exportSdlmssxs(response);
    }

    @ApiOperation(value = "隧道沥青路面渗水系数数据文件导入")
    @PostMapping("importSdlmssxs")
    public Result importSdlmssxs(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcSdgcLmssxsService.importSdlmssxs(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcSdgcLmssxs jjgFbgcSdgcLmssxs){
        //创建page对象
        Page<JjgFbgcSdgcLmssxs> pageParam=new Page<>(current,limit);
        if (jjgFbgcSdgcLmssxs != null){
            QueryWrapper<JjgFbgcSdgcLmssxs> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcSdgcLmssxs.getProname());
            wrapper.like("htd",jjgFbgcSdgcLmssxs.getHtd());
            wrapper.like("fbgc",jjgFbgcSdgcLmssxs.getFbgc());
            Date jcsj = jjgFbgcSdgcLmssxs.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcSdgcLmssxs> pageModel = jjgFbgcSdgcLmssxsService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除隧道沥青路面渗水系数数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcSdgcLmssxsService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }
    }

    @ApiOperation("根据id查询")
    @GetMapping("getLmssxs/{id}")
    public Result getLmssxs(@PathVariable String id) {
        JjgFbgcSdgcLmssxs user = jjgFbgcSdgcLmssxsService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改隧道沥青路面渗水系数")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcSdgcLmssxs user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgFbgcSdgcLmssxsService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setFbgc(user.getFbgc());
            sysOperLog.setTitle("隧道沥青路面渗水系数");
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

