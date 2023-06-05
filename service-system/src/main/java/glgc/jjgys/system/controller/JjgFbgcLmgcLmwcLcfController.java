package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcLjgcLjwcLcf;
import glgc.jjgys.model.project.JjgFbgcLmgcLmwcLcf;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.system.service.JjgFbgcLmgcLmwcLcfService;
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
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * <p>
 *  前端控制器
 * </p>
 *
 * @author wq
 * @since 2023-04-18
 */
@RestController
@RequestMapping("/jjg/fbgc/lmgc/lmwclcf")
public class JjgFbgcLmgcLmwcLcfController {

    @Autowired
    private JjgFbgcLmgcLmwcLcfService jjgFbgcLmgcLmwcLcfService;

    @Autowired
    private OperLogService operLogService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @ApiOperation("查看路面弯沉落锤法鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcLmgcLmwcLcfService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname, String htd) throws IOException {
        String fileName = "13路面弯沉(落锤法).xlsx";
        String p = filespath+ File.separator+proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
        }
    }

    @ApiOperation("生成路面弯沉落锤法鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcLmgcLmwcLcfService.generateJdb(commonInfoVo);

    }

    @ApiOperation("路面弯沉落锤法模板文件导出")
    @GetMapping("exportlmwclcf")
    public void exportlmwclcf(HttpServletResponse response){
        jjgFbgcLmgcLmwcLcfService.exportlmwclcf(response);
    }

    @ApiOperation(value = "路面弯沉落锤法数据文件导入")
    @PostMapping("importlmwclcf")
    public Result importlmwclcf(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcLmgcLmwcLcfService.importlmwclcf(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcLmgcLmwcLcf jjgFbgcLmgcLmwcLcf){
        //创建page对象
        Page<JjgFbgcLmgcLmwcLcf> pageParam=new Page<>(current,limit);
        if (jjgFbgcLmgcLmwcLcf != null){
            QueryWrapper<JjgFbgcLmgcLmwcLcf> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLmgcLmwcLcf.getProname());
            wrapper.like("htd",jjgFbgcLmgcLmwcLcf.getHtd());
            wrapper.like("fbgc",jjgFbgcLmgcLmwcLcf.getFbgc());
            Date jcsj = jjgFbgcLmgcLmwcLcf.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcLmgcLmwcLcf> pageModel = jjgFbgcLmgcLmwcLcfService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除路面弯沉落锤法数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcLmgcLmwcLcfService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }
    }

    @ApiOperation("根据id查询")
    @GetMapping("getLmwcLcf/{id}")
    public Result getLmwcLcf(@PathVariable String id) {
        JjgFbgcLmgcLmwcLcf user = jjgFbgcLmgcLmwcLcfService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改路面弯沉落锤法数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcLmgcLmwcLcf user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgFbgcLmgcLmwcLcfService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setFbgc(user.getFbgc());
            sysOperLog.setTitle("路面弯沉落锤法数据");
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

