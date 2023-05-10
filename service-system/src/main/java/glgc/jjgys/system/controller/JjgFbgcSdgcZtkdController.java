package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcQlgcXbTqd;
import glgc.jjgys.model.project.JjgFbgcSdgcSdhp;
import glgc.jjgys.model.project.JjgFbgcSdgcZtkd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.system.service.JjgFbgcSdgcZtkdService;
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
 *  隧道总体宽度
 * </p>
 *
 * @author wq
 * @since 2023-03-26
 */
@RestController
@RequestMapping("/jjg/fbgc/sdgc/ztkd")
@CrossOrigin
public class JjgFbgcSdgcZtkdController {

    @Autowired
    private JjgFbgcSdgcZtkdService jjgFbgcSdgcZtkdService;

    @Autowired
    private OperLogService operLogService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname, String htd) throws IOException {
        String fileName = "41隧道总体宽度.xlsx";
        String p = filespath+"/"+proname+"/"+htd+"/"+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
        }
    }

    @ApiOperation("生成隧道总体宽度鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcSdgcZtkdService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看隧道总体宽度鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcSdgcZtkdService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("隧道总体宽度模板文件导出")
    @GetMapping("exportsdztkd")
    public void exportsdztkd(HttpServletResponse response){
        jjgFbgcSdgcZtkdService.exportsdztkd(response);
    }

    /**
     * @param file
     * @return
     */
    @ApiOperation(value = "隧道总体宽度数据文件导入")
    @PostMapping("importsdztkd")
    public Result importsdztkd(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcSdgcZtkdService.importsdztkd(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcSdgcZtkd jjgFbgcSdgcZtkd){
        //创建page对象
        Page<JjgFbgcSdgcZtkd> pageParam=new Page<>(current,limit);
        if (jjgFbgcSdgcZtkd != null){
            QueryWrapper<JjgFbgcSdgcZtkd> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcSdgcZtkd.getProname());
            wrapper.like("htd",jjgFbgcSdgcZtkd.getHtd());
            wrapper.like("fbgc",jjgFbgcSdgcZtkd.getFbgc());
            Date jcsj = jjgFbgcSdgcZtkd.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcSdgcZtkd> pageModel = jjgFbgcSdgcZtkdService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除隧道总体宽度数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcSdgcZtkdService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getSdhp{id}")
    public Result getSdhp(@PathVariable String id) {
        JjgFbgcSdgcZtkd user = jjgFbgcSdgcZtkdService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改隧道总体宽度数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcSdgcZtkd user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgFbgcSdgcZtkdService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setFbgc(user.getFbgc());
            sysOperLog.setTitle("隧道总体宽度数据");
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

