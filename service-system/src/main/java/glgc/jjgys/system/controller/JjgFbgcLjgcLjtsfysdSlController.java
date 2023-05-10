package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcLjgcLjtsfysdHt;
import glgc.jjgys.model.project.JjgFbgcLjgcLjtsfysdSl;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.system.service.JjgFbgcLjgcLjtsfysdSlService;
import glgc.jjgys.system.service.OperLogService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.context.request.RequestAttributes;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.text.ParseException;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * <p>
 *  路基土石方压实度_砂砾
 * </p>
 *
 * @author wq
 * @since 2023-02-21
 */
@RestController
@RequestMapping("/jjg/fbgc/ljgc/ljtsfysd")
@CrossOrigin
public class JjgFbgcLjgcLjtsfysdSlController {

    @Autowired
    private JjgFbgcLjgcLjtsfysdSlService jjgFbgcLjgcLjtsfysdSlService;

    @Autowired
    private OperLogService operLogService;


    @ApiOperation("路基土石方压实度_砂砾模板文件导出")
    @GetMapping("exportysdsl")
    public void exportysdsl(HttpServletResponse response) throws IOException {
        jjgFbgcLjgcLjtsfysdSlService.exportysdsl(response);
    }

    @ApiOperation(value = "路基土石方压实度_砂砾数据文件导入")
    @PostMapping("importysdsl")
    public Result importysdsl(@RequestParam("file") MultipartFile file,CommonInfoVo commonInfoVo) throws IOException, ParseException {
        jjgFbgcLjgcLjtsfysdSlService.importysdsl(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPageSL/{current}/{limit}")
    public Result findQueryPageSL(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcLjgcLjtsfysdSl jjgFbgcLjgcLjtsfysdSl){
        //创建page对象
        Page<JjgFbgcLjgcLjtsfysdSl> pageParam=new Page<>(current,limit);
        if (jjgFbgcLjgcLjtsfysdSl != null){
            QueryWrapper<JjgFbgcLjgcLjtsfysdSl> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLjgcLjtsfysdSl.getProname());
            wrapper.like("htd",jjgFbgcLjgcLjtsfysdSl.getHtd());
            wrapper.like("fbgc",jjgFbgcLjgcLjtsfysdSl.getFbgc());
            Date sysj = jjgFbgcLjgcLjtsfysdSl.getSysj();
            if (!StringUtils.isEmpty(sysj)){
                wrapper.like("sysj",sysj);
            }
            //调用方法分页查询
            IPage<JjgFbgcLjgcLjtsfysdSl> pageModel = jjgFbgcLjgcLjtsfysdSlService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");

    }

    @ApiOperation("批量删除路基土石方压实度_砂砾数据")
    @DeleteMapping("removeBeatchSL")
    public Result removeBeatchSL(@RequestBody List<String> idList){
        boolean hd = jjgFbgcLjgcLjtsfysdSlService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getSl/{id}")
    public Result getSl(@PathVariable String id) {
        JjgFbgcLjgcLjtsfysdSl user = jjgFbgcLjgcLjtsfysdSlService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改路基土石方压实度_砂砾数据")
    @PostMapping("updateSl")
    public Result updateSl(@RequestBody JjgFbgcLjgcLjtsfysdSl user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();

        boolean is_Success = jjgFbgcLjgcLjtsfysdSlService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setFbgc(user.getFbgc());
            sysOperLog.setTitle("路基土石方压实度_砂砾数据");
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

