package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabx;
import glgc.jjgys.model.project.JjgHtd;
import glgc.jjgys.model.project.JjgLqsQl;
import glgc.jjgys.model.project.JjgPlaninfo;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.system.service.JjgFbgcGenerateWordService;
import glgc.jjgys.system.service.JjgHtdService;
import glgc.jjgys.system.service.JjgLookProjectPlanService;
import glgc.jjgys.system.service.OperLogService;
import io.swagger.annotations.ApiOperation;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
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
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;


/**
 * 先调用gethtd方法，CommonInfoVo对象中只传项目名称就可以对象中只传项目名称就可以，然后，后端返回合同段的信息
 * 其次调用lookplan方法，CommonInfoVo对象中传递项目名称和合同段名称，后端返回数据
 *
 * @author wq
 * @since 2023-03-01
 */
@RestController
@RequestMapping("/jjg/project/plan")
@CrossOrigin
public class JjgLookProjectPlanController {

    @Autowired
    private JjgLookProjectPlanService jjgLookProjectPlanService;

    @Autowired
    private JjgHtdService jjgHtdService;

    @Autowired
    private OperLogService operLogService;


    @ApiOperation("合同段信息")
    @PostMapping("gethtd")
    public Result gethtd(@RequestBody CommonInfoVo commonInfoVo){
        String proname = commonInfoVo.getProname();
        List<JjgHtd> gethtd = jjgHtdService.gethtd(proname);
        List<String> htd = new ArrayList<>();
        for (JjgHtd jjgHtd : gethtd) {
            String name = jjgHtd.getName();
            htd.add(name);
        }
        return Result.ok(htd);
    }


    @ApiOperation("查看项目进度")
    @PostMapping("lookplan")
    public Result lookplan(@RequestBody CommonInfoVo commonInfoVo){
        /**
         * 项目，合同段
         */
        List<Map<String,Object>> result = jjgLookProjectPlanService.lookplan(commonInfoVo);
        return Result.ok(result);
    }

    //桥梁清单文件导出
    @ApiOperation("项目进度清单文件导出")
    @GetMapping("exportxmjd/{projectname}")
    public void exportxmjd(HttpServletResponse response, @PathVariable String projectname){
        jjgLookProjectPlanService.exportxmjd(response,projectname);
    }


    @ApiOperation(value = "项目进度信息导入")
    @PostMapping("importxmjd")
    public Result importxmjd(@RequestParam("file") MultipartFile file, @RequestParam String proname) {
        jjgLookProjectPlanService.importxmjd(file,proname);
        return Result.ok();
    }


    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgPlaninfo jjgPlaninfo){
        //创建page对象
        Page<JjgPlaninfo> pageParam=new Page<>(current,limit);
        //判断projectQueryVo对象是否为空，直接查全部
        if(jjgPlaninfo == null){
            IPage<JjgPlaninfo> pageModel = jjgLookProjectPlanService.page(pageParam,null);
            return Result.ok(pageModel);
        }else {
            //获取条件值，进行非空判断，条件封装
            String name = jjgPlaninfo.getProname();
            String htd = jjgPlaninfo.getHtd();
            QueryWrapper<JjgPlaninfo> wrapper=new QueryWrapper<>();
            wrapper.like("proname",name);
            wrapper.like("htd",htd);
            wrapper.orderByDesc("create_time");
            //调用方法分页查询
            IPage<JjgPlaninfo> pageModel = jjgLookProjectPlanService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);

        }
    }

    @ApiOperation("批量删除项目进度信息")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean ql = jjgLookProjectPlanService.removeByIds(idList);
        if(ql){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }
    }

    @ApiOperation("根据id查询")
    @GetMapping("getxmjd/{id}")
    public Result getxmjd(@PathVariable String id) {
        JjgPlaninfo user = jjgLookProjectPlanService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改项目进度信息")
    @PostMapping("update")
    public Result update(@RequestBody JjgPlaninfo user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgLookProjectPlanService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setFbgc(user.getFbgc());
            sysOperLog.setTitle("项目进度信息");
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

