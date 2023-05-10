package glgc.jjgys.system.controller;


import cn.hutool.core.io.resource.InputStreamResource;
import com.alibaba.excel.util.FileUtils;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcLjgcHdgqd;
import glgc.jjgys.model.project.JjgLqsQl;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.annotation.Log;
import glgc.jjgys.system.enums.BusinessType;
import glgc.jjgys.system.service.JjgFbgcLjgcHdgqdService;
import glgc.jjgys.system.service.LoginLogService;
import glgc.jjgys.system.service.OperLogService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.FileSystemResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.context.request.RequestAttributes;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URISyntaxException;
import java.net.URLEncoder;
import java.nio.channels.Channels;
import java.nio.channels.FileChannel;
import java.nio.channels.WritableByteChannel;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
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
 * @since 2023-01-14
 */
@Api(tags = "涵洞砼强度")
@RestController
@RequestMapping("/jjg/fbgc/ljgc/hdgqd")
@CrossOrigin
public class JjgFbgcLjgcHdgqdController {
    @Autowired
    private JjgFbgcLjgcHdgqdService jjgFbgcLjgcHdgqdService;

    @Autowired
    private OperLogService operLogService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;


    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response,String proname,String htd) throws IOException {
        String fileName = "08路基涵洞砼强度.xlsx";
        String p = filespath+File.separator+proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
        }
    }

    @ApiOperation("生成涵洞砼强度鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcLjgcHdgqdService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看涵洞砼强度鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcLjgcHdgqdService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("涵洞砼强度模板文件导出")
    @GetMapping("exporthdgqd")
    public void exporthdgqd(HttpServletResponse response){
        jjgFbgcLjgcHdgqdService.exporthdgqd(response);
    }


    /**
     * 导入时需要传入项目名称，合同段，分布工程
     * @param file
     * @return
     */
    @ApiOperation(value = "涵洞砼强度数据文件导入")
    @PostMapping("importhdgqd")
    public Result importhdgqd(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcLjgcHdgqdService.importhdgqd(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcLjgcHdgqd jjgFbgcLjgcHdgqd){
        //创建page对象
        Page<JjgFbgcLjgcHdgqd> pageParam=new Page<>(current,limit);
        if (jjgFbgcLjgcHdgqd != null){
            QueryWrapper<JjgFbgcLjgcHdgqd> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLjgcHdgqd.getProname());
            wrapper.like("htd",jjgFbgcLjgcHdgqd.getHtd());
            wrapper.like("fbgc",jjgFbgcLjgcHdgqd.getFbgc());
            Date jcsj = jjgFbgcLjgcHdgqd.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcLjgcHdgqd> pageModel = jjgFbgcLjgcHdgqdService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
        //判断projectQueryVo对象是否为空，直接查全部
        /*if(jjgFbgcLjgcHdgqd == null){
            IPage<JjgFbgcLjgcHdgqd> pageModel = jjgFbgcLjgcHdgqdService.page(pageParam,null);
            return Result.ok(pageModel);
        }else {
            //获取条件值，进行非空判断，条件封装
            Date jcsj = jjgFbgcLjgcHdgqd.getJcsj();
            QueryWrapper<JjgFbgcLjgcHdgqd> wrapper=new QueryWrapper<>();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            wrapper.orderByAsc("zh","bw1");
//            wrapper.like("fbgc",jjgFbgcLjgcHdgqd.getFbgc());
            //调用方法分页查询
            IPage<JjgFbgcLjgcHdgqd> pageModel = jjgFbgcLjgcHdgqdService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);

        }*/
    }

    @ApiOperation("根据id查询")
    @GetMapping("getHdgqd/{id}")
    public Result getHdgqd(@PathVariable String id) {
        JjgFbgcLjgcHdgqd user = jjgFbgcLjgcHdgqdService.getById(id);
        return Result.ok(user);
    }

    //@Log(title = "涵洞砼强度数据",businessType = BusinessType.UPDATE)
    @ApiOperation("修改涵洞砼强度数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcLjgcHdgqd user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgFbgcLjgcHdgqdService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setFbgc(user.getFbgc());
            sysOperLog.setTitle("涵洞砼强度数据");
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

    @ApiOperation("批量删除涵洞砼强度数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcLjgcHdgqdService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

}

