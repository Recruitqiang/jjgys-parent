package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcLjgcZddmcc;
import glgc.jjgys.model.project.JjgFbgcQlgcQmhp;
import glgc.jjgys.model.project.JjgFbgcQlgcSbJgcc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.system.service.JjgFbgcQlgcSbJgccService;
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
 * @since 2023-03-20
 */
@RestController
@RequestMapping("/jjg/fbgc/qlgc/sb/jgcc")
@CrossOrigin
public class JjgFbgcQlgcSbJgccController {

    @Autowired
    private JjgFbgcQlgcSbJgccService jjgFbgcQlgcSbJgccService;

    @Autowired
    private OperLogService operLogService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname, String htd) throws IOException {
        String fileName = "30桥梁上部主要结构尺寸.xlsx";
        String p = filespath+ File.separator +proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
        }
    }

    @ApiOperation("生成桥梁上部结构尺寸鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcQlgcSbJgccService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看桥梁上部结构尺寸鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcQlgcSbJgccService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("桥梁上部结构尺寸模板文件导出")
    @GetMapping("exportqlsbjgcc")
    public void exportqlsbjgcc(HttpServletResponse response){
        jjgFbgcQlgcSbJgccService.exportqlsbjgcc(response);
    }


    /**
     * @param file
     * @return
     */
    @ApiOperation(value = "桥梁上部结构尺寸数据文件导入")
    @PostMapping("importqlsbjgcc")
    public Result importqlsbjgcc(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcQlgcSbJgccService.importqlsbjgcc(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcQlgcSbJgcc jjgFbgcQlgcSbJgcc){

        //创建page对象
        Page<JjgFbgcQlgcSbJgcc> pageParam=new Page<>(current,limit);
        if (jjgFbgcQlgcSbJgcc != null){
            QueryWrapper<JjgFbgcQlgcSbJgcc> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcQlgcSbJgcc.getProname());
            wrapper.like("htd",jjgFbgcQlgcSbJgcc.getHtd());
            wrapper.like("fbgc",jjgFbgcQlgcSbJgcc.getFbgc());
            Date jcsj = jjgFbgcQlgcSbJgcc.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcQlgcSbJgcc> pageModel = jjgFbgcQlgcSbJgccService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");

    }

    @ApiOperation("批量删除桥梁上部结构尺寸数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean ql = jjgFbgcQlgcSbJgccService.removeByIds(idList);
        if(ql){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getQmhp{id}")
    public Result getQmhp(@PathVariable String id) {
        JjgFbgcQlgcSbJgcc user = jjgFbgcQlgcSbJgccService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改桥梁上部结构尺寸数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcQlgcSbJgcc user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgFbgcQlgcSbJgccService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setFbgc(user.getFbgc());
            sysOperLog.setTitle("桥梁上部结构尺寸数据");
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

