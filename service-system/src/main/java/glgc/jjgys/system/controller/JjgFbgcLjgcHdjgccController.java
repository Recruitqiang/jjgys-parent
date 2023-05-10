package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcLjgcHdgqd;
import glgc.jjgys.model.project.JjgFbgcLjgcHdjgcc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.system.annotation.Log;
import glgc.jjgys.system.enums.BusinessType;
import glgc.jjgys.system.service.JjgFbgcLjgcHdjgccService;
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
 * @since 2023-02-08
 */
@RestController
@RequestMapping("/jjg/fbgc/ljgc/hdjgcc")
@CrossOrigin
public class JjgFbgcLjgcHdjgccController {

    @Autowired
    private JjgFbgcLjgcHdjgccService jjgFbgcLjgcHdjgccService;

    @Autowired
    private OperLogService operLogService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;


    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response,String proname,String htd) throws IOException {
        String fileName = "09路基涵洞结构尺寸.xlsx";
        String p = filespath+ File.separator+proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
        }
    }

    @ApiOperation("生成涵洞结构尺寸鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcLjgcHdjgccService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看涵洞结构尺寸鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcLjgcHdjgccService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }


    @ApiOperation("涵洞结构尺寸模板文件导出")
    @GetMapping("exporthdjgcc")
    public void exporthdjgcc(HttpServletResponse response){
        jjgFbgcLjgcHdjgccService.exporthdjgcc(response);
    }

    /**
     * 遗留问题：导入时需要传入项目名称，合同段，分布工程
     * @param file
     * @return
     */
    @ApiOperation(value = "涵洞结构尺寸数据文件导入")
    @PostMapping("importhdjgcc")
    public Result importhdjgcc(@RequestParam("file") MultipartFile file,CommonInfoVo commonInfoVo) {
        jjgFbgcLjgcHdjgccService.importhdjgcc(file,commonInfoVo);
        return Result.ok();
    }

    //遗留问题：还需要根据项目名和合同段明查询
    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcLjgcHdjgcc jjgFbgcLjgcHdjgcc){
        //创建page对象
        Page<JjgFbgcLjgcHdjgcc> pageParam=new Page<>(current,limit);
        if (jjgFbgcLjgcHdjgcc != null){
            QueryWrapper<JjgFbgcLjgcHdjgcc> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLjgcHdjgcc.getProname());
            wrapper.like("htd",jjgFbgcLjgcHdjgcc.getHtd());
            wrapper.like("fbgc",jjgFbgcLjgcHdjgcc.getFbgc());
            Date jcsj = jjgFbgcLjgcHdjgcc.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcLjgcHdjgcc> pageModel = jjgFbgcLjgcHdjgccService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除涵洞结构尺寸数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean ql = jjgFbgcLjgcHdjgccService.removeByIds(idList);
        if(ql){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getHdjgcc/{id}")
    public Result getHdjgcc(@PathVariable String id) {
        JjgFbgcLjgcHdjgcc user = jjgFbgcLjgcHdjgccService.getById(id);
        return Result.ok(user);
    }

    //@Log(title = "涵洞结构尺寸数据",businessType = BusinessType.UPDATE)
    @ApiOperation("修改涵洞结构尺寸数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcLjgcHdjgcc user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgFbgcLjgcHdjgccService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setFbgc(user.getFbgc());
            sysOperLog.setTitle("涵洞结构尺寸数据");
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

