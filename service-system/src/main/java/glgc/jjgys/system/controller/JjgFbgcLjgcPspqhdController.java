package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcLjgcHdgqd;
import glgc.jjgys.model.project.JjgFbgcLjgcPsdmcc;
import glgc.jjgys.model.project.JjgFbgcLjgcPspqhd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.system.service.JjgFbgcLjgcPspqhdService;
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
 * @since 2023-02-16
 */
@RestController
@RequestMapping("/jjg/fbgc/ljgc/pspqhd")
@CrossOrigin
public class JjgFbgcLjgcPspqhdController {

    @Autowired
    private JjgFbgcLjgcPspqhdService jjgFbgcLjgcPspqhdService;

    @Autowired
    private OperLogService operLogService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;


    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response,String proname,String htd) throws IOException {
        String fileName = "05路基排水铺砌厚度.xlsx";
        String p = filespath+ File.separator+proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
        }
    }

    @ApiOperation("生成排水铺砌厚度鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcLjgcPspqhdService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看排水铺砌厚度鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcLjgcPspqhdService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("排水铺砌厚度模板文件导出")
    @GetMapping("exportpspqhd")
    public void exportpspqhd(HttpServletResponse response){
        jjgFbgcLjgcPspqhdService.exportpspqhd(response);
    }


    /**
     * 遗留问题：导入时需要传入项目名称，合同段，分布工程
     * @param file
     * @return
     */
    @ApiOperation(value = "排水铺砌厚度数据文件导入")
    @PostMapping("importpspqhd")
    public Result importpspqhd(@RequestParam("file") MultipartFile file,CommonInfoVo commonInfoVo) {
        jjgFbgcLjgcPspqhdService.importpspqhd(file,commonInfoVo);
        return Result.ok();
    }

    //遗留问题：还需要根据项目名和合同段明查询
    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcLjgcPspqhd jjgFbgcLjgcPspqhd){
        //创建page对象
        Page<JjgFbgcLjgcPspqhd> pageParam=new Page<>(current,limit);
        if (jjgFbgcLjgcPspqhd != null){
            QueryWrapper<JjgFbgcLjgcPspqhd> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLjgcPspqhd.getProname());
            wrapper.like("htd",jjgFbgcLjgcPspqhd.getHtd());
            wrapper.like("fbgc",jjgFbgcLjgcPspqhd.getFbgc());
            Date jcsj = jjgFbgcLjgcPspqhd.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcLjgcPspqhd> pageModel = jjgFbgcLjgcPspqhdService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除排水铺砌厚度数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean ql = jjgFbgcLjgcPspqhdService.removeByIds(idList);
        if(ql){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getPspqhd/{id}")
    public Result getPspqhd(@PathVariable String id) {
        JjgFbgcLjgcPspqhd user = jjgFbgcLjgcPspqhdService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改排水铺砌厚度数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcLjgcPspqhd user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgFbgcLjgcPspqhdService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setFbgc(user.getFbgc());
            sysOperLog.setTitle("排水铺砌厚度数据");
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

