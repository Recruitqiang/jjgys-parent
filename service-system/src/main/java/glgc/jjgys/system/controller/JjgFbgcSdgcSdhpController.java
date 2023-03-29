package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcQlgcQmhp;
import glgc.jjgys.model.project.JjgFbgcSdgcDmpzd;
import glgc.jjgys.model.project.JjgFbgcSdgcSdhp;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcSdgcSdhpService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

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
 * @since 2023-03-27
 */
@RestController
@RequestMapping("/jjg/fbgc/sdgc/sdhp")
public class JjgFbgcSdgcSdhpController {

    @Autowired
    private JjgFbgcSdgcSdhpService jjgFbgcSdgcSdhpService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public Result downloadExport(HttpServletResponse response, String proname, String htd) throws IOException {
        //暂定
        String fileName = ".xlsx";
        String p = filespath+ File.separator +proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
            return Result.ok();
        }else {
            return Result.fail().message("还未生成鉴定表");
        }
    }

    @ApiOperation("生成隧道横坡鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcSdgcSdhpService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看隧道横坡鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcSdgcSdhpService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("隧道横坡模板文件导出")
    @GetMapping("export")
    public void export(HttpServletResponse response){
        jjgFbgcSdgcSdhpService.export(response);
    }


    /**
     * @param file
     * @return
     */
    @ApiOperation(value = "隧道横坡数据文件导入")
    @PostMapping("importsdhp")
    public Result importsdhp(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcSdgcSdhpService.importsdhp(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcSdgcSdhp jjgFbgcSdgcSdhp){
        //创建page对象
        Page<JjgFbgcSdgcSdhp> pageParam=new Page<>(current,limit);
        if (jjgFbgcSdgcSdhp != null){
            QueryWrapper<JjgFbgcSdgcSdhp> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcSdgcSdhp.getProname());
            wrapper.like("htd",jjgFbgcSdgcSdhp.getHtd());
            wrapper.like("fbgc",jjgFbgcSdgcSdhp.getFbgc());
            Date jcsj = jjgFbgcSdgcSdhp.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcSdgcSdhp> pageModel = jjgFbgcSdgcSdhpService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");

    }

    @ApiOperation("批量删除隧道横坡数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean ql = jjgFbgcSdgcSdhpService.removeByIds(idList);
        if(ql){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getSdhp{id}")
    public Result getSdhp(@PathVariable String id) {
        JjgFbgcSdgcSdhp user = jjgFbgcSdgcSdhpService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改隧道横坡数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcSdgcSdhp user) {
        boolean is_Success = jjgFbgcSdgcSdhpService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }



}

