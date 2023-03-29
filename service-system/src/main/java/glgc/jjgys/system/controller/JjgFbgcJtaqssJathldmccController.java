package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabz;
import glgc.jjgys.model.project.JjgFbgcJtaqssJathldmcc;
import glgc.jjgys.model.project.JjgFbgcLjgcHdjgcc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcJtaqssJathldmccService;
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
 * @since 2023-03-01
 */
@RestController
@RequestMapping("/jjg/fbgc/jtaqss/jathldmcc")
@CrossOrigin
public class JjgFbgcJtaqssJathldmccController {

    @Autowired
    private JjgFbgcJtaqssJathldmccService jjgFbgcJtaqssJathldmccService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public Result downloadExport(HttpServletResponse response, String proname, String htd) throws IOException {
        String fileName = "60交安砼护栏断面尺寸.xlsx";
        String p = filespath+ File.separator +proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
            return Result.ok();
        }else {
            return Result.fail().message("还未生成鉴定表");
        }
    }

    @ApiOperation("生成交安砼护栏断面尺寸鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcJtaqssJathldmccService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看交安砼护栏断面尺寸鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcJtaqssJathldmccService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("交安砼护栏断面尺寸模板文件导出")
    @GetMapping("exportjathldmcc")
    public void exportjathldmcc(HttpServletResponse response){
        jjgFbgcJtaqssJathldmccService.exportjathldmcc(response);
    }

    @ApiOperation(value = "交安砼护栏断面尺寸数据文件导入")
    @PostMapping("importjathldmcc")
    public Result importjathldmcc(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcJtaqssJathldmccService.importjathldmcc(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcJtaqssJathldmcc jjgFbgcJtaqssJathldmcc){
        //创建page对象
        Page<JjgFbgcJtaqssJathldmcc> pageParam=new Page<>(current,limit);
        if (jjgFbgcJtaqssJathldmcc != null){
            QueryWrapper<JjgFbgcJtaqssJathldmcc> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcJtaqssJathldmcc.getProname());
            wrapper.like("htd",jjgFbgcJtaqssJathldmcc.getHtd());
            wrapper.like("fbgc",jjgFbgcJtaqssJathldmcc.getFbgc());
            Date jcsj = jjgFbgcJtaqssJathldmcc.getJcrq();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcJtaqssJathldmcc> pageModel = jjgFbgcJtaqssJathldmccService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除交安砼护栏断面尺寸数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean ql = jjgFbgcJtaqssJathldmccService.removeByIds(idList);
        if(ql){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getJathldmcc/{id}")
    public Result getJathldmcc(@PathVariable String id) {
        JjgFbgcJtaqssJathldmcc user = jjgFbgcJtaqssJathldmccService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改交安砼护栏断面尺寸数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcJtaqssJathldmcc user) {
        boolean is_Success = jjgFbgcJtaqssJathldmccService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

}

