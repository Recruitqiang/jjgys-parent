package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcLjgcLjwc;
import glgc.jjgys.model.project.JjgFbgcLjgcLjwcLcf;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcLjgcLjwcLcfService;
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

/**
 * <p>
 *  路基弯沉落锤法
 * </p>
 *
 * @author wq
 * @since 2023-03-01
 */
@RestController
@RequestMapping("/jjg/fbgc/ljgc/ljwclcf")
@CrossOrigin
public class JjgFbgcLjgcLjwcLcfController {

    @Autowired
    private JjgFbgcLjgcLjwcLcfService jjgFbgcLjgcLjwcLcfService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public Result downloadExport(HttpServletResponse response,String proname,String htd) throws IOException {
        String fileName = "02路基弯沉(落锤法).xlsx";
        String p = filespath+ File.separator+proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
            return Result.ok();
        }else {
            return Result.fail().message("还未生成鉴定表");
        }
    }

    @ApiOperation("生成路基弯沉落锤法鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcLjgcLjwcLcfService.generateJdb(commonInfoVo);

    }

    @ApiOperation("路基弯沉落锤法模板文件导出")
    @GetMapping("exportljwclcf")
    public void exportljwclcf(HttpServletResponse response){
        jjgFbgcLjgcLjwcLcfService.exportljwclcf(response);
    }

    @ApiOperation(value = "路基弯沉落锤法数据文件导入")
    @PostMapping("importljwclcf")
    public Result importljwclcf(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcLjgcLjwcLcfService.importljwclcf(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcLjgcLjwcLcf jjgFbgcLjgcLjwcLcf){
        //创建page对象
        Page<JjgFbgcLjgcLjwcLcf> pageParam=new Page<>(current,limit);
        if (jjgFbgcLjgcLjwcLcf != null){
            QueryWrapper<JjgFbgcLjgcLjwcLcf> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLjgcLjwcLcf.getProname());
            wrapper.like("htd",jjgFbgcLjgcLjwcLcf.getHtd());
            wrapper.like("fbgc",jjgFbgcLjgcLjwcLcf.getFbgc());
            Date jcsj = jjgFbgcLjgcLjwcLcf.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcLjgcLjwcLcf> pageModel = jjgFbgcLjgcLjwcLcfService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除路基弯沉落锤法数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcLjgcLjwcLcfService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }
    }

    @ApiOperation("根据id查询")
    @GetMapping("getLjwcLcf/{id}")
    public Result getLjwcLcf(@PathVariable String id) {
        JjgFbgcLjgcLjwcLcf user = jjgFbgcLjgcLjwcLcfService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改路基弯沉落锤法数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcLjgcLjwcLcf user) {
        boolean is_Success = jjgFbgcLjgcLjwcLcfService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }


}

