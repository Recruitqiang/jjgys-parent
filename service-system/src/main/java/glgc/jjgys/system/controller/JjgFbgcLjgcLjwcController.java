package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcLjgcHdgqd;
import glgc.jjgys.model.project.JjgFbgcLjgcLjtsfysdSl;
import glgc.jjgys.model.project.JjgFbgcLjgcLjwc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcLjgcLjwcService;
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
 * @since 2023-02-27
 */
@RestController
@RequestMapping("/jjg/fbgc/ljgc/ljwc")
@CrossOrigin
public class
JjgFbgcLjgcLjwcController {

    @Autowired
    private JjgFbgcLjgcLjwcService jjgFbgcLjgcLjwcService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public Result downloadExport(HttpServletResponse response,String proname,String htd) throws IOException {
        String fileName = "02路基弯沉(贝克曼梁法).xlsx";
        String p = filespath+ File.separator+proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
            return Result.ok();
        }else {
            return Result.fail().message("还未生成鉴定表");
        }
    }



    @ApiOperation("生成路基弯沉鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcLjgcLjwcService.generateJdb(commonInfoVo);

    }
    @ApiOperation("查看路基弯沉鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcLjgcLjwcService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("路基弯沉模板文件导出")
    @GetMapping("exportljwc")
    public void exportljwc(HttpServletResponse response){
        jjgFbgcLjgcLjwcService.exportljwc(response);
    }

    @ApiOperation(value = "路基弯沉数据文件导入")
    @PostMapping("importljwc")
    public Result importljwc(@RequestParam("file") MultipartFile file,CommonInfoVo commonInfoVo) {
        jjgFbgcLjgcLjwcService.importljwc(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcLjgcLjwc jjgFbgcLjgcLjwc){
        //创建page对象
        Page<JjgFbgcLjgcLjwc> pageParam=new Page<>(current,limit);
        if (jjgFbgcLjgcLjwc != null){
            QueryWrapper<JjgFbgcLjgcLjwc> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLjgcLjwc.getProname());
            wrapper.like("htd",jjgFbgcLjgcLjwc.getHtd());
            wrapper.like("fbgc",jjgFbgcLjgcLjwc.getFbgc());
            Date jcsj = jjgFbgcLjgcLjwc.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcLjgcLjwc> pageModel = jjgFbgcLjgcLjwcService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除路基弯沉数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcLjgcLjwcService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }
    }

    @ApiOperation("根据id查询")
    @GetMapping("getLjwc/{id}")
    public Result getLjwc(@PathVariable String id) {
        JjgFbgcLjgcLjwc user = jjgFbgcLjgcLjwcService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改路基弯沉数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcLjgcLjwc user) {
        boolean is_Success = jjgFbgcLjgcLjwcService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }
}

