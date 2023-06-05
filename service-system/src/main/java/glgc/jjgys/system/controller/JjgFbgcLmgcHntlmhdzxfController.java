package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcLmgcGslqlmhdzxf;
import glgc.jjgys.model.project.JjgFbgcLmgcHntlmhdzxf;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcLmgcHntlmhdzxfService;
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
 * @since 2023-04-25
 */
@RestController
@RequestMapping("/jjg/fbgc/lmgc/hntlmhdzxf")
public class JjgFbgcLmgcHntlmhdzxfController {

    @Autowired
    private JjgFbgcLmgcHntlmhdzxfService jjgFbgcLmgcHntlmhdzxfService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname, String htd) throws IOException {
        String fileName = "23混凝土路面厚度-钻芯法.xlsx";
        String p = filespath+ File.separator+proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
        }
    }


    @ApiOperation("生成混凝土路面厚度鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcLmgcHntlmhdzxfService.generateJdb(commonInfoVo);

    }
    @ApiOperation("查看混凝土路面厚度鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcLmgcHntlmhdzxfService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("混凝土路面厚度模板文件导出")
    @GetMapping("exportHntlmhd")
    public void exportHntlmhd(HttpServletResponse response){
        jjgFbgcLmgcHntlmhdzxfService.exportHntlmhd(response);
    }

    @ApiOperation(value = "混凝土路面厚度数据文件导入")
    @PostMapping("importHntlmhd")
    public Result importHntlmhd(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcLmgcHntlmhdzxfService.importHntlmhd(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcLmgcHntlmhdzxf jjgFbgcLmgcHntlmhdzxf){
        //创建page对象
        Page<JjgFbgcLmgcHntlmhdzxf> pageParam=new Page<>(current,limit);
        if (jjgFbgcLmgcHntlmhdzxf != null){
            QueryWrapper<JjgFbgcLmgcHntlmhdzxf> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLmgcHntlmhdzxf.getProname());
            wrapper.like("htd",jjgFbgcLmgcHntlmhdzxf.getHtd());
            wrapper.like("fbgc",jjgFbgcLmgcHntlmhdzxf.getFbgc());
            Date jcsj = jjgFbgcLmgcHntlmhdzxf.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcLmgcHntlmhdzxf> pageModel = jjgFbgcLmgcHntlmhdzxfService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除混凝土路面厚度数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcLmgcHntlmhdzxfService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }
    }

    @ApiOperation("根据id查询")
    @GetMapping("getLmssxs/{id}")
    public Result getLmssxs(@PathVariable String id) {
        JjgFbgcLmgcHntlmhdzxf user = jjgFbgcLmgcHntlmhdzxfService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改混凝土路面厚度数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcLmgcHntlmhdzxf user) {
        boolean is_Success = jjgFbgcLmgcHntlmhdzxfService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

}

