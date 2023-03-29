package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcLjgcHdgqd;
import glgc.jjgys.model.project.JjgFbgcLjgcHdjgcc;
import glgc.jjgys.model.project.JjgFbgcLjgcXqjgcc;
import glgc.jjgys.model.project.JjgFbgcLjgcZddmcc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcLjgcZddmccService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.autoconfigure.AutoConfigureOrder;
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
 * @since 2023-02-16
 */
@RestController
@RequestMapping("/jjg/fbgc/ljgc/zddmcc")
@CrossOrigin
public class JjgFbgcLjgcZddmccController {

    @Autowired
    private JjgFbgcLjgcZddmccService jjgFbgcLjgcZddmccService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;


    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public Result downloadExport(HttpServletResponse response,String proname,String htd) throws IOException {
        String fileName = "11路基支挡断面尺寸.xlsx";
        String p = filespath+ File.separator+proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
            return Result.ok();
        }else {
            return Result.fail().message("还未生成鉴定表");
        }
    }

    @ApiOperation("生成支档断面尺寸鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcLjgcZddmccService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看支档断面尺寸鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcLjgcZddmccService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("支档断面尺寸模板文件导出")
    @GetMapping("exportzddmcc")
    public void exportzddmcc(HttpServletResponse response){
        jjgFbgcLjgcZddmccService.exportzddmcc(response);
    }


    /**
     * 遗留问题：导入时需要传入项目名称，合同段，分布工程
     * @param file
     * @return
     */
    @ApiOperation(value = "支档断面尺寸数据文件导入")
    @PostMapping("importzddmcc")
    public Result importzddmcc(@RequestParam("file") MultipartFile file,CommonInfoVo commonInfoVo) {
        jjgFbgcLjgcZddmccService.importzddmcc(file,commonInfoVo);
        return Result.ok();
    }

    //遗留问题：还需要根据项目名和合同段明查询
    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcLjgcZddmcc jjgFbgcLjgcZddmcc){

        //创建page对象
        Page<JjgFbgcLjgcZddmcc> pageParam=new Page<>(current,limit);
        if (jjgFbgcLjgcZddmcc != null){
            QueryWrapper<JjgFbgcLjgcZddmcc> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLjgcZddmcc.getProname());
            wrapper.like("htd",jjgFbgcLjgcZddmcc.getHtd());
            wrapper.like("fbgc",jjgFbgcLjgcZddmcc.getFbgc());
            Date jcsj = jjgFbgcLjgcZddmcc.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcLjgcZddmcc> pageModel = jjgFbgcLjgcZddmccService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");

        }

    @ApiOperation("批量删除支档断面尺寸数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcLjgcZddmccService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getZddmcc/{id}")
    public Result getZddmcc(@PathVariable String id) {
        JjgFbgcLjgcZddmcc user = jjgFbgcLjgcZddmccService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改支档断面尺寸数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcLjgcZddmcc user) {
        boolean is_Success = jjgFbgcLjgcZddmccService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

}

