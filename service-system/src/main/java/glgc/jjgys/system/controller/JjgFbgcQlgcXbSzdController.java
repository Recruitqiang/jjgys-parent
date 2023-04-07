package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcQlgcXbJgcc;
import glgc.jjgys.model.project.JjgFbgcQlgcXbSzd;
import glgc.jjgys.model.project.JjgFbgcQlgcXbTqd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcQlgcXbSzdService;
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
 * @since 2023-03-22
 */
@RestController
@RequestMapping("/jjg/fbgc/qlgc/xb/szd")
@CrossOrigin
public class JjgFbgcQlgcXbSzdController {

    @Autowired
    private JjgFbgcQlgcXbSzdService jjgFbgcQlgcXbSzdService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname, String htd) throws IOException {
        String fileName = "28桥梁下部墩台垂直度.xlsx";
        String p = filespath+ File.separator +proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
        }
    }

    @ApiOperation("生成桥梁下部竖直度鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcQlgcXbSzdService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看桥梁下部竖直度鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcQlgcXbSzdService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("桥梁下部竖直度模板文件导出")
    @GetMapping("exportqlxbszd")
    public void exportqlxbszd(HttpServletResponse response){
        jjgFbgcQlgcXbSzdService.exportqlxbszd(response);
    }

    /**
     * @param file
     * @return
     */
    @ApiOperation(value = "桥梁下部竖直度数据文件导入")
    @PostMapping("importqlxbszd")
    public Result importqlxbszd(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcQlgcXbSzdService.importqlxbszd(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcQlgcXbSzd JjgFbgcQlgcXbSzd){
        //创建page对象
        Page<JjgFbgcQlgcXbSzd> pageParam=new Page<>(current,limit);
        if (JjgFbgcQlgcXbSzd != null){
            QueryWrapper<JjgFbgcQlgcXbSzd> wrapper=new QueryWrapper<>();
            wrapper.like("proname",JjgFbgcQlgcXbSzd.getProname());
            wrapper.like("htd",JjgFbgcQlgcXbSzd.getHtd());
            wrapper.like("fbgc",JjgFbgcQlgcXbSzd.getFbgc());
            Date jcsj = JjgFbgcQlgcXbSzd.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcQlgcXbSzd> pageModel = jjgFbgcQlgcXbSzdService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除桥梁下部竖直度数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcQlgcXbSzdService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getXbSzd{id}")
    public Result getXbSzd(@PathVariable String id) {
        JjgFbgcQlgcXbSzd user = jjgFbgcQlgcXbSzdService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改桥桥梁下部竖直度数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcQlgcXbSzd user) {
        boolean is_Success = jjgFbgcQlgcXbSzdService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

}

