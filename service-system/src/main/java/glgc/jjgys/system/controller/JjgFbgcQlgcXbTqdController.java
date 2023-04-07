package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcQlgcSbTqd;
import glgc.jjgys.model.project.JjgFbgcQlgcXbSzd;
import glgc.jjgys.model.project.JjgFbgcQlgcXbTqd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcQlgcXbTqdService;
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
@RequestMapping("/jjg/fbgc/qlgc/xb/tqd")
public class JjgFbgcQlgcXbTqdController {

    @Autowired
    private JjgFbgcQlgcXbTqdService jjgFbgcQlgcXbTqdService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;


    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname, String htd) throws IOException {
        String fileName = "25桥梁下部墩台砼强度.xlsx";
        String p = filespath+ File.separator +proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
        }
    }

    @ApiOperation("生成桥梁下部砼强度鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcQlgcXbTqdService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看桥梁下部砼强度鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcQlgcXbTqdService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("桥梁下部砼强度模板文件导出")
    @GetMapping("exportqlxbtqd")
    public void exportqlxbtqd(HttpServletResponse response){
        jjgFbgcQlgcXbTqdService.exportqlxbtqd(response);
    }

    /**
     * @param file
     * @return
     */
    @ApiOperation(value = "桥梁下部砼强度数据文件导入")
    @PostMapping("importqlxbtqd")
    public Result importqlxbtqd(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcQlgcXbTqdService.importqlxbtqd(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcQlgcXbTqd jjgFbgcQlgcXbTqd){
        //创建page对象
        Page<JjgFbgcQlgcXbTqd> pageParam=new Page<>(current,limit);
        if (jjgFbgcQlgcXbTqd != null){
            QueryWrapper<JjgFbgcQlgcXbTqd> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcQlgcXbTqd.getProname());
            wrapper.like("htd",jjgFbgcQlgcXbTqd.getHtd());
            wrapper.like("fbgc",jjgFbgcQlgcXbTqd.getFbgc());
            Date jcsj = jjgFbgcQlgcXbTqd.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcQlgcXbTqd> pageModel = jjgFbgcQlgcXbTqdService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除桥梁上部砼强度数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcQlgcXbTqdService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getXbTqd{id}")
    public Result getXbTqd(@PathVariable String id) {
        JjgFbgcQlgcXbTqd user = jjgFbgcQlgcXbTqdService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改桥桥梁上部砼强度数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcQlgcXbTqd user) {
        boolean is_Success = jjgFbgcQlgcXbTqdService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

}

