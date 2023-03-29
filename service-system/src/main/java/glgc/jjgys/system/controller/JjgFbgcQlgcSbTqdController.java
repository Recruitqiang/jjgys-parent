package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcLjgcZdgqd;
import glgc.jjgys.model.project.JjgFbgcQlgcSbJgcc;
import glgc.jjgys.model.project.JjgFbgcQlgcSbTqd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcQlgcSbTqdService;
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
 *  桥梁上部砼强度
 * </p>
 *
 * @author wq
 * @since 2023-03-20
 */
@RestController
@RequestMapping("/jjg/fbgc/qlgc/sb/tqd")
@CrossOrigin
public class JjgFbgcQlgcSbTqdController {

    @Autowired
    private JjgFbgcQlgcSbTqdService jjgFbgcQlgcSbTqdService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;


    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public Result downloadExport(HttpServletResponse response, String proname, String htd) throws IOException {
        String fileName = "29桥梁上部砼强度.xlsx";
        String p = filespath+ File.separator +proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
            return Result.ok();
        }else {
            return Result.fail().message("还未生成鉴定表");
        }
    }

    @ApiOperation("生成桥梁上部砼强度鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcQlgcSbTqdService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看桥梁上部砼强度鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcQlgcSbTqdService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("桥梁上部砼强度模板文件导出")
    @GetMapping("exportqlsbtqd")
    public void exportqlsbtqd(HttpServletResponse response){
        jjgFbgcQlgcSbTqdService.exportqlsbtqd(response);
    }

    /**
     * @param file
     * @return
     */
    @ApiOperation(value = "桥梁上部砼强度数据文件导入")
    @PostMapping("importqlsbtqd")
    public Result importqlsbtqd(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcQlgcSbTqdService.importqlsbtqd(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcQlgcSbTqd jjgFbgcQlgcSbTqd){
        //创建page对象
        Page<JjgFbgcQlgcSbTqd> pageParam=new Page<>(current,limit);
        if (jjgFbgcQlgcSbTqd != null){
            QueryWrapper<JjgFbgcQlgcSbTqd> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcQlgcSbTqd.getProname());
            wrapper.like("htd",jjgFbgcQlgcSbTqd.getHtd());
            wrapper.like("fbgc",jjgFbgcQlgcSbTqd.getFbgc());
            Date jcsj = jjgFbgcQlgcSbTqd.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcQlgcSbTqd> pageModel = jjgFbgcQlgcSbTqdService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除桥梁上部砼强度数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcQlgcSbTqdService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getSbTqd{id}")
    public Result getSbTqd(@PathVariable String id) {
        JjgFbgcQlgcSbTqd user = jjgFbgcQlgcSbTqdService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改桥梁桥梁上部砼强度数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcQlgcSbTqd user) {
        boolean is_Success = jjgFbgcQlgcSbTqdService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

}

