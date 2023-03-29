package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcLjgcHdgqd;
import glgc.jjgys.model.project.JjgFbgcQlgcQmgzsd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcQlgcQmgzsdService;
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
 *  桥面构造深度
 * </p>
 *
 * @author wq
 * @since 2023-03-20
 */
@RestController
@RequestMapping("/jjg/fbgc/qlgc/qmgzsd")
public class JjgFbgcQlgcQmgzsdController {

    @Autowired
    private JjgFbgcQlgcQmgzsdService jjgFbgcQlgcQmgzsdService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname, String htd) throws IOException {
        String fileName = "桥面构造深度.xlsx";
        String p = filespath+ File.separator+proname+File.separator+htd+File.separator+fileName;
        JjgFbgcCommonUtils.download(response,p,fileName);
    }

    @ApiOperation("生成桥面构造深度鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcQlgcQmgzsdService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看桥面构造深度鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcQlgcQmgzsdService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("桥面构造深度模板文件导出")
    @GetMapping("exportqmgzsd")
    public void exportqmgzsd(HttpServletResponse response){
        jjgFbgcQlgcQmgzsdService.exportqmgzsd(response);
    }

    /**
     * @param file
     * @return
     */
    @ApiOperation(value = "桥面构造深度数据文件导入")
    @PostMapping("importqmgzsd")
    public Result importqmgzsd(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcQlgcQmgzsdService.importqmgzsd(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcQlgcQmgzsd jjgFbgcQlgcQmgzsd){
        //创建page对象
        Page<JjgFbgcQlgcQmgzsd> pageParam=new Page<>(current,limit);
        if (jjgFbgcQlgcQmgzsd != null){
            QueryWrapper<JjgFbgcQlgcQmgzsd> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcQlgcQmgzsd.getProname());
            wrapper.like("htd",jjgFbgcQlgcQmgzsd.getHtd());
            wrapper.like("fbgc",jjgFbgcQlgcQmgzsd.getFbgc());
            Date jcsj = jjgFbgcQlgcQmgzsd.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcQlgcQmgzsd> pageModel = jjgFbgcQlgcQmgzsdService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("根据id查询")
    @GetMapping("getQmgzsd/{id}")
    public Result getQmgzsd(@PathVariable String id) {
        JjgFbgcQlgcQmgzsd user = jjgFbgcQlgcQmgzsdService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改桥面构造深度数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcQlgcQmgzsd user) {
        boolean is_Success = jjgFbgcQlgcQmgzsdService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

    @ApiOperation("批量删除桥面构造深度数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcQlgcQmgzsdService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

}

