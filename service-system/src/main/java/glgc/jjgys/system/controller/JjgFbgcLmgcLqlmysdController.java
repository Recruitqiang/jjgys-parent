package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcLjgcLjtsfysdHt;
import glgc.jjgys.model.project.JjgFbgcLmgcLqlmysd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcLmgcLqlmysdService;
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
import java.text.ParseException;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * <p>
 *  前端控制器
 * </p>
 *
 * @author wq
 * @since 2023-04-10
 */
@RestController
@RequestMapping("/jjg/fbgc/lmgc/lqlmysd")
public class JjgFbgcLmgcLqlmysdController {

    @Autowired
    private JjgFbgcLmgcLqlmysdService jjgFbgcLmgcLqlmysdService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname, String htd) throws IOException {
        String fileName = "12沥青路面压实度.xlsx";
        String p = filespath+ File.separator+proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
        }
    }

    @ApiOperation("生成沥青路面压实度鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcLmgcLqlmysdService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看沥青路面压实度鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcLmgcLqlmysdService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("沥青路面压实度模板文件导出")
    @GetMapping("exportlqlmysd")
    public void exportlqlmysd(HttpServletResponse response) {
        jjgFbgcLmgcLqlmysdService.exportlqlmysd(response);
    }

    @ApiOperation(value = "沥青路面压实度数据文件导入")
    @PostMapping("importlqlmysd")
    public Result importlqlmysd(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcLmgcLqlmysdService.importlqlmysd(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPageHt/{current}/{limit}")
    public Result findQueryPageHt(@PathVariable long current,
                                  @PathVariable long limit,
                                  @RequestBody JjgFbgcLmgcLqlmysd jjgFbgcLmgcLqlmysd){
        //创建page对象
        Page<JjgFbgcLmgcLqlmysd> pageParam=new Page<>(current,limit);
        if (jjgFbgcLmgcLqlmysd != null){
            QueryWrapper<JjgFbgcLmgcLqlmysd> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLmgcLqlmysd.getProname());
            wrapper.like("htd",jjgFbgcLmgcLqlmysd.getHtd());
            wrapper.like("fbgc",jjgFbgcLmgcLqlmysd.getFbgc());
            Date sysj = jjgFbgcLmgcLqlmysd.getJcsj();
            if (!StringUtils.isEmpty(sysj)){
                wrapper.like("sysj",sysj);
            }
            //调用方法分页查询
            IPage<JjgFbgcLmgcLqlmysd> pageModel = jjgFbgcLmgcLqlmysdService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");

    }

    @ApiOperation("批量沥青路面压实度数据")
    @DeleteMapping("removeBeatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcLmgcLqlmysdService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getysd/{id}")
    public Result getysd(@PathVariable String id) {
        JjgFbgcLmgcLqlmysd user = jjgFbgcLmgcLqlmysdService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改沥青路面压实度数据")
    @PostMapping("updateysd")
    public Result updateysd(@RequestBody JjgFbgcLmgcLqlmysd ysd) {
        boolean is_Success = jjgFbgcLmgcLqlmysdService.updateById(ysd);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

}

