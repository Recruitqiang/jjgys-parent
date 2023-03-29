package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcJtaqssJathlqd;
import glgc.jjgys.model.project.JjgFbgcLjgcHdgqd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcJtaqssJathlqdService;
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
@RequestMapping("/jjg/fbgc/jtaqss/jathlqd")
@CrossOrigin
public class JjgFbgcJtaqssJathlqdController {

    @Autowired
    private JjgFbgcJtaqssJathlqdService jjgFbgcJtaqssJathlqdService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;


    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public Result downloadExport(HttpServletResponse response, String proname, String htd) throws IOException {
        String fileName = "59交安砼护栏强度.xlsx";
        String p = filespath+ File.separator+proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
            return Result.ok();
        }else {
            return Result.fail().message("还未生成鉴定表");
        }
    }

    @ApiOperation("生成交安砼护栏强度鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcJtaqssJathlqdService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看交安砼护栏强度鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcJtaqssJathlqdService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("交安砼护栏强度模板文件导出")
    @GetMapping("exportjathlqd")
    public void exportjathlqd(HttpServletResponse response){
        jjgFbgcJtaqssJathlqdService.exportjathlqd(response);
    }

    @ApiOperation(value = "交安砼护栏强度数据文件导入")
    @PostMapping("importjathlqd")
    public Result importjathlqd(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcJtaqssJathlqdService.importjathlqd(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcJtaqssJathlqd jjgFbgcJtaqssJathlqd) {
        //创建page对象
        Page<JjgFbgcJtaqssJathlqd> pageParam = new Page<>(current, limit);
        if (jjgFbgcJtaqssJathlqd != null) {
            QueryWrapper<JjgFbgcJtaqssJathlqd> wrapper = new QueryWrapper<>();
            wrapper.like("proname", jjgFbgcJtaqssJathlqd.getProname());
            wrapper.like("htd", jjgFbgcJtaqssJathlqd.getHtd());
            wrapper.like("fbgc", jjgFbgcJtaqssJathlqd.getFbgc());
            Date jcsj = jjgFbgcJtaqssJathlqd.getJcsj();
            if (!StringUtils.isEmpty(jcsj)) {
                wrapper.like("jcsj", jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcJtaqssJathlqd> pageModel = jjgFbgcJtaqssJathlqdService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("根据id查询")
    @GetMapping("getJathlqd/{id}")
    public Result getJathlqd(@PathVariable String id) {
        JjgFbgcJtaqssJathlqd user = jjgFbgcJtaqssJathlqdService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改交安砼护栏强度数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcJtaqssJathlqd user) {
        boolean is_Success = jjgFbgcJtaqssJathlqdService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

    @ApiOperation("批量删除交安砼护栏强度数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcJtaqssJathlqdService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

}

