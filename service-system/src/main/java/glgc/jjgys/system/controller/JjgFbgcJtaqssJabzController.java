package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabx;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabxfhl;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabz;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcJtaqssJabzService;
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
 *  前端控制器
 * </p>
 *
 * @author wq
 * @since 2023-03-01
 */
@RestController
@RequestMapping("/jjg/fbgc/jtaqss/jabz")
@CrossOrigin
public class JjgFbgcJtaqssJabzController {

    @Autowired
    private JjgFbgcJtaqssJabzService jjgFbgcJtaqssJabzService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;


    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public Result downloadExport(HttpServletResponse response,String proname,String htd) throws IOException {
        String fileName = "56交安标志.xlsx";
        String p = filespath+ File.separator+proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
            return Result.ok();
        }else {
            return Result.fail().message("还未生成鉴定表");
        }
    }

    @ApiOperation("生成交安标志鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcJtaqssJabzService.generateJdb(commonInfoVo);

    }

    @ApiOperation("交安标志模板文件导出")
    @GetMapping("exportjabz")
    public void exportjabz(HttpServletResponse response){
        jjgFbgcJtaqssJabzService.exportjabz(response);
    }

    @ApiOperation(value = "交安标志数据文件导入")
    @PostMapping("importjabz")
    public Result importjabz(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcJtaqssJabzService.importjabz(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcJtaqssJabz jjgFbgcJtaqssJabz) {
        //创建page对象
        Page<JjgFbgcJtaqssJabz> pageParam = new Page<>(current, limit);
        if (jjgFbgcJtaqssJabz != null) {
            QueryWrapper<JjgFbgcJtaqssJabz> wrapper = new QueryWrapper<>();
            wrapper.like("proname", jjgFbgcJtaqssJabz.getProname());
            wrapper.like("htd", jjgFbgcJtaqssJabz.getHtd());
            wrapper.like("fbgc", jjgFbgcJtaqssJabz.getFbgc());
            Date jcsj = jjgFbgcJtaqssJabz.getJcsj();
            if (!StringUtils.isEmpty(jcsj)) {
                wrapper.like("jcsj", jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcJtaqssJabz> pageModel = jjgFbgcJtaqssJabzService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除交安标志数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcJtaqssJabzService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getJabz/{id}")
    public Result getJabz(@PathVariable String id) {
        JjgFbgcJtaqssJabz user = jjgFbgcJtaqssJabzService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改交安标志数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcJtaqssJabz user) {
        boolean is_Success = jjgFbgcJtaqssJabzService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

}

