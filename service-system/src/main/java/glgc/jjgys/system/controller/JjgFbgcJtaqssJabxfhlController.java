package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabx;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabxfhl;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabz;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcJtaqssJabxfhlService;
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

/**
 * <p>
 *  前端控制器
 * </p>
 *
 * @author wq
 * @since 2023-03-01
 */
@RestController
@RequestMapping("/jjg/fbgc/jtaqss/jabxfhl")
@CrossOrigin
public class JjgFbgcJtaqssJabxfhlController {

    @Autowired
    private JjgFbgcJtaqssJabxfhlService jjgFbgcJtaqssJabxfhlService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public Result downloadExport(HttpServletResponse response, String proname, String htd) throws IOException {
        String fileName = "58交安钢防护栏.xlsx";
        String p = filespath+ File.separator+proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
            return Result.ok();
        }else {
            return Result.fail().message("还未生成鉴定表");
        }
    }

    @ApiOperation("生成交安波形防护栏鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcJtaqssJabxfhlService.generateJdb(commonInfoVo);

    }

    @ApiOperation("交安波形防护模板文件导出")
    @GetMapping("exportjabxfhl")
    public void exportjabxfhl(HttpServletResponse response) throws IOException {
        jjgFbgcJtaqssJabxfhlService.exportjabxfhl(response);
    }


    @ApiOperation(value = "交安波形防护数据文件导入")
    @PostMapping("importjabxfhl")
    public Result importjabxfhl(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) throws IOException, ParseException {
        jjgFbgcJtaqssJabxfhlService.importjabxfhl(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcJtaqssJabxfhl jjgFbgcJtaqssJabxfhl) {
        //创建page对象
        Page<JjgFbgcJtaqssJabxfhl> pageParam = new Page<>(current, limit);
        if (jjgFbgcJtaqssJabxfhl != null) {
            QueryWrapper<JjgFbgcJtaqssJabxfhl> wrapper = new QueryWrapper<>();
            wrapper.like("proname", jjgFbgcJtaqssJabxfhl.getProname());
            wrapper.like("htd", jjgFbgcJtaqssJabxfhl.getHtd());
            wrapper.like("fbgc", jjgFbgcJtaqssJabxfhl.getFbgc());
            Date jcsj = jjgFbgcJtaqssJabxfhl.getJcsj();
            if (!StringUtils.isEmpty(jcsj)) {
                wrapper.like("jcsj", jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcJtaqssJabxfhl> pageModel = jjgFbgcJtaqssJabxfhlService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除交安波形防护数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcJtaqssJabxfhlService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getJabxfhl/{id}")
    public Result getJabxfhl(@PathVariable String id) {
        JjgFbgcJtaqssJabxfhl user = jjgFbgcJtaqssJabxfhlService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改交安波形防护数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcJtaqssJabxfhl user) {
        boolean is_Success = jjgFbgcJtaqssJabxfhlService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }



}

