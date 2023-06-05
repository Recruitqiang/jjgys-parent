package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcLmgcHntlmqd;
import glgc.jjgys.model.project.JjgFbgcLmgcLmhp;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcLmgcLmhpService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * <p>
 *  前端控制器
 * </p>
 *
 * @author wq
 * @since 2023-04-22
 */
@RestController
@RequestMapping("/jjg/fbgc/lmgc/lmhp")
public class JjgFbgcLmgcLmhpController {

    @Autowired
    private JjgFbgcLmgcLmhpService jjgFbgcLmgcLmhpService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletRequest request, HttpServletResponse response, String proname, String htd, String fbgc) throws IOException {
        List<Map<String,String>> mclist = jjgFbgcLmgcLmhpService.selectmc(proname,htd,fbgc);
        List list = new ArrayList<>();
        for (int i=0;i<mclist.size();i++){
            list.add(mclist.get(i).get("lxlx"));
        }
        String zipName = "24路面横坡";
        JjgFbgcCommonUtils.batchDownloadlmhpFile(request,response,zipName,list,filespath+File.separator+proname+File.separator+htd);
    }


    @ApiOperation("生成路面横坡鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcLmgcLmhpService.generateJdb(commonInfoVo);
    }
    @ApiOperation("查看路面横坡鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,String>> jdjg = jjgFbgcLmgcLmhpService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("路面横坡模板文件导出")
    @GetMapping("exportlmhp")
    public void exportlmhp(HttpServletResponse response){
        jjgFbgcLmgcLmhpService.exportlmhp(response);
    }


    @ApiOperation(value = "路面横坡数据文件导入")
    @PostMapping("importlmhp")
    public Result importlmhp(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcLmgcLmhpService.importlmhp(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcLmgcLmhp jjgFbgcLmgcLmhp){
        //创建page对象
        Page<JjgFbgcLmgcLmhp> pageParam=new Page<>(current,limit);
        if (jjgFbgcLmgcLmhp != null){
            QueryWrapper<JjgFbgcLmgcLmhp> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLmgcLmhp.getProname());
            wrapper.like("htd",jjgFbgcLmgcLmhp.getHtd());
            wrapper.like("fbgc",jjgFbgcLmgcLmhp.getFbgc());
            Date jcsj = jjgFbgcLmgcLmhp.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcLmgcLmhp> pageModel = jjgFbgcLmgcLmhpService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删路面横坡强度数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcLmgcLmhpService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }
    }

    @ApiOperation("根据id查询")
    @GetMapping("getLmhp/{id}")
    public Result getLmhp(@PathVariable String id) {
        JjgFbgcLmgcLmhp user = jjgFbgcLmgcLmhpService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改路面横坡数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcLmgcLmhp user) {
        boolean is_Success = jjgFbgcLmgcLmhpService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

}

