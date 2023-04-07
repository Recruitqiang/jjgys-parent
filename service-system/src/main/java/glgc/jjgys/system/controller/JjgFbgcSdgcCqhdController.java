package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcQlgcXbTqd;
import glgc.jjgys.model.project.JjgFbgcSdgcCqhd;
import glgc.jjgys.model.project.JjgFbgcSdgcCqtqd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcSdgcCqhdService;
import glgc.jjgys.system.service.JjgFbgcSdgcSdhpService;
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
 * @since 2023-03-27
 */
@RestController
@RequestMapping("/jjg/fbgc/sdgc/cqhd")
@CrossOrigin
public class JjgFbgcSdgcCqhdController {

    @Autowired
    private JjgFbgcSdgcCqhdService jjgFbgcSdgcCqhdService;


    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletRequest request,HttpServletResponse response, String proname, String htd, String fbgc) throws IOException {
        List<Map<String,Object>> sdmclist = jjgFbgcSdgcCqhdService.selectsdmc(proname,htd,fbgc);
        List list = new ArrayList<>();
        for (int i=0;i<sdmclist.size();i++){
            list.add(sdmclist.get(i).get("sdmc"));
        }
        String zipName = "39隧道衬砌厚度";
        JjgFbgcCommonUtils.batchDownloadFile(request,response,zipName,list,filespath+File.separator+proname+File.separator+htd);
    }

    @ApiOperation("生成隧道衬砌厚度鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcSdgcCqhdService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看隧道衬砌厚度鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcSdgcCqhdService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("隧道衬砌厚度模板文件导出")
    @GetMapping("exportsdcqhd")
    public void exportsdcqhd(HttpServletResponse response){
        jjgFbgcSdgcCqhdService.exportsdcqhd(response);
    }

    /**
     * @param file
     * @return
     */
    @ApiOperation(value = "隧道衬砌厚度数据文件导入")
    @PostMapping("importsdcqhd")
    public Result importsdcqhd(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcSdgcCqhdService.importsdcqhd(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcSdgcCqhd jjgFbgcSdgcCqhd){
        //创建page对象
        Page<JjgFbgcSdgcCqhd> pageParam=new Page<>(current,limit);
        if (jjgFbgcSdgcCqhd != null){
            QueryWrapper<JjgFbgcSdgcCqhd> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcSdgcCqhd.getProname());
            wrapper.like("htd",jjgFbgcSdgcCqhd.getHtd());
            wrapper.like("fbgc",jjgFbgcSdgcCqhd.getFbgc());
            Date jcsj = jjgFbgcSdgcCqhd.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcSdgcCqhd> pageModel = jjgFbgcSdgcCqhdService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除隧道衬砌厚度数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcSdgcCqhdService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getCqhd{id}")
    public Result getCqhd(@PathVariable String id) {
        JjgFbgcSdgcCqhd user = jjgFbgcSdgcCqhdService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改隧道衬砌厚度数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcSdgcCqhd user) {
        boolean is_Success = jjgFbgcSdgcCqhdService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }


}

