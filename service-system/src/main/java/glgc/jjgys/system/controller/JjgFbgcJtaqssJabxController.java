package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabx;
import glgc.jjgys.model.project.JjgFbgcLjgcHdgqd;
import glgc.jjgys.model.project.JjgFbgcLjgcHdjgcc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcJtaqssJabxService;
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
@RequestMapping("/jjg/fbgc/jtaqss/jabx")
@CrossOrigin
public class JjgFbgcJtaqssJabxController {

    @Autowired
    private JjgFbgcJtaqssJabxService jjgFbgcJtaqssJabxService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;


    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public Result downloadExport(HttpServletResponse response,String proname,String htd) throws IOException {
        String fileName1 = "57交安标线厚度.xlsx";
        String fileName2 = "57交安标线白线逆反射系数.xlsx";
        String fileName3 = "57交安标线黄线逆反射系数.xlsx";
        String p = filespath+File.separator+proname+File.separator+htd+File.separator+fileName1;
        String bx = filespath+File.separator+proname+File.separator+htd+File.separator+fileName2;
        String hx = filespath+File.separator+proname+File.separator+htd+File.separator+fileName3;
        File hdfile = new File(p);
        File hf = new File(bx);
        File hxf = new File(hx);
        if (hdfile.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName1);
            if (hf.exists()){
                JjgFbgcCommonUtils.download(response,bx,fileName2);
            }
            if (hxf.exists()){
                JjgFbgcCommonUtils.download(response,hx,fileName3);
            }
            return Result.ok();
        }else {
            return Result.fail().message("还未生成鉴定表");
        }
    }

    @ApiOperation("生成交安标线鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcJtaqssJabxService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看交安标线鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcJtaqssJabxService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("交安标线模板文件导出")
    @GetMapping("exportjabx")
    public void exportjabx(HttpServletResponse response){
        jjgFbgcJtaqssJabxService.exportjabx(response);
    }

    @ApiOperation(value = "交安标线数据文件导入")
    @PostMapping("importjabx")
    public Result importjabx(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcJtaqssJabxService.importjabx(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcJtaqssJabx jjgFbgcJtaqssJabx) {
        //创建page对象
        Page<JjgFbgcJtaqssJabx> pageParam = new Page<>(current, limit);
        if (jjgFbgcJtaqssJabx != null) {
            QueryWrapper<JjgFbgcJtaqssJabx> wrapper = new QueryWrapper<>();
            wrapper.like("proname", jjgFbgcJtaqssJabx.getProname());
            wrapper.like("htd", jjgFbgcJtaqssJabx.getHtd());
            wrapper.like("fbgc", jjgFbgcJtaqssJabx.getFbgc());
            Date jcsj = jjgFbgcJtaqssJabx.getJcsj();
            if (!StringUtils.isEmpty(jcsj)) {
                wrapper.like("jcsj", jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcJtaqssJabx> pageModel = jjgFbgcJtaqssJabxService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }


    @ApiOperation("批量删除交安标线数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcJtaqssJabxService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getJabx/{id}")
    public Result getJabx(@PathVariable String id) {
        JjgFbgcJtaqssJabx user = jjgFbgcJtaqssJabxService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改交安标线数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcJtaqssJabx user) {
        boolean is_Success = jjgFbgcJtaqssJabxService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

}

