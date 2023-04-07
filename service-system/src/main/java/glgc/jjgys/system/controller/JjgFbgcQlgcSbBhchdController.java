package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcQlgcSbBhchd;
import glgc.jjgys.model.project.JjgFbgcQlgcSbJgcc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcQlgcSbBhchdService;
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
 *  保护层厚度
 * </p>
 *
 * @author wq
 * @since 2023-03-20
 */
@RestController
@RequestMapping("/jjg/fbgc/qlgc/sb/bhchd")
@CrossOrigin
public class JjgFbgcQlgcSbBhchdController {

    @Autowired
    private JjgFbgcQlgcSbBhchdService jjgFbgcQlgcSbBhchdService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname, String htd) throws IOException {
        String fileName = "31桥梁上部保护层厚度.xlsx";
        String p = filespath+ File.separator +proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
        }
    }

    @ApiOperation("生成保护层厚度鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcQlgcSbBhchdService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看保护层厚度鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcQlgcSbBhchdService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("保护层厚度模板文件导出")
    @GetMapping("export")
    public void export(HttpServletResponse response){
        jjgFbgcQlgcSbBhchdService.export(response);
    }


    /**
     * @param file
     * @return
     */
    @ApiOperation(value = "保护层厚度数据文件导入")
    @PostMapping("importbhchd")
    public Result importbhchd(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcQlgcSbBhchdService.importbhchd(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcQlgcSbBhchd jjgFbgcQlgcSbBhchd){
        //创建page对象
        Page<JjgFbgcQlgcSbBhchd> pageParam=new Page<>(current,limit);
        if (jjgFbgcQlgcSbBhchd != null){
            QueryWrapper<JjgFbgcQlgcSbBhchd> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcQlgcSbBhchd.getProname());
            wrapper.like("htd",jjgFbgcQlgcSbBhchd.getHtd());
            wrapper.like("fbgc",jjgFbgcQlgcSbBhchd.getFbgc());
            Date jcsj = jjgFbgcQlgcSbBhchd.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcQlgcSbBhchd> pageModel = jjgFbgcQlgcSbBhchdService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");

    }

    @ApiOperation("批量删除保护层厚度数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean ql = jjgFbgcQlgcSbBhchdService.removeByIds(idList);
        if(ql){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

}

