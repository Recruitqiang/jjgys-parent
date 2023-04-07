package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcLjgcHdgqd;
import glgc.jjgys.model.project.JjgFbgcLjgcHdjgcc;
import glgc.jjgys.model.project.JjgFbgcLjgcXqgqd;
import glgc.jjgys.model.project.JjgFbgcLjgcXqjgcc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcLjgcXqjgccService;
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
 * @since 2023-02-15
 */
@RestController
@RequestMapping("/jjg/fbgc/ljgc/xqjgcc")
@CrossOrigin
public class JjgFbgcLjgcXqjgccController {

    @Autowired
    private JjgFbgcLjgcXqjgccService jjgFbgcLjgcXqjgccService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;


    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response,String proname,String htd) throws IOException {
        String fileName = "07路基小桥结构尺寸.xlsx";
        String p = filespath+ File.separator+proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
        }
    }

    @ApiOperation("生成小桥结构尺寸鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcLjgcXqjgccService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看小桥结构尺寸鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcLjgcXqjgccService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("小桥结构尺寸模板文件导出")
    @GetMapping("exportxqjgcc")
    public void exportxqjgcc(HttpServletResponse response){
        jjgFbgcLjgcXqjgccService.exportxqjgcc(response);
    }


    @ApiOperation(value = "小桥结构尺寸数据文件导入")
    @PostMapping("importxqjgcc")
    public Result importxqjgcc(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcLjgcXqjgccService.importxqjgcc(file,commonInfoVo);
        return Result.ok();
    }
    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcLjgcXqjgcc jjgFbgcLjgcXqjgcc){
        //创建page对象
        Page<JjgFbgcLjgcXqjgcc> pageParam=new Page<>(current,limit);
        if (jjgFbgcLjgcXqjgcc != null){
            QueryWrapper<JjgFbgcLjgcXqjgcc> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLjgcXqjgcc.getProname());
            wrapper.like("htd",jjgFbgcLjgcXqjgcc.getHtd());
            wrapper.like("fbgc",jjgFbgcLjgcXqjgcc.getFbgc());
            Date jcsj = jjgFbgcLjgcXqjgcc.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcLjgcXqjgcc> pageModel = jjgFbgcLjgcXqjgccService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除小桥结构尺寸数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean ql = jjgFbgcLjgcXqjgccService.removeByIds(idList);
        if(ql){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getcXqjgcc/{id}")
    public Result getcXqjgcc(@PathVariable String id) {
        JjgFbgcLjgcXqjgcc user = jjgFbgcLjgcXqjgccService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改小桥结构尺寸数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcLjgcXqjgcc user) {
        boolean is_Success = jjgFbgcLjgcXqjgccService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

}

