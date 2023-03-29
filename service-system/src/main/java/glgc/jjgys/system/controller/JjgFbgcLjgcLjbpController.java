package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcJtaqssJathldmcc;
import glgc.jjgys.model.project.JjgFbgcLjgcHdjgcc;
import glgc.jjgys.model.project.JjgFbgcLjgcLjbp;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcLjgcLjbpService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.beans.IntrospectionException;
import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.text.ParseException;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * <p>
 *  路基边坡
 * </p>
 *
 * @author wq
 * @since 2023-02-21
 */
@RestController
@RequestMapping("/jjg/fbgc/ljgc/ljbp")
@CrossOrigin
public class JjgFbgcLjgcLjbpController {

    @Autowired
    private JjgFbgcLjgcLjbpService jjgFbgcLjgcLjbpService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public Result downloadExport(HttpServletResponse response,String proname,String htd) throws IOException {
        String fileName = "03路基边坡.xlsx";
        String p = filespath+ File.separator+proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
            return Result.ok();
        }else {
            return Result.fail().message("还未生成鉴定表");
        }
    }

    @ApiOperation("生成路基边坡鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcLjgcLjbpService.generateJdb(commonInfoVo);

    }
    @ApiOperation("查看路基边坡鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcLjgcLjbpService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("路基边坡模板文件导出")
    @GetMapping("exportljbp")
    public void exportljbp(HttpServletResponse response) throws IOException {
        jjgFbgcLjgcLjbpService.exportljbp(response);
    }

    @ApiOperation(value = "路基边坡数据文件导入")
    @PostMapping("importljbp")
    public Result importljbp(@RequestParam("file") MultipartFile file,CommonInfoVo commonInfoVo) throws IOException, ParseException {
        jjgFbgcLjgcLjbpService.importljbp(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcLjgcLjbp jjgFbgcLjgcLjbp){
        //创建page对象
        Page<JjgFbgcLjgcLjbp> pageParam=new Page<>(current,limit);
        if (jjgFbgcLjgcLjbp != null){
            QueryWrapper<JjgFbgcLjgcLjbp> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLjgcLjbp.getProname());
            wrapper.like("htd",jjgFbgcLjgcLjbp.getHtd());
            wrapper.like("fbgc",jjgFbgcLjgcLjbp.getFbgc());
            Date jcsj = jjgFbgcLjgcLjbp.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcLjgcLjbp> pageModel = jjgFbgcLjgcLjbpService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除路基边坡数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean ql = jjgFbgcLjgcLjbpService.removeByIds(idList);
        if(ql){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getLjbp/{id}")
    public Result getLjbp(@PathVariable String id) {
        JjgFbgcLjgcLjbp user = jjgFbgcLjgcLjbpService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改路基边坡数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcLjgcLjbp user) {
        boolean is_Success = jjgFbgcLjgcLjbpService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

}

