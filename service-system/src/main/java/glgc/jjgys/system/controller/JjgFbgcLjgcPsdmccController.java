package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcLjgcHdjgcc;
import glgc.jjgys.model.project.JjgFbgcLjgcLjwcLcf;
import glgc.jjgys.model.project.JjgFbgcLjgcPsdmcc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcLjgcPsdmccService;
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
 * @since 2023-02-16
 */
@RestController
@RequestMapping("/jjg/fbgc/ljgc/psdmcc")
@CrossOrigin
public class JjgFbgcLjgcPsdmccController {

    @Autowired
    private JjgFbgcLjgcPsdmccService jjgFbgcLjgcPsdmccService;
    @Value(value = "${jjgys.path.filepath}")
    private String filespath;


    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response,String proname,String htd) throws IOException {
        String fileName = "04路基排水断面尺寸.xlsx";
        String p = filespath+ File.separator+proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
        }
    }

    @ApiOperation("生成排水断面尺寸鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcLjgcPsdmccService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看排水断面尺寸鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcLjgcPsdmccService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("排水断面尺寸模板文件导出")
    @GetMapping("exportpsdmcc")
    public void exportpsdmcc(HttpServletResponse response){
        jjgFbgcLjgcPsdmccService.exportpsdmcc(response);
    }


    /**
     * 遗留问题：导入时需要传入项目名称，合同段，分布工程
     * @param file
     * @return
     */
    @ApiOperation(value = "排水断面尺寸数据文件导入")
    @PostMapping("importpsdmcc")
    public Result importpsdmcc(@RequestParam("file") MultipartFile file,CommonInfoVo commonInfoVo) {
        jjgFbgcLjgcPsdmccService.importpsdmcc(file,commonInfoVo);
        return Result.ok();
    }

    //遗留问题：还需要根据项目名和合同段明查询
    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcLjgcPsdmcc jjgFbgcLjgcPsdmcc){
        //创建page对象
        Page<JjgFbgcLjgcPsdmcc> pageParam=new Page<>(current,limit);
        if (jjgFbgcLjgcPsdmcc != null){
            QueryWrapper<JjgFbgcLjgcPsdmcc> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLjgcPsdmcc.getProname());
            wrapper.like("htd",jjgFbgcLjgcPsdmcc.getHtd());
            wrapper.like("fbgc",jjgFbgcLjgcPsdmcc.getFbgc());
            Date jcsj = jjgFbgcLjgcPsdmcc.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcLjgcPsdmcc> pageModel = jjgFbgcLjgcPsdmccService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }

    @ApiOperation("批量删除排水断面尺寸数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean ql = jjgFbgcLjgcPsdmccService.removeByIds(idList);
        if(ql){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }
    @ApiOperation("根据id查询")
    @GetMapping("getPsdmcc/{id}")
    public Result getPsdmcc(@PathVariable String id) {
        JjgFbgcLjgcPsdmcc user = jjgFbgcLjgcPsdmccService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改排水断面尺寸数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcLjgcPsdmcc user) {
        boolean is_Success = jjgFbgcLjgcPsdmccService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

}

