package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcLjgcHdgqd;
import glgc.jjgys.model.project.JjgFbgcLjgcLjcj;
import glgc.jjgys.model.project.JjgFbgcLjgcLjtsfysdHt;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcLjgcLjtsfysdHtService;
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
import java.util.Map;

/**
 * <p>
 *  路基土石方压实度_灰土
 * </p>
 *
 * @author wq
 * @since 2023-02-21
 */
@RestController
@RequestMapping("/jjg/fbgc/ljgc/ljtsfysd")
@CrossOrigin
public class JjgFbgcLjgcLjtsfysdHtController {

    @Autowired
    private JjgFbgcLjgcLjtsfysdHtService jjgFbgcLjgcLjtsfysdHtService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public Result downloadExport(HttpServletResponse response,String proname,String htd) throws IOException {
        String fileName = "01路基土石方压实度.xlsx";
        String p = filespath+ File.separator+proname+File.separator+htd+File.separator+fileName;
        File file = new File(p);
        if (file.exists()){
            JjgFbgcCommonUtils.download(response,p,fileName);
            return Result.ok();
        }else {
            return Result.fail().message("还未生成鉴定表");
        }
    }

    @ApiOperation("生成路基土石方压实度(灰土+砂砾)鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcLjgcLjtsfysdHtService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看路基土石方压实度_灰土鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcLjgcLjtsfysdHtService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("路基土石方压实度_灰土模板文件导出")
    @GetMapping("exportysdht")
    public void exportysdht(HttpServletResponse response) throws IOException {
        jjgFbgcLjgcLjtsfysdHtService.exportysdht(response);
    }

    @ApiOperation(value = "路基土石方压实度_灰土数据文件导入")
    @PostMapping("importysdht")
    public Result importysdht(@RequestParam("file") MultipartFile file,CommonInfoVo commonInfoVo) throws IOException, ParseException {
        jjgFbgcLjgcLjtsfysdHtService.importysdht(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPageHt/{current}/{limit}")
    public Result findQueryPageHt(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcLjgcLjtsfysdHt jjgFbgcLjgcLjtsfysdHt){
        //创建page对象
        Page<JjgFbgcLjgcLjtsfysdHt> pageParam=new Page<>(current,limit);
        if (jjgFbgcLjgcLjtsfysdHt != null){
            QueryWrapper<JjgFbgcLjgcLjtsfysdHt> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcLjgcLjtsfysdHt.getProname());
            wrapper.like("htd",jjgFbgcLjgcLjtsfysdHt.getHtd());
            wrapper.like("fbgc",jjgFbgcLjgcLjtsfysdHt.getFbgc());
            Date sysj = jjgFbgcLjgcLjtsfysdHt.getSysj();
            if (!StringUtils.isEmpty(sysj)){
                wrapper.like("sysj",sysj);
            }
            //调用方法分页查询
            IPage<JjgFbgcLjgcLjtsfysdHt> pageModel = jjgFbgcLjgcLjtsfysdHtService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");

    }

    @ApiOperation("批量删除路基土石方压实度_灰土数据")
    @DeleteMapping("removeBeatchHt")
    public Result removeBeatchHt(@RequestBody List<String> idList){
        boolean hd = jjgFbgcLjgcLjtsfysdHtService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getHt/{id}")
    public Result getHt(@PathVariable String id) {
        JjgFbgcLjgcLjtsfysdHt user = jjgFbgcLjgcLjtsfysdHtService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改路基土石方压实度_灰土数据")
    @PostMapping("updateHt")
    public Result updateHt(@RequestBody JjgFbgcLjgcLjtsfysdHt user) {
        boolean is_Success = jjgFbgcLjgcLjtsfysdHtService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }


}

