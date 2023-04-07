package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgFbgcQlgcQmgzsd;
import glgc.jjgys.model.project.JjgFbgcQlgcQmhp;
import glgc.jjgys.model.project.JjgFbgcQlgcSbBhchd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcQlgcQmhpService;
import glgc.jjgys.system.utils.JjgFbgcCommonUtils;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.*;

/**
 * <p>
 *  前端控制器
 * </p>
 *
 * @author wq
 * @since 2023-03-20
 */
@RestController
@RequestMapping("/jjg/fbgc/qlgc/qmhp")
@CrossOrigin
public class JjgFbgcQlgcQmhpController {

    @Autowired
    private JjgFbgcQlgcQmhpService jjgFbgcQlgcQmhpService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletRequest request, HttpServletResponse response, String proname, String htd, String fbgc) throws IOException {
        List<Map<String,Object>> qlmclist = jjgFbgcQlgcQmhpService.selectqlmc(proname,htd,fbgc);
        List list = new ArrayList<>();
        for (int i=0;i<qlmclist.size();i++){
            list.add(qlmclist.get(i).get("qlmc"));
        }
        String zipName = "35桥面横坡";
        JjgFbgcCommonUtils.batchDownloadFile(request,response,zipName,list,filespath+File.separator+proname+File.separator+htd);

    }

    @ApiOperation("生成桥面横坡鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcQlgcQmhpService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看桥面横坡鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcQlgcQmhpService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("桥面横坡模板文件导出")
    @GetMapping("export")
    public void export(HttpServletResponse response){
        jjgFbgcQlgcQmhpService.export(response);
    }


    /**
     * @param file
     * @return
     */
    @ApiOperation(value = "桥面横坡数据文件导入")
    @PostMapping("importqmhp")
    public Result importqmhp(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcQlgcQmhpService.importqmhp(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcQlgcQmhp jjgFbgcQlgcQmhp){
        //创建page对象
        Page<JjgFbgcQlgcQmhp> pageParam=new Page<>(current,limit);
        if (jjgFbgcQlgcQmhp != null){
            QueryWrapper<JjgFbgcQlgcQmhp> wrapper=new QueryWrapper<>();
            wrapper.like("proname",jjgFbgcQlgcQmhp.getProname());
            wrapper.like("htd",jjgFbgcQlgcQmhp.getHtd());
            wrapper.like("fbgc",jjgFbgcQlgcQmhp.getFbgc());
            Date jcsj = jjgFbgcQlgcQmhp.getJcsj();
            if (!StringUtils.isEmpty(jcsj)){
                wrapper.like("jcsj",jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcQlgcQmhp> pageModel = jjgFbgcQlgcQmhpService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");

    }

    @ApiOperation("批量删除桥面横坡数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean ql = jjgFbgcQlgcQmhpService.removeByIds(idList);
        if(ql){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getQmhp{id}")
    public Result getQmhp(@PathVariable String id) {
        JjgFbgcQlgcQmhp user = jjgFbgcQlgcQmhpService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改桥面横坡数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcQlgcQmhp user) {
        boolean is_Success = jjgFbgcQlgcQmhpService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }




}

