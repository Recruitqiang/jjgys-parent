package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgLqsQl;
import glgc.jjgys.system.service.JjgLqsQlService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.util.List;

/**
 * <p>
 *  前端控制器
 * </p>
 *
 * @author wq
 * @since 2022-12-14
 */
@RestController
@RequestMapping("/project/info/lqs/ql")
@CrossOrigin
public class JjgLqsQlController {
    @Autowired
    private JjgLqsQlService jjgLqsQlService;

    //桥梁清单文件导出
    @ApiOperation("桥梁清单文件导出")
    @GetMapping("exportQL/{projectname}")
    public void exportQL(HttpServletResponse response,
                         @PathVariable String projectname){
        jjgLqsQlService.exportQL(response,projectname);
    }


    @ApiOperation(value = "桥梁信息导入")
    @PostMapping("importQL")
    public Result importQL(@RequestParam("file") MultipartFile file,@RequestParam String proname) {
        jjgLqsQlService.importQL(file,proname);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgLqsQl jjgLqsQl){
        //创建page对象
        Page<JjgLqsQl> pageParam=new Page<>(current,limit);
        //判断projectQueryVo对象是否为空，直接查全部
        if(jjgLqsQl == null){
            IPage<JjgLqsQl> pageModel = jjgLqsQlService.page(pageParam,null);
            return Result.ok(pageModel);
        }else {
            //获取条件值，进行非空判断，条件封装
            String qlname = jjgLqsQl.getQlname();
            QueryWrapper<JjgLqsQl> wrapper=new QueryWrapper<>();
            if (!StringUtils.isEmpty(qlname)){
                wrapper.like("qlname",qlname);
            }
            wrapper.like("proname",jjgLqsQl.getProname());
            wrapper.orderByDesc("create_time");
            //调用方法分页查询
            IPage<JjgLqsQl> pageModel = jjgLqsQlService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);

        }
    }

    @ApiOperation("批量删除桥梁清单信息")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean ql = jjgLqsQlService.removeByIds(idList);
        if(ql){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

}

