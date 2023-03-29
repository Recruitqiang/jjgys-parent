package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgLqsFhlm;
import glgc.jjgys.model.project.JjgLqsQl;
import glgc.jjgys.system.service.JjgLqsFhlmService;
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
 * @since 2023-01-10
 */
@RestController
@RequestMapping("/project/info/lqs/fhlm")
@CrossOrigin
public class JjgLqsFhlmController {

    @Autowired
    private JjgLqsFhlmService jjgLqsFhlmService;

    @ApiOperation("复合路面清单文件导出")
    @GetMapping("export")
    public void export(HttpServletResponse response){
        jjgLqsFhlmService.export(response);
    }


    @ApiOperation(value = "复合路面信息导入")
    @PostMapping("importFhlm")
    public Result importFhlm(@RequestParam("file") MultipartFile file, @RequestParam String proname) {
        jjgLqsFhlmService.importFhlm(file,proname);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgLqsFhlm jjgLqsFhlm){
        //创建page对象
        Page<JjgLqsFhlm> pageParam=new Page<>(current,limit);
        //判断projectQueryVo对象是否为空，直接查全部
        if(jjgLqsFhlm == null){
            IPage<JjgLqsFhlm> pageModel = jjgLqsFhlmService.page(pageParam,null);
            return Result.ok(pageModel);
        }else {
            //获取条件值，进行非空判断，条件封装
            String fhlmname = jjgLqsFhlm.getFhlmname();
            QueryWrapper<JjgLqsFhlm> wrapper=new QueryWrapper<>();
            if (!StringUtils.isEmpty(fhlmname)){
                wrapper.like("name",fhlmname);
            }
            wrapper.like("proname",jjgLqsFhlm.getProname());
            wrapper.orderByDesc("create_time");
            //调用方法分页查询
            IPage<JjgLqsFhlm> pageModel = jjgLqsFhlmService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);

        }
    }

    @ApiOperation("批量删除复合路面清单信息")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean ql = jjgLqsFhlmService.removeByIds(idList);
        if(ql){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }


}

