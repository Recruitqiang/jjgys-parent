package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgLjxQl;
import glgc.jjgys.model.project.JjgLqsQl;
import glgc.jjgys.system.service.JjgLqsLjxqlService;
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
@RequestMapping("/project/info/lqs/ljxql")
@CrossOrigin
public class JjgLqsLjxqlController {

    @Autowired
    private JjgLqsLjxqlService jjgLqsLjxqlService;

    @ApiOperation("连接线桥梁清单文件导出")
    @GetMapping("exportLjxQL")
    public void exportLjxQL(HttpServletResponse response){
        jjgLqsLjxqlService.exportLjxQL(response);
    }


    @ApiOperation(value = "连接线桥梁清单信息导入")
    @PostMapping("importLjxQL")
    public Result importLjxQL(@RequestParam("file") MultipartFile file, @RequestParam String proname) {
        jjgLqsLjxqlService.importLjxQL(file,proname);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgLjxQl jjgLjxQl){
        //创建page对象
        Page<JjgLjxQl> pageParam=new Page<>(current,limit);
        //判断projectQueryVo对象是否为空，直接查全部
        if(jjgLjxQl == null){
            IPage<JjgLjxQl> pageModel = jjgLqsLjxqlService.page(pageParam,null);
            return Result.ok(pageModel);
        }else {
            //获取条件值，进行非空判断，条件封装
            String ljxqlname = jjgLjxQl.getLjxqlname();
            QueryWrapper<JjgLjxQl> wrapper=new QueryWrapper<>();
            if (!StringUtils.isEmpty(ljxqlname)){
                wrapper.like("name",ljxqlname);
            }
            wrapper.like("proname",jjgLjxQl.getProname());
            wrapper.orderByDesc("create_time");
            //调用方法分页查询
            IPage<JjgLjxQl> pageModel = jjgLqsLjxqlService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);

        }
    }

    @ApiOperation("批量删除连接线桥梁清单信息")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean ql = jjgLqsLjxqlService.removeByIds(idList);
        if(ql){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

}

