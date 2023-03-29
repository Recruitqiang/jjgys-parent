package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgLqsFhlm;
import glgc.jjgys.model.project.JjgLqsHntlmzd;
import glgc.jjgys.system.service.JjgLqsHntlmzdService;
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
@RequestMapping("/project/info/lqs/hntlmzd")
@CrossOrigin
public class JjgLqsHntlmzdController {
    @Autowired
    private JjgLqsHntlmzdService jjgLqsHntlmzdService;

    @ApiOperation("复合路面清单文件导出")
    @GetMapping("export")
    public void export(HttpServletResponse response){
        jjgLqsHntlmzdService.export(response);
    }


    @ApiOperation(value = "复合路面信息导入")
    @PostMapping("importhntlmzd")
    public Result importhntlmzd(@RequestParam("file") MultipartFile file, @RequestParam String proname) {
        jjgLqsHntlmzdService.importhntlmzd(file,proname);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgLqsHntlmzd jjgLqsHntlmzd){
        //创建page对象
        Page<JjgLqsHntlmzd> pageParam=new Page<>(current,limit);
        //判断projectQueryVo对象是否为空，直接查全部
        if(jjgLqsHntlmzd == null){
            IPage<JjgLqsHntlmzd> pageModel = jjgLqsHntlmzdService.page(pageParam,null);
            return Result.ok(pageModel);
        }else {
            //获取条件值，进行非空判断，条件封装
            String hntlmzdName = jjgLqsHntlmzd.getHntlmname();
            QueryWrapper<JjgLqsHntlmzd> wrapper=new QueryWrapper<>();
            if (!StringUtils.isEmpty(hntlmzdName)){
                wrapper.like("name",hntlmzdName);
            }
            wrapper.like("proname",jjgLqsHntlmzd.getProname());
            wrapper.orderByDesc("create_time");
            //调用方法分页查询
            IPage<JjgLqsHntlmzd> pageModel = jjgLqsHntlmzdService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);

        }
    }

    @ApiOperation("批量删除复合路面清单信息")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean ql = jjgLqsHntlmzdService.removeByIds(idList);
        if(ql){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

}

