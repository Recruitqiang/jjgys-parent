package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgLjxSd;
import glgc.jjgys.model.project.JjgLqsQl;
import glgc.jjgys.system.service.JjgLqsLjxsdService;
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
@RequestMapping("/project/info/lqs/ljxsd")
@CrossOrigin
public class JjgLqsLjxsdController {

    @Autowired
    private JjgLqsLjxsdService jjgLqsLjxsdService;

    @ApiOperation("连接线隧道清单文件导出")
    @GetMapping("exportLjxsd")
    public void exportLjxsd(HttpServletResponse response){
        jjgLqsLjxsdService.exportLjxsd(response);
    }


    @ApiOperation(value = "连接线隧道清单信息导入")
    @PostMapping("importLjxsd")
    public Result importLjxsd(@RequestParam("file") MultipartFile file, @RequestParam String proname) {
        jjgLqsLjxsdService.importLjxsd(file,proname);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgLjxSd jjgLjxSd){
        //创建page对象
        Page<JjgLjxSd> pageParam=new Page<>(current,limit);
        //判断projectQueryVo对象是否为空，直接查全部
        if(jjgLjxSd == null){
            IPage<JjgLjxSd> pageModel = jjgLqsLjxsdService.page(pageParam,null);
            return Result.ok(pageModel);
        }else {
            //获取条件值，进行非空判断，条件封装
            String ljxsdname = jjgLjxSd.getLjxsdname();
            QueryWrapper<JjgLjxSd> wrapper=new QueryWrapper<>();
            if (!StringUtils.isEmpty(ljxsdname)){
                wrapper.like("name",ljxsdname);
            }
            wrapper.like("proname",jjgLjxSd.getProname());
            wrapper.orderByDesc("create_time");
            //调用方法分页查询
            IPage<JjgLjxSd> pageModel = jjgLqsLjxsdService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);

        }
    }

    @ApiOperation("批量删除连接线隧道清单信息")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean ql = jjgLqsLjxsdService.removeByIds(idList);
        if(ql){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

}

