package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgLqsSd;
import glgc.jjgys.model.project.JjgSfz;
import glgc.jjgys.system.service.JjgLqsSfzService;
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
@RequestMapping("/project/info/lqs/sfz")
@CrossOrigin
public class JjgLqsSfzController {

    @Autowired
    private JjgLqsSfzService jjgLqsSfzService;

    @ApiOperation("收费站清单文件导出")
    @GetMapping("export")
    public void exportSD(HttpServletResponse response){
        jjgLqsSfzService.exportSD(response);
    }

    @ApiOperation(value = "收费站清单导入")
    @PostMapping("import")
    public Result importQL(@RequestParam("file") MultipartFile file, @RequestParam String proname) {
        jjgLqsSfzService.importSD(file,proname);
        return Result.ok();
    }

    /**
     * 分页查询
     * @param current
     * @param limit
     * @param jjgSfz
     * @return
     */
    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgSfz jjgSfz){
        //创建page对象
        Page<JjgSfz> pageParam=new Page<>(current,limit);
        //判断projectQueryVo对象是否为空，直接查全部
        if(jjgSfz == null){
            IPage<JjgSfz> pageModel = jjgLqsSfzService.page(pageParam,null);
            return Result.ok(pageModel);
        }else {
            //获取条件值，进行非空判断，条件封装
            String sdname = jjgSfz.getZdsfzname();
            QueryWrapper<JjgSfz> wrapper=new QueryWrapper<>();
            if (!StringUtils.isEmpty(sdname)){
                wrapper.like("name",sdname);
            }
            wrapper.like("proname",jjgSfz.getProname());
            wrapper.orderByDesc("create_time");
            //调用方法分页查询
            IPage<JjgSfz> pageModel = jjgLqsSfzService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);

        }
    }

    @ApiOperation("批量删除收费站清单信息")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean ql = jjgLqsSfzService.removeByIds(idList);
        if(ql){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }


}

