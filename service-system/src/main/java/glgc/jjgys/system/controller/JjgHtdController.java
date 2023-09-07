package glgc.jjgys.system.controller;

import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgHtd;
import glgc.jjgys.system.service.JjgHtdService;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.bind.annotation.*;

import java.util.ArrayList;
import java.util.List;

@Api(tags = "合同段")
@RestController
@Transactional
@RequestMapping("/project/info/htd")
public class JjgHtdController {

    @Autowired
    private JjgHtdService jjgHtdService;


    @ApiOperation("批量删除合同段信息")
    @Transactional
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        jjgHtdService.removeData(idList);
        jjgHtdService.removeByIds(idList);
        return Result.ok();

    }

    @ApiOperation("添加合同段")
    @PostMapping("save")
    public Result save(@RequestBody JjgHtd jjgHtd) {
        boolean res = jjgHtdService.addhtd(jjgHtd);
        return Result.ok();
    }



    @ApiOperation("删除")
    @PostMapping("remove")
    public Result remove(@RequestBody JjgHtd jjgHtd) {
        jjgHtdService.removeById(jjgHtd.getId());
        return Result.ok();
    }

    /**
     * 分页查询
     */
    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgHtd jjgHtd){
        System.out.println(jjgHtd.getProname());
        //创建page对象
        Page<JjgHtd> pageParam=new Page<>(current,limit);
        //判断projectQueryVo对象是否为空，直接查全部
        if(jjgHtd == null){
            IPage<JjgHtd> pageModel = jjgHtdService.page(pageParam,null);
            return Result.ok(pageModel);
        }else {
            //获取条件值，进行非空判断，条件封装
            String jjgHtdName = jjgHtd.getName();
            QueryWrapper<JjgHtd> htdWrapper=new QueryWrapper<>();
            if (!StringUtils.isEmpty(jjgHtdName)){
                htdWrapper.like("name",jjgHtdName);
            }
            htdWrapper.like("proname",jjgHtd.getProname());
            htdWrapper.orderByDesc("create_time");
            //调用方法分页查询
            IPage<JjgHtd> pageModel = jjgHtdService.page(pageParam, htdWrapper);
            //返回
            return Result.ok(pageModel);

        }
    }

}
