package glgc.jjgys.system.controller;

import com.baomidou.mybatisplus.core.metadata.IPage;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.MD5;
import glgc.jjgys.model.system.SysLoginLog;
import glgc.jjgys.model.system.SysPost;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.model.vo.SysLoginLogQueryVo;
import glgc.jjgys.model.vo.SysPostQueryVo;
import glgc.jjgys.system.service.SysPostService;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

@Api(tags = "岗位管理接口")
@RestController
@RequestMapping("/admin/system/sysPost")
public class SysPostController {

    @Autowired
    private SysPostService sysPostService;

    //条件分页查询登录日志
    @ApiOperation(value = "获取分页列表")
    @GetMapping("{page}/{limit}")
    public Result index(@PathVariable long page,
                        @PathVariable long limit,
                        SysPostQueryVo sysPostQueryVo) {
        IPage<SysPost> pageModel =  sysPostService.selectPage(page,limit,sysPostQueryVo);
        return Result.ok(pageModel);
    }

    @ApiOperation("根据id查询")
    @GetMapping("findPostById/{id}")
    public Result getPost(@PathVariable String id) {
        SysPost sysPost = sysPostService.getById(id);
        return Result.ok(sysPost);
    }

    @ApiOperation("添加岗位")
    @PostMapping("save")
    public Result save(@RequestBody SysPost sysPost) {
        boolean res = sysPostService.save(sysPost);
        if(res) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

    @ApiOperation("修改岗位")
    @PostMapping("update")
    public Result update(@RequestBody SysPost sysPost) {
        boolean is_Success = sysPostService.updateById(sysPost);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

    @ApiOperation("删除岗位")
    @DeleteMapping("remove/{id}")
    public Result remove(@PathVariable String id) {
        boolean is_Success = sysPostService.removeById(id);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }



}
