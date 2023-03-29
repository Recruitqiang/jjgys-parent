package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.system.Project;
import glgc.jjgys.model.vo.ProjectQueryVo;
import glgc.jjgys.system.service.ProjectService;
import glgc.jjgys.system.service.SysMenuService;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@Api(tags = "项目管理接口")
@RestController
@RequestMapping(value = "/project")
@CrossOrigin//解决跨域
public class ProjectController {

    @Autowired
    private ProjectService projectService;

    /**
     * 新增项目 包括新增合同段信息 路桥隧信息
     */
    @PostMapping("/add")
    public Result add(@RequestBody Project project){
        String proName = project.getProName();
        projectService.addOtherInfo(proName);
        boolean res = projectService.save(project);
        if(res){
            return Result.ok(null).message("增加成功");
        }else {
            return Result.fail(null).message("增加失败");
        }

    }

    /**
     * 删除项目
     */
    @DeleteMapping("/remove/{id}")
    public Result remove(@PathVariable String id){
        projectService.removeById(id);
        return Result.ok(null);
    }

    @ApiOperation("批量删除项目信息")
    @DeleteMapping("removeBatch")
    //传json数组[1,2,3]，用List接收
    public Result removeBeatch(@RequestBody List<String> idList){
        System.out.println(idList);
        boolean isSuccess = projectService.removeByIds(idList);
        if(isSuccess){
            return Result.ok(null);
        } else {
            return Result.fail(null).message("删除失败！");
        }

    }

    /**
     * 查询所有的项目
     */
    @GetMapping("findAll")
    public Result findAllProject(){
        List<Project> list = projectService.list();
        return Result.ok(list);
    }

    /**
     * 分页查询
     */
    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody ProjectQueryVo projectQueryVo){
        //创建page对象
        Page<Project> pageParam=new Page<>(current,limit);
        //判断projectQueryVo对象是否为空，直接查全部
        if(projectQueryVo == null){
            IPage<Project> pageModel = projectService.page(pageParam,null);
            return Result.ok(pageModel);
        }else {
            //获取条件值，进行非空判断，条件封装
            String proName = projectQueryVo.getProName();
            QueryWrapper<Project> wrapper=new QueryWrapper<>();
            if (!StringUtils.isEmpty(proName)){
                wrapper.like("proName",proName);
            }
            wrapper.orderByDesc("create_time");
            //调用方法分页查询
            IPage<Project> pageModel = projectService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);

        }
    }
}
