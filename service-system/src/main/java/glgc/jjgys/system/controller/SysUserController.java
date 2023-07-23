package glgc.jjgys.system.controller;


import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.MD5;
import glgc.jjgys.model.system.SysDept;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.model.vo.SysUserQueryVo;
import glgc.jjgys.system.service.SysDepartService;
import glgc.jjgys.system.service.SysUserService;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.util.HashMap;
import java.util.Map;

/**
 * <p>
 * 用户表 前端控制器
 * </p>
 */
@Api(tags = "用户管理接口")
@RestController
@RequestMapping("/admin/system/sysUser")
public class SysUserController {

    @Autowired
    private SysUserService sysUserService;
    @Autowired
    private SysDepartService sysDepartService;

    @ApiOperation("校验旧密码")
    @GetMapping("checkSysUserPwd/{oldpw}/{username}")
    public Result checkSysUserPwd(@PathVariable String oldpw,
                                  @PathVariable String username) {
        boolean pwd = sysUserService.checkSysUserPwd(oldpw,username);
        if (pwd){
            return Result.ok(null).message("旧密码校验成功");
        } else {
            return Result.fail(null).message("旧密码校验失败");
        }

    }

    @ApiOperation("用户修改密码")
    @PostMapping("updatepwd/{username}/{newpass}")
    public Result updatepwd(@PathVariable String username,
                            @PathVariable String newpass) {
        boolean is_Success = sysUserService.updatepwd(username,newpass);
        if(is_Success) {
            return Result.ok(null);
        } else {
            return Result.fail(null);
        }
    }

    @ApiOperation("根据用户名查询")
    @GetMapping("getUserByName/{username}")
    public Result getUserByName(@PathVariable String username) {
        SysUser user = sysUserService.getUserInfoByUserName(username);
        SysDept sysDept = sysDepartService.getById(user.getDeptId());
        String deptId = sysDept.getId().toString();
        String  departNames = sysDepartService.selectParentDeptName(deptId);
        Map<String,Object> map=new HashMap<>();
        map.put("departname",departNames);
        user.setParam(map);
        return Result.ok(user);
    }

    @ApiOperation("用户列表")
    @GetMapping("/{page}/{limit}")
    public Result list(@PathVariable Long page,
                       @PathVariable Long limit,
                       SysUserQueryVo sysUserQueryVo) {
        //创建page对象
        Page<SysUser> pageParam = new Page<>(page,limit);
        //调用service方法
        IPage<SysUser> pageModel = sysUserService.selectPage(pageParam,sysUserQueryVo);
        return Result.ok(pageModel);
    }

    @ApiOperation("添加用户")
    @PostMapping("save")
    public Result save(@RequestBody SysUser user) {
        //把输入密码进行加密 MD5
        String encrypt = MD5.encrypt(user.getPassword());
        user.setPassword(encrypt);
        boolean is_Success = sysUserService.save(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }



    @ApiOperation("根据id查询")
    @GetMapping("getUser/{id}")
    public Result getUser(@PathVariable String id) {
        SysUser user = sysUserService.getById(id);
        return Result.ok(user);
    }

    @ApiOperation("修改用户")
    @PostMapping("update")
    public Result update(@RequestBody SysUser user) {
        boolean is_Success = sysUserService.updateById(user);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

    @ApiOperation("删除用户")
    @DeleteMapping("remove/{id}")
    public Result remove(@PathVariable String id) {
        boolean is_Success = sysUserService.removeById(id);
        if(is_Success) {
            return Result.ok();
        } else {
            return Result.fail();
        }
    }
}

