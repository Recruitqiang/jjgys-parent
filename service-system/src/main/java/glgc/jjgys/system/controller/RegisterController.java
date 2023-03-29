package glgc.jjgys.system.controller;

import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.MD5;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.model.system.SysUserRole;
import glgc.jjgys.system.service.RegisterService;
import glgc.jjgys.system.service.SysUserService;
import io.swagger.annotations.Api;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.util.UUID;

@Api(tags = "用户注册接口")
@RestController
@RequestMapping(value = "/user")
@CrossOrigin//解决跨域
public class RegisterController {

    @Autowired
    private RegisterService registerService;

    @Autowired
    private SysUserService sysUserService;


    @PostMapping("/register")
    public Result addUser(@RequestBody SysUser sysUser) {
        boolean isExist = registerService.userNameIsExist(sysUser.getUsername());
        String encrypt = MD5.encrypt(sysUser.getPassword());
        sysUser.setPassword(encrypt);
        boolean is_Success = sysUserService.save(sysUser);

        if (!isExist && is_Success){
            return Result.ok().message("注册成功");
        }else {
            return Result.fail().message("用户名重复，请重新注册");
        }

    }

}
