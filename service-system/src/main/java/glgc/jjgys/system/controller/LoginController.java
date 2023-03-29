package glgc.jjgys.system.controller;

import com.baomidou.mybatisplus.core.conditions.query.LambdaQueryWrapper;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.common.utils.MD5;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.model.vo.LoginVo;
import glgc.jjgys.system.service.SysUserService;
import io.swagger.annotations.Api;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletRequest;
import java.util.HashMap;
import java.util.Map;

@Api(tags = "用户登录接口")
@RestController
@RequestMapping("/admin/system/index")
public class LoginController {

    @Autowired
    private SysUserService sysUserService;

    @PostMapping("login")
    public Result login(@RequestBody LoginVo loginVo) {
        String username = loginVo.getUsername();
        //1. 校验用户是否存在
        LambdaQueryWrapper<SysUser> queryWrapper = new LambdaQueryWrapper<>();
        queryWrapper.eq(SysUser::getUsername,username);
        SysUser user = sysUserService.getOne(queryWrapper);
        if(user == null){
            return Result.fail(null).message("用户不存在，请注册!");
        }
        //2. 校验用户名或密码是否正确
        String userpassword = sysUserService.selectpwd(loginVo.getUsername());
        //把输入的密码加密后和数据库的密码做比对
        String encrypt = MD5.encrypt(loginVo.getPassword());
        if (!userpassword.equals(encrypt)) {
            return Result.fail(null).message("密码错误,请重新输入!");
        }
        //根据userid和username生成token字符串，通过map返回
        String token = JwtHelper.createToken(user.getId().toString(), user.getUsername());
        //map存放响应的数据
        Map<String,Object> map=new HashMap<>();
        map.put("token",token);
        return Result.ok(map);
    }

    @GetMapping("info")
    public Result info(HttpServletRequest request) {
        //获取请求头token字符串,因为把token放到请求头不存在跨域
        String token = request.getHeader("token");
        //从token字符串获取用户名称（id）
        String username = JwtHelper.getUsername(token);
        //根据用户名称获取用户信息
        Map<String,Object> map = sysUserService.getUserInfo(username);
        return Result.ok(map);
    }

}
