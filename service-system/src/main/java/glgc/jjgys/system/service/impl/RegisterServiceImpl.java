package glgc.jjgys.system.service.impl;

import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.model.system.SysUser;
import glgc.jjgys.system.mapper.SysUserMapper;
import glgc.jjgys.system.service.RegisterService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class RegisterServiceImpl extends ServiceImpl<SysUserMapper, SysUser> implements RegisterService {

    @Autowired
    private SysUserMapper sysUserMapper;

    @Override
    public boolean userNameIsExist(String username) {
        List<String> checkName = sysUserMapper.selectUserName();
        boolean res = checkName.contains(username);
        if (!res){
            return false;
        }
        return true;
    }
}
