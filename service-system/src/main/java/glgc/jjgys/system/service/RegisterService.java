package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.system.SysUser;


public interface RegisterService extends IService<SysUser> {

    boolean userNameIsExist(String username);
}
