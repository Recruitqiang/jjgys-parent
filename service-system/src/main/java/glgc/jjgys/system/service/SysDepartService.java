package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.system.SysDept;


public interface SysDepartService extends IService<SysDept> {

    String selectParentDeptName(String deptId);
}
