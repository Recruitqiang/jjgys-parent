package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.system.Project;


public interface ProjectService extends IService<Project> {

    void addOtherInfo(String proName);

    Integer getlevel(String proname);
}
