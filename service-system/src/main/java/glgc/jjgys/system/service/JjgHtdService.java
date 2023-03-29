package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgHtd;

import java.util.List;

public interface JjgHtdService extends IService<JjgHtd> {

    boolean addhtd(JjgHtd jjgHtd);
}
