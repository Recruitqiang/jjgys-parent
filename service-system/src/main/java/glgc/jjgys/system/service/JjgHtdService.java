package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgHtd;

import java.util.List;

public interface JjgHtdService extends IService<JjgHtd> {

    boolean addhtd(JjgHtd jjgHtd);

    JjgHtd selectlx(String proname, String htd);

    JjgHtd selectInfo(String proname, String htd);

    String getAllzh(String proname);

    List<JjgHtd> gethtd(String proname);
}
