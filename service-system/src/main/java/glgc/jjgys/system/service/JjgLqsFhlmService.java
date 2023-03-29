package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgLqsFhlm;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;

/**
 * <p>
 *  服务类
 * </p>
 *
 * @author wq
 * @since 2023-01-10
 */
public interface JjgLqsFhlmService extends IService<JjgLqsFhlm> {

    void export(HttpServletResponse response);

    void importFhlm(MultipartFile file, String proname);
}
