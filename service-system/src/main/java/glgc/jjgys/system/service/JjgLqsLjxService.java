package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgLjx;
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
public interface JjgLqsLjxService extends IService<JjgLjx> {

    void export(HttpServletResponse response);

    void importLjx(MultipartFile file, String proname);
}
