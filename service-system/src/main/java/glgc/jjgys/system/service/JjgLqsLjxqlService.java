package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgLjxQl;
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
public interface JjgLqsLjxqlService extends IService<JjgLjxQl> {

    void exportLjxQL(HttpServletResponse response);

    void importLjxQL(MultipartFile file, String proname);
}
