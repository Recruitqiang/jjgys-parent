package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgLjxhntlm;
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
public interface JjgLqsLjxhntlmService extends IService<JjgLjxhntlm> {

    void exportLjxhntlm(HttpServletResponse response);

    void importLjxhntlm(MultipartFile file, String proname);
}
