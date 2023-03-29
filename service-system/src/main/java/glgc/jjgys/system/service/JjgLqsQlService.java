package glgc.jjgys.system.service;


import glgc.jjgys.model.project.JjgLqsQl;
import com.baomidou.mybatisplus.extension.service.IService;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

/**
 * <p>
 *  服务类
 * </p>
 *
 * @author wq
 * @since 2022-12-14
 */
public interface JjgLqsQlService extends IService<JjgLqsQl> {

    void importQL(MultipartFile file,String proname);

    void exportQL(HttpServletResponse response,String proname);
}
