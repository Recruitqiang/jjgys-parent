package glgc.jjgys.system.service;


import glgc.jjgys.model.project.JjgLqsQl;
import com.baomidou.mybatisplus.extension.service.IService;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;
import java.util.Map;

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

    List<JjgLqsQl> getqlName(String proname, String htd);
}
