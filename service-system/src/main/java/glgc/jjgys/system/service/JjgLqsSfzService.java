package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgSfz;
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
public interface JjgLqsSfzService extends IService<JjgSfz> {

    void exportSD(HttpServletResponse response);

    void importSD(MultipartFile file, String proname);
}
