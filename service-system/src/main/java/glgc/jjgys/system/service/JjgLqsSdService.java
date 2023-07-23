package glgc.jjgys.system.service;

import glgc.jjgys.model.project.JjgLjxSd;
import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgLqsSd;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.util.List;

/**
 * <p>
 *  服务类
 * </p>
 *
 * @author wq
 * @since 2023-01-09
 */
public interface JjgLqsSdService extends IService<JjgLqsSd> {

    void exportSD(HttpServletResponse response,String projectname);

    void importSD(MultipartFile file, String proname);

    List<JjgLqsSd> getsdName(String proname, String htd);
}
