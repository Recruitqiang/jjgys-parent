package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcLjgcLjwcLcf;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

/**
 * <p>
 *  服务类
 * </p>
 *
 * @author wq
 * @since 2023-03-01
 */
public interface JjgFbgcLjgcLjwcLcfService extends IService<JjgFbgcLjgcLjwcLcf> {

    void generateJdb(CommonInfoVo commonInfoVo) throws IOException;

    void exportljwclcf(HttpServletResponse response);

    void importljwclcf(MultipartFile file, CommonInfoVo commonInfoVo);
}
