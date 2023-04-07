package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcLjgcLjwc;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * <p>
 *  服务类
 * </p>
 *
 * @author wq
 * @since 2023-02-27
 */
public interface JjgFbgcLjgcLjwcService extends IService<JjgFbgcLjgcLjwc> {

    void generateJdb(CommonInfoVo commonInfoVo) throws IOException;

    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;

    void exportljwc(HttpServletResponse response);

    void importljwc(MultipartFile file, CommonInfoVo commonInfoVo);
}
