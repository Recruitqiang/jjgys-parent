package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcQlgcQmpzd;
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
 * @since 2023-03-20
 */
public interface JjgFbgcQlgcQmpzdService extends IService<JjgFbgcQlgcQmpzd> {

    void generateJdb(CommonInfoVo commonInfoVo) throws IOException;

    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;

    void exportQmpzd(HttpServletResponse response);

    void importQmpzd(MultipartFile file, CommonInfoVo commonInfoVo);

    List<Map<String, Object>> selectqlmc(String proname, String htd, String fbgc);

    List<Map<String, Object>> lookjg(CommonInfoVo commonInfoVo, String value) throws IOException;
}
