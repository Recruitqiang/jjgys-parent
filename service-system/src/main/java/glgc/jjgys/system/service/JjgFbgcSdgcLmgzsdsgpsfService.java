package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcSdgcLmgzsdsgpsf;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
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
 * @since 2023-05-03
 */
public interface JjgFbgcSdgcLmgzsdsgpsfService extends IService<JjgFbgcSdgcLmgzsdsgpsf> {

    void generateJdb(CommonInfoVo commonInfoVo) throws IOException;

    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;

    void exportlmgzsdsgpsf(HttpServletResponse response);

    void importlmgzsdsgpsf(MultipartFile file, CommonInfoVo commonInfoVo);

    List<Map<String, Object>> selectsdmc(String proname, String htd, String fbgc);
}
