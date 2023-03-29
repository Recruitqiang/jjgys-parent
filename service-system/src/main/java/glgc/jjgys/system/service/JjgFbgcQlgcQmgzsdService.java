package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcQlgcQmgzsd;
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
 * @since 2023-03-20
 */
public interface JjgFbgcQlgcQmgzsdService extends IService<JjgFbgcQlgcQmgzsd> {

    void generateJdb(CommonInfoVo commonInfoVo) throws IOException;

    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo);

    void exportqmgzsd(HttpServletResponse response);

    void importqmgzsd(MultipartFile file, CommonInfoVo commonInfoVo);
}
