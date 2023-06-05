package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcLmgcLmhp;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.text.ParseException;
import java.util.List;
import java.util.Map;

/**
 * <p>
 *  服务类
 * </p>
 *
 * @author wq
 * @since 2023-04-22
 */
public interface JjgFbgcLmgcLmhpService extends IService<JjgFbgcLmgcLmhp> {

    void generateJdb(CommonInfoVo commonInfoVo) throws Exception;

    List<Map<String, String>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;

    void exportlmhp(HttpServletResponse response);

    void importlmhp(MultipartFile file, CommonInfoVo commonInfoVo);

    List<Map<String, String>> selectmc(String proname, String htd, String fbgc);
}
