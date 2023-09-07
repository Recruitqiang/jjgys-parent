package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcLjgcLjcj;
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
 * @since 2023-02-21
 */
public interface JjgFbgcLjgcLjcjService extends IService<JjgFbgcLjgcLjcj> {

    void generateJdb(CommonInfoVo commonInfoVo) throws IOException;

    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;

    void exportljcj(HttpServletResponse response);

    void importljcj(MultipartFile file, CommonInfoVo commonInfoVo);

    List<String> selectyxps(String proname, String htd);

    int selectnum(String proname, String htd);
}
