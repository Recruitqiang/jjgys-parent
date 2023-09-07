package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcLmgcLmwcLcf;
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
 * @since 2023-04-18
 */
public interface JjgFbgcLmgcLmwcLcfService extends IService<JjgFbgcLmgcLmwcLcf> {

    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;

    void generateJdb(CommonInfoVo commonInfoVo) throws IOException;

    void exportlmwclcf(HttpServletResponse response);

    void importlmwclcf(MultipartFile file, CommonInfoVo commonInfoVo);

    int selectnum(String proname, String htd);
}
