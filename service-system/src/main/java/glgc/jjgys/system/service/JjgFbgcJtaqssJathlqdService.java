package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcJtaqssJathlqd;
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
 * @since 2023-03-01
 */
public interface JjgFbgcJtaqssJathlqdService extends IService<JjgFbgcJtaqssJathlqd> {

    void importjathlqd(MultipartFile file, CommonInfoVo commonInfoVo);

    void exportjathlqd(HttpServletResponse response);

    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;

    void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException;
}
