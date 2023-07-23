package glgc.jjgys.system.service;

import glgc.jjgys.model.project.JjgZdhMcxs;
import com.baomidou.mybatisplus.extension.service.IService;
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
 * @since 2023-06-06
 */
public interface JjgZdhMcxsService extends IService<JjgZdhMcxs> {

    void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException;

    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;

    void exportmcxs(HttpServletResponse response, String cdsl) throws IOException;

    void importmcxs(MultipartFile file, CommonInfoVo commonInfoVo) throws IOException;
}
