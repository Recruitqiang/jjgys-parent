package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcSdgcCqtqd;
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
 * @since 2023-03-26
 */
public interface JjgFbgcSdgcCqtqdService extends IService<JjgFbgcSdgcCqtqd> {

    void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException;

    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;

    void exportsdcqtqd(HttpServletResponse response);

    void importsdcqtsd(MultipartFile file, CommonInfoVo commonInfoVo);

    List<Map<String, Object>> lookjg(CommonInfoVo commonInfoVo);

    List<Map<String, Object>> getsdnum(String proname, String htd);
}
