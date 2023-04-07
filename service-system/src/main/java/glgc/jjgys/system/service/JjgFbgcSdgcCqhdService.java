package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcSdgcCqhd;
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
 * @since 2023-03-27
 */
public interface JjgFbgcSdgcCqhdService extends IService<JjgFbgcSdgcCqhd> {

    void generateJdb(CommonInfoVo commonInfoVo) throws IOException;

    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;

    void exportsdcqhd(HttpServletResponse response);

    void importsdcqhd(MultipartFile file, CommonInfoVo commonInfoVo);

    List<Map<String, Object>> selectsdmc(String proname, String htd, String fbgc);
}
