package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcLjgcXqjgcc;
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
 * @since 2023-02-15
 */
public interface JjgFbgcLjgcXqjgccService extends IService<JjgFbgcLjgcXqjgcc> {

    void generateJdb(CommonInfoVo commonInfoVo) throws Exception;

    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;

    void exportxqjgcc(HttpServletResponse response);

    void importxqjgcc(MultipartFile file,CommonInfoVo commonInfoVo);

    List<Map<String, Object>> selectyxps(String proname, String htd);
}
