package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabx;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
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
public interface JjgFbgcJtaqssJabxService extends IService<JjgFbgcJtaqssJabx> {

    void generateJdb(CommonInfoVo commonInfoVo) throws Exception;

    void exportjabx(HttpServletResponse response);

    void importjabx(MultipartFile file, CommonInfoVo commonInfoVo);

    void bxnfsxs(List<JjgFbgcJtaqssJabx> data) throws IOException, ParseException;

    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;
}
