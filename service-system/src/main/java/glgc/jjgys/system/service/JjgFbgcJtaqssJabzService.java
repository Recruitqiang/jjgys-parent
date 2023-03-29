package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabz;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.text.ParseException;

/**
 * <p>
 *  服务类
 * </p>
 *
 * @author wq
 * @since 2023-03-01
 */
public interface JjgFbgcJtaqssJabzService extends IService<JjgFbgcJtaqssJabz> {

    void importjabz(MultipartFile file, CommonInfoVo commonInfoVo);

    void exportjabz(HttpServletResponse response);

    void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException;
}
