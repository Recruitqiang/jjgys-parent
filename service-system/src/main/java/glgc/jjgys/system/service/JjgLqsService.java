package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.projectvo.lqs.QlVo;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.UnsupportedEncodingException;

public interface JjgLqsService {

    void exportLqs(HttpServletResponse response,String projectname);

    void importlqs(MultipartFile file, String projectname) throws IOException;
}
