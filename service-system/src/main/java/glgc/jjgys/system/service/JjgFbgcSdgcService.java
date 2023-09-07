package glgc.jjgys.system.service;

import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

/**
 * <p>
 *  服务类
 * </p>
 *
 * @author lhy
 * @since 2023-07-12
 */
public interface JjgFbgcSdgcService {
    void exportsdgc(HttpServletResponse response,String workpath) ;
    void importsdgc( CommonInfoVo commonInfoVo,String workpath) ;
    void generateJdb(CommonInfoVo commonInfoVo) throws IOException;
    void download(HttpServletResponse response,String filename,String proname,String htd,String workpath);
}
