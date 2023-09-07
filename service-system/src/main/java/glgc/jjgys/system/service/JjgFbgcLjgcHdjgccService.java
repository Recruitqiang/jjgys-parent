package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcLjgcHdjgcc;
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
 * @since 2023-02-08
 */
public interface JjgFbgcLjgcHdjgccService extends IService<JjgFbgcLjgcHdjgcc> {

    void generateJdb(CommonInfoVo commonInfoVo) throws Exception;

    void exporthdjgcc(HttpServletResponse response);

    void importhdjgcc(MultipartFile file,CommonInfoVo commonInfoVo);

    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;

    List<String> selectyxps(String proname, String htd);

    int selectnum(String proname, String htd);
}
