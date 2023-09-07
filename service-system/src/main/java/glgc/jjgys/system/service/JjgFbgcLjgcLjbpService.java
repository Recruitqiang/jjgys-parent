package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcLjgcLjbp;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.beans.IntrospectionException;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.InvocationTargetException;
import java.text.ParseException;
import java.util.List;
import java.util.Map;

/**
 * <p>
 *  服务类
 * </p>
 *
 * @author wq
 * @since 2023-02-21
 */
public interface JjgFbgcLjgcLjbpService extends IService<JjgFbgcLjgcLjbp> {

    void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException;

    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;

    void exportljbp(HttpServletResponse response) throws IOException;

    void importljbp(MultipartFile file, CommonInfoVo commonInfoVo) throws IOException, ParseException;

    List<String> selectyxps(String proname, String htd);

    int selectnum(String proname, String htd);
}
