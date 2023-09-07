package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcLjgcLjtsfysdHt;
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
 * @since 2023-02-21
 */
public interface JjgFbgcLjgcLjtsfysdHtService extends IService<JjgFbgcLjgcLjtsfysdHt> {

    void generateJdb(CommonInfoVo commonInfoVo) throws IOException;

    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;

    void exportysdht(HttpServletResponse response) throws IOException;

    void importysdht(MultipartFile file, CommonInfoVo commonInfoVo) throws IOException, ParseException;

    int selectnum(String proname, String htd);
}
