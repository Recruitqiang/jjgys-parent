package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcLjgcHdgqd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * <p>
 *  服务类
 * </p>
 *
 * @author wq
 * @since 2023-01-14
 */
public interface JjgFbgcLjgcHdgqdService extends IService<JjgFbgcLjgcHdgqd> {

    void exporthdgqd(HttpServletResponse response);

    void importhdgqd(MultipartFile file, CommonInfoVo commonInfoVo);

    void generateJdb(CommonInfoVo commonInfoVo) throws Exception;

    List<Map<String,Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;

    void calculateSheet(XSSFSheet sheet);

    List<String> selectsjqd(String proname, String htd);

    Map<String, Object> selectchs(String proname, String htd);

    int selectnum(String proname, String htd);
}
