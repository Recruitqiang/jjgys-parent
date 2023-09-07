package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgFbgcLjgcZdgqd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
 * @since 2023-02-15
 */
public interface JjgFbgcLjgcZdgqdService extends IService<JjgFbgcLjgcZdgqd> {

    void generateJdb(CommonInfoVo commonInfoVo) throws IOException, ParseException;

    List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) throws IOException;

    void exportzdgqd(HttpServletResponse response);

    void importzdgqd(MultipartFile file,CommonInfoVo commonInfoVo);

    double getjdxzz(XSSFSheet sheet, double T, String U);

    double getjzmxzz(XSSFSheet sheet, double W, String X);

    double getthsdhsz(XSSFSheet sheet, double Z, String AA);

    String getBridgeName(String name);

    void calculateSheet(XSSFWorkbook wb);

    List<Map<String, Object>> selectsjqd(String proname, String htd);

    Map<String, Object> selectchs(String proname, String htd);

    int selectnum(String proname, String htd);
}
