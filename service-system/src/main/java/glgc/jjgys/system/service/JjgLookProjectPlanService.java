package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.project.JjgPlaninfo;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
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
public interface JjgLookProjectPlanService extends IService<JjgPlaninfo> {

    List<Map<String,Object>> lookplan(CommonInfoVo commonInfoVo);

    void exportxmjd(HttpServletResponse response, String projectname);

    void importxmjd(MultipartFile file, String proname);
}
