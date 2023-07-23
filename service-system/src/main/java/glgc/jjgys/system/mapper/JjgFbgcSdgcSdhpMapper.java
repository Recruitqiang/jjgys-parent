package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcSdgcSdhp;
import glgc.jjgys.system.service.JjgFbgcSdgcSdhpService;
import org.apache.ibatis.annotations.Mapper;
import org.springframework.beans.factory.annotation.Autowired;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-03-27
 */
@Mapper
public interface JjgFbgcSdgcSdhpMapper extends BaseMapper<JjgFbgcSdgcSdhp> {

    List<Map<String,Object>> selectsdmc(String proname, String htd, String fbgc);

    List<Map<String,Object>> selectzh(String proname, String htd, String fbgc,String sdmc);


    List<Map<String, Object>> selectsdmc1(String proname, String htd);

    List<Map<String, Object>> selectzh1(String proname, String htd, String sdmc);
}
