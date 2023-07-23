package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcSdgcSdhntlmhdzxf;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-05-20
 */
@Mapper
public interface JjgFbgcSdgcSdhntlmhdzxfMapper extends BaseMapper<JjgFbgcSdgcSdhntlmhdzxf> {

    List<Map<String, Object>> selectsdmc(String proname, String htd);
}
