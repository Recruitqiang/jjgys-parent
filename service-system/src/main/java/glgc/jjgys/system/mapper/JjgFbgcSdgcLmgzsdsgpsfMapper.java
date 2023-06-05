package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcSdgcLmgzsdsgpsf;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-05-03
 */
@Mapper
public interface JjgFbgcSdgcLmgzsdsgpsfMapper extends BaseMapper<JjgFbgcSdgcLmgzsdsgpsf> {

    List<Map<String, Object>> selectsdmc(String proname, String htd, String fbgc);
}
