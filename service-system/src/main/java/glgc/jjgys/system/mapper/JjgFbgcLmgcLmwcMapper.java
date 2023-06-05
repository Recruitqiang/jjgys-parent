package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcLmgcLmwc;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-04-16
 */
@Mapper
public interface JjgFbgcLmgcLmwcMapper extends BaseMapper<JjgFbgcLmgcLmwc> {

    List<Map<String, Object>> selectwdata(String proname, String htd, String fbgc);
}
