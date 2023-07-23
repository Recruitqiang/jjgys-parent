package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcLjgcZdgqd;
import org.apache.ibatis.annotations.Mapper;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-02-15
 */
@Repository
@Mapper
public interface JjgFbgcLjgcZdgqdMapper extends BaseMapper<JjgFbgcLjgcZdgqd> {

    List<Map<String, Object>> selectsjqd(String proname, String htd);

    Map<String, Object> selectchs(String proname, String htd);
}
