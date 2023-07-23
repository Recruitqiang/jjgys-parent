package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcLjgcXqjgcc;
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
public interface JjgFbgcLjgcXqjgccMapper extends BaseMapper<JjgFbgcLjgcXqjgcc> {

    List<Map<String, Object>> selectyxps(String proname, String htd);
}
