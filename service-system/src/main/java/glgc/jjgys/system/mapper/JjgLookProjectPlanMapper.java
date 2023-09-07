package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgPlaninfo;
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
 * @since 2023-03-01
 */
@Repository
@Mapper
public interface JjgLookProjectPlanMapper extends BaseMapper<JjgPlaninfo> {

    List<Map<String, Object>> selectplan(String proname, String htd);

}
