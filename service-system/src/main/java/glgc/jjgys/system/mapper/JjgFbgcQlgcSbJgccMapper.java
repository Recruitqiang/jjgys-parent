package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcQlgcSbJgcc;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-03-20
 */
@Mapper
public interface JjgFbgcQlgcSbJgccMapper extends BaseMapper<JjgFbgcQlgcSbJgcc> {

    List<Map<String,Object>> selectnum(String proname, String htd, String fbgc);

}
