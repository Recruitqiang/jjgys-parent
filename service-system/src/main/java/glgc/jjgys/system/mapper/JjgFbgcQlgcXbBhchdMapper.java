package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcQlgcXbBhchd;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-03-22
 */
@Mapper
public interface JjgFbgcQlgcXbBhchdMapper extends BaseMapper<JjgFbgcQlgcXbBhchd> {
    List<Map<String,Object>> selectnum(String proname, String htd, String fbgc);

}
