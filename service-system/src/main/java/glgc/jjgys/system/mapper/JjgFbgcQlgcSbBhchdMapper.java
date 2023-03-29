package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcQlgcSbBhchd;
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
public interface JjgFbgcQlgcSbBhchdMapper extends BaseMapper<JjgFbgcQlgcSbBhchd> {

    List<Map<String,Object>> selectnum(String proname, String htd, String fbgc);

}
