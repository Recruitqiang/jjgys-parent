package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcLmgcTlmxlbgc;
import org.apache.ibatis.annotations.Mapper;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-04-21
 */
@Mapper
public interface JjgFbgcLmgcTlmxlbgcMapper extends BaseMapper<JjgFbgcLmgcTlmxlbgc> {

    int selectnum(String proname, String htd);
}
