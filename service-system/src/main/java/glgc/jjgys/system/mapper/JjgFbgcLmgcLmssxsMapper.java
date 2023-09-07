package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcLmgcLmssxs;
import org.apache.ibatis.annotations.Mapper;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-04-19
 */
@Mapper
public interface JjgFbgcLmgcLmssxsMapper extends BaseMapper<JjgFbgcLmgcLmssxs> {

    int selectnum(String proname, String htd);
}
