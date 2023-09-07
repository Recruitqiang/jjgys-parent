package glgc.jjgys.system.mapper;

import glgc.jjgys.model.project.JjgFbgcJtaqssJabxfhl;
import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import org.apache.ibatis.annotations.Mapper;
import org.springframework.stereotype.Repository;

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
public interface JjgFbgcJtaqssJabxfhlMapper extends BaseMapper<JjgFbgcJtaqssJabxfhl> {

    int selectnum(String proname, String htd);
}
