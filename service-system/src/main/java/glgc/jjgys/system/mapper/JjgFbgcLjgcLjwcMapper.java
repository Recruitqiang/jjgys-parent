package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcLjgcLjwc;
import org.apache.ibatis.annotations.Mapper;
import org.springframework.stereotype.Repository;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-02-27
 */
@Repository
@Mapper
public interface JjgFbgcLjgcLjwcMapper extends BaseMapper<JjgFbgcLjgcLjwc> {

    int selectnum(String proname, String htd);
}
