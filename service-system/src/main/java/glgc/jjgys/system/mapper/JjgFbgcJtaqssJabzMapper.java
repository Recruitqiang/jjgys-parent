package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabz;
import org.apache.ibatis.annotations.Mapper;
import org.springframework.stereotype.Repository;

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
public interface JjgFbgcJtaqssJabzMapper extends BaseMapper<JjgFbgcJtaqssJabz> {

    Map<String, Object> selectchs(String proname, String htd);

    int selectnum(String proname, String htd);
}
