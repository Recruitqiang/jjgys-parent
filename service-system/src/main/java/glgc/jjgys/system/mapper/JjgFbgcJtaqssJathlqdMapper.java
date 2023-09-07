package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcJtaqssJathlqd;
import org.apache.ibatis.annotations.Mapper;

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

@Mapper
public interface JjgFbgcJtaqssJathlqdMapper extends BaseMapper<JjgFbgcJtaqssJathlqd> {

    List<String> selectsjqd(String proname, String htd);

    Map<String, Object> selectchs(String proname, String htd);

    int selectnum(String proname, String htd);
}
