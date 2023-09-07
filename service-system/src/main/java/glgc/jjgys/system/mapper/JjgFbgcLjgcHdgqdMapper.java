package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcLjgcHdgqd;
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
 * @since 2023-01-14
 */
@Repository
@Mapper
public interface JjgFbgcLjgcHdgqdMapper extends BaseMapper<JjgFbgcLjgcHdgqd> {

    List<String> selectzh(String proname,String htd);

    List<String> selectsjqd(String proname, String htd);

    Map<String, Object> selectchs(String proname, String htd);

    int selectnum(String proname, String htd);
}
