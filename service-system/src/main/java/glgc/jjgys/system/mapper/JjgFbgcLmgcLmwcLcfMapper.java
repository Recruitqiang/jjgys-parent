package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcLmgcLmwcLcf;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-04-18
 */
@Mapper
public interface JjgFbgcLmgcLmwcLcfMapper extends BaseMapper<JjgFbgcLmgcLmwcLcf> {

    List<Map<String, Object>> selectwdata(String proname, String htd, String fbgc);

    int selectnum(String proname, String htd);
}
