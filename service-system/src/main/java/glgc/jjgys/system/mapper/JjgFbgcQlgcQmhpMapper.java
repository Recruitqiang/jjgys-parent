package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcQlgcQmhp;
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
public interface JjgFbgcQlgcQmhpMapper extends BaseMapper<JjgFbgcQlgcQmhp> {

    List<Map<String,Object>> selectqlmc(String proname, String htd, String fbgc);

    List<Map<String,Object>> selectzh(String proname, String htd, String fbgc,String qlmc);

}
