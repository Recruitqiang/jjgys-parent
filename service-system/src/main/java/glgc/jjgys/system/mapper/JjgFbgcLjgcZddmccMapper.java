package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcLjgcZddmcc;
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
 * @since 2023-02-16
 */
@Repository
@Mapper
public interface JjgFbgcLjgcZddmccMapper extends BaseMapper<JjgFbgcLjgcZddmcc> {

    List<Map<String,Object>> selectyxps(String proname, String htd);

    Map<String, Object> selectchs(String proname, String htd);
}
