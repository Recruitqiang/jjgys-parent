package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcLjgcLjtsfysdSl;
import org.apache.ibatis.annotations.Mapper;
import org.springframework.stereotype.Repository;

import java.util.List;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-02-21
 */
@Repository
@Mapper
public interface JjgFbgcLjgcLjtsfysdSlMapper extends BaseMapper<JjgFbgcLjgcLjtsfysdSl> {
    List<String> selectNums(String proname, String htd, String fbgc);

}
