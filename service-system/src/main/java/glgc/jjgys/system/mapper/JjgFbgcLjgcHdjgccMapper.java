package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcLjgcHdjgcc;
import org.apache.ibatis.annotations.Mapper;
import org.springframework.stereotype.Repository;

import java.util.List;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-02-08
 */
@Repository
@Mapper
public interface JjgFbgcLjgcHdjgccMapper extends BaseMapper<JjgFbgcLjgcHdjgcc> {

    List<String> selectyxps(String proname, String htd);
}
