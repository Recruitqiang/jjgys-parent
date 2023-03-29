package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabx;
import org.apache.ibatis.annotations.Mapper;
import org.springframework.stereotype.Repository;

import java.util.List;

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
public interface JjgFbgcJtaqssJabxMapper extends BaseMapper<JjgFbgcJtaqssJabx> {

    List<JjgFbgcJtaqssJabx> selectbxnfsxs(String proname, String htd, String fbgc);

    List<JjgFbgcJtaqssJabx> selecthxnfsxs(String proname, String htd, String fbgc);
}
