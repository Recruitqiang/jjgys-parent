package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcLmgcHntlmqd;
import org.apache.ibatis.annotations.Mapper;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-04-20
 */
@Mapper
public interface JjgFbgcLmgcHntlmqdMapper extends BaseMapper<JjgFbgcLmgcHntlmqd> {

    int selectnum(String proname, String htd);
}
