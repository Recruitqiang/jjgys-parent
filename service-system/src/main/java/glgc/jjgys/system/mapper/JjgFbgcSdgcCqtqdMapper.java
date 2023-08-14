package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcSdgcCqtqd;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-03-26
 */
@Mapper
public interface JjgFbgcSdgcCqtqdMapper extends BaseMapper<JjgFbgcSdgcCqtqd> {

    List<Map<String, Object>> getsdnum(String proname, String htd);
}
