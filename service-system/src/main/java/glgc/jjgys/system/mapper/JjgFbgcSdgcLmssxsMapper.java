package glgc.jjgys.system.mapper;

import glgc.jjgys.model.project.JjgFbgcLmgcLmssxs;
import glgc.jjgys.model.project.JjgFbgcSdgcLmssxs;
import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-05-04
 */
@Mapper
public interface JjgFbgcSdgcLmssxsMapper extends BaseMapper<JjgFbgcSdgcLmssxs> {

    List<Map<String, Object>> selectsdmc(String proname, String htd, String fbgc);

    List<JjgFbgcLmgcLmssxs> selectsdzfdata(String proname, String htd, String fbgc, String sdmc);

    List<JjgFbgcLmgcLmssxs> selectsdyfdata(String proname, String htd, String fbgc, String sdmc);

    List<Map<String, Object>> selectsdmc1(String proname, String htd);
}
