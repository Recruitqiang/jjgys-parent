package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcSdgcSdlqlmysd;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-05-07
 */
@Mapper
public interface JjgFbgcSdgcSdlqlmysdMapper extends BaseMapper<JjgFbgcSdgcSdlqlmysd> {

    List<Map<String, Object>> selectsdmc(String proname, String htd, String fbgc);

    List<JjgFbgcSdgcSdlqlmysd> selectzxzf(String proname, String htd, String fbgc,String sdmc);

    List<JjgFbgcSdgcSdlqlmysd> selectzxyf(String proname, String htd, String fbgc,String sdmc);

    List<JjgFbgcSdgcSdlqlmysd> selectsdzf(String proname, String htd, String fbgc,String sdmc);

    List<JjgFbgcSdgcSdlqlmysd> selectsdyf(String proname, String htd, String fbgc,String sdmc);

    List<JjgFbgcSdgcSdlqlmysd> selectzd(String proname, String htd, String fbgc,String sdmc);

    List<JjgFbgcSdgcSdlqlmysd> selectljx(String proname, String htd, String fbgc,String sdmc);

    List<JjgFbgcSdgcSdlqlmysd> selectljxsd(String proname, String htd, String fbgc,String sdmc);
}
