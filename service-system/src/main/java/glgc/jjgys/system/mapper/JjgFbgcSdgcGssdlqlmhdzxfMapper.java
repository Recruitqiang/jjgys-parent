package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcSdgcGssdlqlmhdzxf;
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
 * @since 2023-05-20
 */
@Mapper
public interface JjgFbgcSdgcGssdlqlmhdzxfMapper extends BaseMapper<JjgFbgcSdgcGssdlqlmhdzxf> {

    List<Map<String, Object>> selectsdmc(String proname, String htd, String fbgc);

    List<JjgFbgcSdgcGssdlqlmhdzxf> selectzxzf(String proname, String htd, String fbgc, String sdmc);

    List<JjgFbgcSdgcGssdlqlmhdzxf> selectzxyf(String proname, String htd, String fbgc, String sdmc);

    List<JjgFbgcSdgcGssdlqlmhdzxf> selectsdzf(String proname, String htd, String fbgc, String sdmc);

    List<JjgFbgcSdgcGssdlqlmhdzxf> selectsdyf(String proname, String htd, String fbgc, String sdmc);

    List<JjgFbgcSdgcGssdlqlmhdzxf> selectqlzf(String proname, String htd, String fbgc, String sdmc);

    List<JjgFbgcSdgcGssdlqlmhdzxf> selectqlyf(String proname, String htd, String fbgc, String sdmc);

    List<JjgFbgcSdgcGssdlqlmhdzxf> selectzd(String proname, String htd, String fbgc, String sdmc);

    List<Map<String, Object>> selectsdmc1(String proname, String htd);
}
