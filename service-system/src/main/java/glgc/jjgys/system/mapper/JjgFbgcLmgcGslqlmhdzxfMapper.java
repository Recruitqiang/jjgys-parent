package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcLmgcGslqlmhdzxf;
import glgc.jjgys.model.project.JjgFbgcSdgcGssdlqlmhdzxf;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-04-25
 */
@Mapper
public interface JjgFbgcLmgcGslqlmhdzxfMapper extends BaseMapper<JjgFbgcLmgcGslqlmhdzxf> {

    List<JjgFbgcLmgcGslqlmhdzxf> selectzxzf(String proname, String htd, String fbgc);

    List<JjgFbgcLmgcGslqlmhdzxf> selectzxyf(String proname, String htd, String fbgc);

    List<JjgFbgcLmgcGslqlmhdzxf> selectsdzf(String proname, String htd, String fbgc);

    List<JjgFbgcLmgcGslqlmhdzxf> selectsdyf(String proname, String htd, String fbgc);

    List<JjgFbgcLmgcGslqlmhdzxf> selectqlzf(String proname, String htd, String fbgc);

    List<JjgFbgcLmgcGslqlmhdzxf> selectqlyf(String proname, String htd, String fbgc);

    List<JjgFbgcLmgcGslqlmhdzxf> selectzd(String proname, String htd, String fbgc);

    List<JjgFbgcLmgcGslqlmhdzxf> selectljx(String proname, String htd, String fbgc);

    List<JjgFbgcLmgcGslqlmhdzxf> selectljxq(String proname, String htd, String fbgc);

    List<JjgFbgcLmgcGslqlmhdzxf> selectljxsd(String proname, String htd, String fbgc);
}
