package glgc.jjgys.system.mapper;

import glgc.jjgys.model.project.JjgZdhGzsd;
import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import org.apache.ibatis.annotations.Mapper;

import java.util.Collection;
import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-06-15
 */
@Mapper
public interface JjgZdhGzsdMapper extends BaseMapper<JjgZdhGzsd> {

    List<Map<String, Object>> selectlx(String proname, String htd);

    int selectcdnum(String proname, String htd, String zx);

    List<Map<String, Object>> selectzfList(String proname, String htd, String zx);

    List<Map<String, Object>> selectyfList(String proname, String htd, String zx);

    Collection<? extends Map<String, Object>> selectSdZfData(String proname, String htd, String zx,String zhq, String zhz);

    Collection<? extends Map<String, Object>> selectSdyfData(String proname, String htd, String zx,String zhq, String zhz);

    Collection<? extends Map<String, Object>> selectQlZfData(String proname, String htd, String zx, String zhq, String zhz);

    Collection<? extends Map<String, Object>> selectQlYfData(String proname, String htd, String zx, String zhq, String zhz);

    Collection<? extends Map<String, Object>> selectsdgzsd(String proname,String bz,String lf, String sdzhq, String sdzhz, String zx, String zhq1, String zhz1);

    Collection<? extends Map<String, Object>> selectqlgzsd(String proname, String bz, String lf, String qlzhq, String qlzhz, String zx, String zhq1, String zhz1);
}
