package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgZdhPzd;
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
 * @since 2023-06-17
 */
@Mapper
public interface JjgZdhPzdMapper extends BaseMapper<JjgZdhPzd> {

    List<Map<String, Object>> selectlx(String proname, String htd);

    int selectcdnum(String proname, String htd, String zx);

    List<Map<String,Object>> selectSdZfData(String proname, String htd, String zx,String zhq, String zhz,String result);

    List<Map<String,Object>>  selectSdyfData(String proname, String htd, String zx, String zhq, String zhz,String result);

    List<Map<String,Object>> selectQlZfData(String proname, String htd, String zx, String zhq, String zhz,String result);

    List<Map<String,Object>> selectQlYfData(String proname, String htd, String zx, String zhq, String zhz,String result);

    List<Map<String, Object>> selectzfList(String proname, String htd, String zx,String result);

    List<Map<String, Object>> selectyfList(String proname, String htd, String zx,String result);

    List<Map<String, Object>> selectSdData(String proname, String htd, String zx, String zhq, String zhz,String result);

    List<Map<String, Object>> selectsdpzd(String proname, String bz, String lf, String zx, String zhq1, String zhz1,String sdzhz);

    List<Map<String, Object>> selectqlpzd(String proname, String bz, String lf, String qlzhq, String qlzhz, String zx,String qlzhzj);

    Collection<? extends Map<String, Object>> selectsdpzd1(String proname, String bz, String lf, String zx, String zhq1, String zhz1, String sdzhq, String sdzhz);

    Collection<? extends Map<String, Object>> selectqlpzd1(String proname, String bz, String lf, String qlzhq, String qlzhz, String zx, String zhq1, String zhz1);

}
