package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcLmgcLmhp;
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
 * @since 2023-04-22
 */
@Mapper
public interface JjgFbgcLmgcLmhpMapper extends BaseMapper<JjgFbgcLmgcLmhp> {

    List<Map<String, String>> selectlx(String proname, String htd);

    List<Map<String, Object>> selectljx(String proname, String htd, String fbgc);

    List<Map<String,String>> selectSdZfData(String proname, String htd, String fbgc, String zhq, String zhz);

    List<Map<String,String>> selectSdYfData(String proname, String htd, String fbgc, String zhq, String zhz);

    List<Map<String,String>> selectQlZfData(String proname, String htd, String fbgc, String zhq, String zhz);

    List<Map<String,String>> selectQlYfData(String proname, String htd, String fbgc, String zhq, String zhz);

    List<Map<String, String>> selectAllList(String proname, String htd, String fbgc);

    List<JjgFbgcLmgcLmhp> selecthpzd(Map<String, String> map);

    List<Map<String, String>> selecthpzdData(Map<String, String> mapzd);

    List<Map<String, String>> selecthpzdsdData(Map<String, String> mapzdql);

    List<JjgFbgcLmgcLmhp> selecthpljx(Map<String, String> map);

    int selectnum(String proname, String htd);

}
