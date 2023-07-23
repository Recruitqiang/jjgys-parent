package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcSdgcCqhd;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-03-27
 */
@Mapper
public interface JjgFbgcSdgcCqhdMapper extends BaseMapper<JjgFbgcSdgcCqhd> {

    List<Map<String,Object>> selectsdmc(String proname, String htd, String fbgc);

    List<Map<String,Object>> selectcd(String proname, String htd, String fbgc,String sdmc);

    List<Map<String,Object>> selectwz(String proname, String htd, String fbgc,String sdmc);

    List<Map<String, Object>> selectsdmc2(String proname, String htd);

    List<Map<String, Object>> selectcd2(String proname, String htd, String sdmc);

    List<Map<String, Object>> selectwz2(String proname, String htd, String sdmc);

    List<Map<String, Object>> selectsjhd(String proname, String htd,String s);

    List<Map<String, Object>> getds(String proname, String htd, String sdmc);
}
