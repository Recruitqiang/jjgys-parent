package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcLmgcLqlmysd;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-04-10
 */
@Mapper
public interface JjgFbgcLmgcLqlmysdMapper extends BaseMapper<JjgFbgcLmgcLqlmysd> {

    Map<String, String> selectsffl(String proname, String htd, String fbgc);

    List<JjgFbgcLmgcLqlmysd> selectzd(String proname, String htd, String fbgc);

    List<JjgFbgcLmgcLqlmysd> selectljx(String proname, String htd, String fbgc);

    List<JjgFbgcLmgcLqlmysd> selectljxsd(String proname, String htd, String fbgc);

    //沥青路面压实度主线左幅上面层
    List<JjgFbgcLmgcLqlmysd> selectzxsmc(String proname, String htd, String fbgc, String sffl, String cw);

    //沥青路面压实度主线左幅中下面层
    List<JjgFbgcLmgcLqlmysd> selectzxzxmc(String proname, String htd, String fbgc, String sffl);

    List<JjgFbgcLmgcLqlmysd> selectzxyfsmc(String proname, String htd, String fbgc, String sffl, String cw);

    List<JjgFbgcLmgcLqlmysd> selectzxyfzxmc(String proname, String htd, String fbgc, String sffl);

    List<JjgFbgcLmgcLqlmysd> selectsdzfsmc(String proname, String htd, String fbgc, String sffl, String cw);

    List<JjgFbgcLmgcLqlmysd> selectsdyfsmc(String proname, String htd, String fbgc, String sffl, String cw);

    List<JjgFbgcLmgcLqlmysd> selectsdzfzxmc(String proname, String htd, String fbgc, String sffl);

    List<JjgFbgcLmgcLqlmysd> selectsdyfzxmc(String proname, String htd, String fbgc, String sffl);

    List<JjgFbgcLmgcLqlmysd> selectzdsmc(String proname, String htd, String fbgc);

    List<JjgFbgcLmgcLqlmysd> selectzdsxmc(String proname, String htd, String fbgc);

    List<JjgFbgcLmgcLqlmysd> selectljxsmc(String proname, String htd, String fbgc);

    List<JjgFbgcLmgcLqlmysd> selectljxzxmc(String proname, String htd, String fbgc);

    List<JjgFbgcLmgcLqlmysd> selectljxsdsmc(String proname, String htd, String fbgc);

    List<JjgFbgcLmgcLqlmysd> selectljxsdzxmc(String proname, String htd, String fbgc);
}
