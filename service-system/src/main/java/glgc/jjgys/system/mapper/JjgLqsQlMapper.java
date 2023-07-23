package glgc.jjgys.system.mapper;

import glgc.jjgys.model.project.JjgLqsQl;
import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgLqsSd;
import org.apache.ibatis.annotations.Mapper;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Map;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2022-12-14
 */
@Repository
@Mapper
public interface JjgLqsQlMapper extends BaseMapper<JjgLqsQl> {

    List<JjgLqsQl> selectqlzf(String proname, String htdzhq, String htdzhz, String lf);

    List<JjgLqsQl> selectqlyf(String proname, String htdzhq, String htdzhz, String lf);

    List<JjgLqsQl> selectqlList(String proname, String zhq, String zhz, String bz,String wz,String zdlf);

    List<JjgLqsQl> getqlName(String proname, String htd);

}
