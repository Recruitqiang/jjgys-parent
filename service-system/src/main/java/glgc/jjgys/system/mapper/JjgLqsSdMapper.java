package glgc.jjgys.system.mapper;

import glgc.jjgys.model.project.JjgLjxSd;
import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgLqsSd;
import org.apache.ibatis.annotations.Mapper;
import org.springframework.stereotype.Repository;

import java.util.List;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-01-09
 */
@Repository
@Mapper
public interface JjgLqsSdMapper extends BaseMapper<JjgLqsSd> {

    List<JjgLqsSd> selectsdzf(String proname, String htdzhq, String htdzhz, String lf);

    List<JjgLqsSd> selectsdyf(String proname, String htdzhq, String htdzhz, String lf);

    List<JjgLqsSd> selectsdList(String proname, String zhq, String zhz, String bz,String wz,String zdlf);

    List<JjgLqsSd> selectsd(String proname, String htdzhq, String htdzhz);

    List<JjgLqsSd> getsdName(String proname, String htd);

}
