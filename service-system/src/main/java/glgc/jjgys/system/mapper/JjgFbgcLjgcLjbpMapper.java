package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcLjgcLjbp;
import org.apache.ibatis.annotations.Mapper;
import org.springframework.stereotype.Repository;

import javax.annotation.ManagedBean;
import java.util.List;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-02-21
 */
@Repository
@Mapper
public interface JjgFbgcLjgcLjbpMapper extends BaseMapper<JjgFbgcLjgcLjbp> {

    List<String> selectyxps(String proname, String htd);

    int selectnum(String proname, String htd);
}
