package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgFbgcLjgcPspqhd;
import org.apache.ibatis.annotations.Mapper;
import org.springframework.stereotype.Repository;

import javax.annotation.Resource;
import java.util.List;

/**
 * <p>
 *  Mapper 接口
 * </p>
 *
 * @author wq
 * @since 2023-02-16
 */
@Repository
@Mapper
public interface JjgFbgcLjgcPspqhdMapper extends BaseMapper<JjgFbgcLjgcPspqhd> {

    List<String> selectyxps(String proname, String htd);

}
