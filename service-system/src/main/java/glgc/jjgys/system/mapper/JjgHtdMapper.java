package glgc.jjgys.system.mapper;

import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import glgc.jjgys.model.project.JjgHtd;
import org.apache.ibatis.annotations.Mapper;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Map;

@Repository
@Mapper
public interface JjgHtdMapper extends BaseMapper<JjgHtd> {
    String selectscDataId(String proname);

    List<JjgHtd> selecthtd(String proname);

}
