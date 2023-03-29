package glgc.jjgys.system.mapper;

import glgc.jjgys.model.system.SysOperLog;
import com.baomidou.mybatisplus.core.mapper.BaseMapper;
import org.apache.ibatis.annotations.Mapper;
import org.springframework.stereotype.Repository;

/**
 * <p>
 * 用户表 Mapper 接口
 * </p>
 */
@Repository
@Mapper
public interface OperLogMapper extends BaseMapper<SysOperLog> {

}
