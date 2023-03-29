package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.system.SysPost;
import glgc.jjgys.model.vo.SysPostQueryVo;

public interface SysPostService extends IService<SysPost> {
    IPage<SysPost> selectPage(long page, long limit, SysPostQueryVo sysPostQueryVo);
}
