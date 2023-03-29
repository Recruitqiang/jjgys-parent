package glgc.jjgys.system.service.impl;

import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.model.system.SysPost;
import glgc.jjgys.model.vo.SysPostQueryVo;
import glgc.jjgys.system.mapper.SysPostMapper;
import glgc.jjgys.system.service.SysPostService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;

@Service
public class SysPostServiceImpl extends ServiceImpl<SysPostMapper, SysPost> implements SysPostService {

    @Autowired
    private SysPostMapper sysPostMapper;

    @Override
    public IPage<SysPost> selectPage(long page, long limit, SysPostQueryVo sysPostQueryVo) {
        Page<SysPost> pageParam = new Page<>(page,limit);
        //获取条件值
        String postName = sysPostQueryVo.getName();//岗位名称
        //封装参数
        QueryWrapper<SysPost> wrapper = new QueryWrapper<>();
        if(!StringUtils.isEmpty(postName)) {
            wrapper.like("name",postName);
        }
        wrapper.orderByDesc("create_time");
        //调用mapper方法实现分页条件查询
        IPage<SysPost> sysPostPage = sysPostMapper.selectPage(pageParam, wrapper);
        return sysPostPage;
    }
}
