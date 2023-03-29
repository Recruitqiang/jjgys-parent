package glgc.jjgys.system.service.impl;

import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.model.system.SysDept;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.system.mapper.SysDepartMapper;
import glgc.jjgys.system.service.SysDepartService;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

@Service
public class SysDepartServiceImpl extends ServiceImpl<SysDepartMapper, SysDept> implements SysDepartService {


    @Override
    public String selectParentDeptName(String deptId) {
        // 根据用户表中的部门id，查询在部门表中的信息
        QueryWrapper<SysDept> wrapper = new QueryWrapper<>();
        if(!StringUtils.isEmpty(deptId)) {
            wrapper.like("id",deptId);
        }
        SysDept sysDept = baseMapper.selectOne(wrapper);
        //获取当前部门id下的treePath
        String treePath = sysDept.getTreePath();
        List<String> list = Arrays.asList(treePath.split(","));
        List<SysDept> sysDeptParent = baseMapper.selectBatchIds(list);
        StringBuffer sb = new StringBuffer();
        for (SysDept dept : sysDeptParent) {
            String name = dept.getName();
            sb.append(name).append(",");
        }
        String str = sb.deleteCharAt(sb.length() - 1).toString();
        return str;
    }
}
