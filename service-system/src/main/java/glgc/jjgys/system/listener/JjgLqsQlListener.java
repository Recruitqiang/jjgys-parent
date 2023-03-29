package glgc.jjgys.system.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import glgc.jjgys.model.project.JjgLqsQl;
import glgc.jjgys.model.projectvo.lqs.QlVo;
import glgc.jjgys.system.mapper.JjgLqsQlMapper;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

@Component
public class JjgLqsQlListener extends AnalysisEventListener<QlVo> {

    @Autowired
    private JjgLqsQlMapper jjgLqsQlMapper;

    @Override
    public void invoke(QlVo qlVo, AnalysisContext analysisContext) {
        JjgLqsQl jjgLqsQl = new JjgLqsQl();
        BeanUtils.copyProperties(qlVo,jjgLqsQl);
        jjgLqsQlMapper.insert(jjgLqsQl);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {

    }
}
