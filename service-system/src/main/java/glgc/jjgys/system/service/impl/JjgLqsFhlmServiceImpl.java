package glgc.jjgys.system.service.impl;


import com.alibaba.excel.EasyExcel;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgLqsFhlm;
import glgc.jjgys.model.project.JjgLqsQl;
import glgc.jjgys.model.projectvo.lqs.FhlmVo;
import glgc.jjgys.model.projectvo.lqs.QlVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgLqsFhlmMapper;
import glgc.jjgys.system.service.JjgLqsFhlmService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.Date;
import java.util.List;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-01-10
 */
@Service
public class JjgLqsFhlmServiceImpl extends ServiceImpl<JjgLqsFhlmMapper, JjgLqsFhlm> implements JjgLqsFhlmService {

    @Autowired
    private JjgLqsFhlmMapper jjgLqsFhlmMapper;

    @Override
    public void export(HttpServletResponse response) {
        String fileName = "复合路面清单";
        String sheetName = "复合路面清单";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new FhlmVo()).finish();
    }

    @Override
    public void importFhlm(MultipartFile file, String proname) {

        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(FhlmVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<FhlmVo>(FhlmVo.class) {
                                @Override
                                public void handle(List<FhlmVo> dataList) {
                                    for(FhlmVo fhlm: dataList)
                                    {
                                        JjgLqsFhlm jjgLqsFhlm = new JjgLqsFhlm();
                                        BeanUtils.copyProperties(fhlm,jjgLqsFhlm);
                                        jjgLqsFhlm.setProname(proname);
                                        jjgLqsFhlm.setCreateTime(new Date());
                                        jjgLqsFhlmMapper.insert(jjgLqsFhlm);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }
}
