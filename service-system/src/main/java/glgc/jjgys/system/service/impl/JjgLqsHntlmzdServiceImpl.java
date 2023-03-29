package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgLqsFhlm;
import glgc.jjgys.model.project.JjgLqsHntlmzd;
import glgc.jjgys.model.projectvo.lqs.FhlmVo;
import glgc.jjgys.model.projectvo.lqs.HntlmzdVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgLqsHntlmzdMapper;
import glgc.jjgys.system.service.JjgLqsHntlmzdService;
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
public class JjgLqsHntlmzdServiceImpl extends ServiceImpl<JjgLqsHntlmzdMapper, JjgLqsHntlmzd> implements JjgLqsHntlmzdService {

    @Autowired
    private JjgLqsHntlmzdMapper jjgLqsHntlmzdMapper;

    @Override
    public void export(HttpServletResponse response) {
        String fileName = "混凝土路面及匝道清单";
        String sheetName = "混凝土路面及匝道清单";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new HntlmzdVo()).finish();
    }

    @Override
    public void importhntlmzd(MultipartFile file, String proname) {

        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(HntlmzdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<HntlmzdVo>(HntlmzdVo.class) {
                                @Override
                                public void handle(List<HntlmzdVo> dataList) {
                                    for (HntlmzdVo hntlmzdVo : dataList) {
                                        JjgLqsHntlmzd jjgLqsHntlmzd = new JjgLqsHntlmzd();
                                        BeanUtils.copyProperties(hntlmzdVo, jjgLqsHntlmzd);
                                        jjgLqsHntlmzd.setProname(proname);
                                        jjgLqsHntlmzd.setCreateTime(new Date());
                                        jjgLqsHntlmzdMapper.insert(jjgLqsHntlmzd);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001, "解析excel出错，请传入正确格式的excel");
        }

    }
}
