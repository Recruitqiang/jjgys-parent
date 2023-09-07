package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgLjxQl;
import glgc.jjgys.model.project.JjgLqsQl;
import glgc.jjgys.model.projectvo.lqs.LjxqlVo;
import glgc.jjgys.model.projectvo.lqs.QlVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgLqsLjxqlMapper;
import glgc.jjgys.system.service.JjgLqsLjxqlService;
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
public class JjgLqsLjxqlServiceImpl extends ServiceImpl<JjgLqsLjxqlMapper, JjgLjxQl> implements JjgLqsLjxqlService {

    @Autowired
    private JjgLqsLjxqlMapper jjgLqsLjxqlMapper;

    @Override
    public void exportLjxQL(HttpServletResponse response) {
        String fileName = "连接线桥梁清单";
        String sheetName = "连接线桥梁清单";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new LjxqlVo()).finish();
    }

    @Override
    public void importLjxQL(MultipartFile file, String proname) {

        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(LjxqlVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<LjxqlVo>(LjxqlVo.class) {
                                @Override
                                public void handle(List<LjxqlVo> dataList) {
                                    for(LjxqlVo ljxqlVo: dataList)
                                    {
                                        JjgLjxQl jjgLjxQl = new JjgLjxQl();
                                        BeanUtils.copyProperties(ljxqlVo,jjgLjxQl);
                                        jjgLjxQl.setZhq(Double.valueOf(ljxqlVo.getZhq()));
                                        jjgLjxQl.setZhz(Double.valueOf(ljxqlVo.getZhz()));
                                        jjgLjxQl.setProname(proname);
                                        jjgLjxQl.setCreateTime(new Date());
                                        jjgLqsLjxqlMapper.insert(jjgLjxQl);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }


    }
}
