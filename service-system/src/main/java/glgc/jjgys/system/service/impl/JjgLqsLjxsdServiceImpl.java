package glgc.jjgys.system.service.impl;


import com.alibaba.excel.EasyExcel;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgLjxSd;
import glgc.jjgys.model.project.JjgSfz;
import glgc.jjgys.model.projectvo.lqs.LjxsdVo;
import glgc.jjgys.model.projectvo.lqs.SfzVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgLqsLjxsdMapper;
import glgc.jjgys.system.service.JjgLqsLjxsdService;
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
public class JjgLqsLjxsdServiceImpl extends ServiceImpl<JjgLqsLjxsdMapper, JjgLjxSd> implements JjgLqsLjxsdService {

    @Autowired
    private JjgLqsLjxsdMapper jjgLqsLjxsdMapper;

    @Override
    public void exportLjxsd(HttpServletResponse response) {
        String fileName = "连接线隧道清单";
        String sheetName = "连接线隧道清单";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new LjxsdVo()).finish();
    }

    @Override
    public void importLjxsd(MultipartFile file, String proname) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(LjxsdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<LjxsdVo>(LjxsdVo.class) {
                                @Override
                                public void handle(List<LjxsdVo> dataList) {
                                    for(LjxsdVo ljxsdVo: dataList)
                                    {
                                        JjgLjxSd jjgLjxSd = new JjgLjxSd();
                                        BeanUtils.copyProperties(ljxsdVo,jjgLjxSd);
                                        jjgLjxSd.setZhq(Double.valueOf(ljxsdVo.getZhq()));
                                        jjgLjxSd.setZhz(Double.valueOf(ljxsdVo.getZhz()));
                                        jjgLjxSd.setProname(proname);
                                        jjgLjxSd.setCreateTime(new Date());
                                        jjgLqsLjxsdMapper.insert(jjgLjxSd);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }



    }
}
