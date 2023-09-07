package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.util.StringUtils;

import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgLqsQl;
import glgc.jjgys.model.projectvo.lqs.QlVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.listener.JjgLqsQlListener;
import glgc.jjgys.system.mapper.JjgLqsQlMapper;
import glgc.jjgys.system.service.JjgLqsQlService;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2022-12-14
 */
@Service
public class JjgLqsQlServiceImpl extends ServiceImpl<JjgLqsQlMapper, JjgLqsQl> implements JjgLqsQlService {

    @Autowired
    private JjgLqsQlMapper jjgLqsQlMapper;

    @Override
    public void importQL(MultipartFile file,String proname)  {

        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(QlVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<QlVo>(QlVo.class) {
                                @Override
                                public void handle(List<QlVo> dataList) {
                                    for(QlVo ql: dataList)
                                    {
                                        JjgLqsQl jjgLqsQl = new JjgLqsQl();
                                        BeanUtils.copyProperties(ql,jjgLqsQl);
                                        jjgLqsQl.setZhq(Double.valueOf(ql.getZhq()));
                                        jjgLqsQl.setZhz(Double.valueOf(ql.getZhz()));
                                        jjgLqsQl.setProname(proname);
                                        jjgLqsQl.setCreateTime(new Date());
                                        jjgLqsQlMapper.insert(jjgLqsQl);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
           throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }


    }

    @Override
    public void exportQL(HttpServletResponse response,String proname) {
        String fileName = proname+"桥梁清单";
        String sheetName = "桥梁清单";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new QlVo()).finish();
    }

    @Override
    public List<JjgLqsQl> getqlName(String proname, String htd) {
        List<JjgLqsQl> list = jjgLqsQlMapper.getqlName(proname,htd);
        return list;
    }
}
