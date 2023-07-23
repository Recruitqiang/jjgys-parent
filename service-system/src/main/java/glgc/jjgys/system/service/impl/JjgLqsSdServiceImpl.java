package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgLqsQl;
import glgc.jjgys.model.project.JjgLqsSd;
import glgc.jjgys.model.projectvo.lqs.SdVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgLqsSdMapper;
import glgc.jjgys.system.service.JjgLqsSdService;
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
 * @since 2023-01-09
 */
@Service
public class JjgLqsSdServiceImpl extends ServiceImpl<JjgLqsSdMapper, JjgLqsSd> implements JjgLqsSdService {

    @Autowired
    private JjgLqsSdMapper jjgLqsSdMapper;

    /**
     * 隧道清单模板文件导出
     * @param response
     */
    @Override
    public void exportSD(HttpServletResponse response,String projectname) {
        String fileName = projectname+"隧道清单";
        String sheetName = "隧道清单";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new SdVo()).finish();

    }

    /**
     * 隧道信息导入
     * @param file
     * @param proname
     */
    @Override
    public void importSD(MultipartFile file, String proname) {

        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(SdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<SdVo>(SdVo.class) {
                                @Override
                                public void handle(List<SdVo> dataList) {
                                    for(SdVo sd: dataList)
                                    {
                                        JjgLqsSd jjgLqsSd = new JjgLqsSd();
                                        BeanUtils.copyProperties(sd,jjgLqsSd);
                                        jjgLqsSd.setProname(proname);
                                        jjgLqsSd.setCreateTime(new Date());
                                        jjgLqsSdMapper.insert(jjgLqsSd);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }


    }

    @Override
    public List<JjgLqsSd> getsdName(String proname, String htd) {
        List<JjgLqsSd> list = jjgLqsSdMapper.getsdName(proname,htd);
        return list;

    }
}
