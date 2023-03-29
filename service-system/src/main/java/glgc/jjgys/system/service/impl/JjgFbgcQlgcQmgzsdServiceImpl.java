package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.JjgFbgcLjgcHdgqd;
import glgc.jjgys.model.project.JjgFbgcQlgcQmgzsd;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.projectvo.ljgc.JjgFbgcLjgcHdgqdVo;
import glgc.jjgys.model.projectvo.qlgc.JjgFbgcQlgcQmgzsdVo;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.JjgFbgcQlgcQmgzsdMapper;
import glgc.jjgys.system.service.JjgFbgcQlgcQmgzsdService;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * <p>
 *  服务实现类
 * </p>
 *
 * @author wq
 * @since 2023-03-20
 */
@Service
public class JjgFbgcQlgcQmgzsdServiceImpl extends ServiceImpl<JjgFbgcQlgcQmgzsdMapper, JjgFbgcQlgcQmgzsd> implements JjgFbgcQlgcQmgzsdService {

    @Autowired
    private JjgFbgcQlgcQmgzsdMapper jjgFbgcQlgcQmgzsdMapper;

    @Value(value = "${jjgys.path.filepath}")
    private String filepath;

    @Override
    public void generateJdb(CommonInfoVo commonInfoVo) throws IOException {


    }

    @Override
    public List<Map<String, Object>> lookJdbjg(CommonInfoVo commonInfoVo) {
        return null;
    }

    @Override
    public void exportqmgzsd(HttpServletResponse response) {
        String fileName = "桥面构造深度实测数据";
        String sheetName = "桥面构造深度实测数据";
        ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName, new JjgFbgcQlgcQmgzsdVo()).finish();

    }

    @Override
    public void importqmgzsd(MultipartFile file, CommonInfoVo commonInfoVo) {
        try {
            EasyExcel.read(file.getInputStream())
                    .sheet(0)
                    .head(JjgFbgcQlgcQmgzsdVo.class)
                    .headRowNumber(1)
                    .registerReadListener(
                            new ExcelHandler<JjgFbgcQlgcQmgzsdVo>(JjgFbgcQlgcQmgzsdVo.class) {
                                @Override
                                public void handle(List<JjgFbgcQlgcQmgzsdVo> dataList) {
                                    for(JjgFbgcQlgcQmgzsdVo qmgzsdVo: dataList)
                                    {
                                        JjgFbgcQlgcQmgzsd fbgcQlgcQmgzsd = new JjgFbgcQlgcQmgzsd();
                                        BeanUtils.copyProperties(qmgzsdVo,fbgcQlgcQmgzsd);
                                        fbgcQlgcQmgzsd.setCreatetime(new Date());
                                        fbgcQlgcQmgzsd.setProname(commonInfoVo.getProname());
                                        fbgcQlgcQmgzsd.setHtd(commonInfoVo.getHtd());
                                        fbgcQlgcQmgzsd.setFbgc(commonInfoVo.getFbgc());
                                        jjgFbgcQlgcQmgzsdMapper.insert(fbgcQlgcQmgzsd);
                                    }
                                }
                            }
                    ).doRead();
        } catch (IOException e) {
            throw new JjgysException(20001,"解析excel出错，请传入正确格式的excel");
        }

    }
}
