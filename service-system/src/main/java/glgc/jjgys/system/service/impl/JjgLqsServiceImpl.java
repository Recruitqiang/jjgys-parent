package glgc.jjgys.system.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import glgc.jjgys.common.excel.ExcelUtil;
import glgc.jjgys.model.project.*;
import glgc.jjgys.model.projectvo.lqs.*;
import glgc.jjgys.system.easyexcel.ExcelHandler;
import glgc.jjgys.system.exception.JjgysException;
import glgc.jjgys.system.mapper.*;
import glgc.jjgys.system.service.JjgLqsService;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.Date;
import java.util.List;

@Service
public class JjgLqsServiceImpl implements JjgLqsService {

    //路桥隧文件导出
    @Override
    public void exportLqs(HttpServletResponse response,String projectname) {

        try {
            String fileName = projectname+"路桥隧数据文件";
            String sheetName1 = "桥梁清单";
            String sheetName2 = "隧道清单";
            String sheetName3 = "复合路面清单";
            String sheetName4 = "混凝土路面及匝道清单";
            String sheetName5 = "收费站清单";
            String sheetName6 = "连接线清单";
            String sheetName7 = "连接线桥梁清单";
            String sheetName8 = "连接线隧道清单";
            String sheetName9 = "连接线混凝土路面清单";
            ExcelUtil.writeExcelWithSheets(response, null, fileName, sheetName1, new QlVo())
                    .write(null, sheetName2, new SdVo())
                    .write(null, sheetName3, new FhlmVo())
                    .write(null, sheetName4, new HntlmzdVo())
                    .write(null, sheetName5, new SfzVo())
                    .write(null, sheetName6, new LjxVo())
                    .write(null, sheetName7, new LjxqlVo())
                    .write(null, sheetName8, new LjxsdVo())
                    .write(null, sheetName9, new LjxhntlmVo())
                    .finish();
        } catch (Exception e) {
            throw new JjgysException(20001,"导出失败");
        }

    }

    @Autowired
    private JjgLqsQlMapper jjgLqsQlMapper;
    @Autowired
    private JjgLqsSdMapper jjgLqsSdMapper;
    @Autowired
    private JjgLqsFhlmMapper jjgLqsFhlmMapper;
    @Autowired
    private JjgLqsHntlmzdMapper jjgLqsHntlmzdMapper;
    @Autowired
    private JjgLqsSfzMapper jjgLqsSfzMapper;
    @Autowired
    private JjgLqsLjxMapper jjgLqsLjxMapper;
    @Autowired
    private JjgLqsLjxqlMapper jjgLqsLjxqlMapper;
    @Autowired
    private JjgLqsLjxsdMapper jjgLqsLjxsdMapper;
    @Autowired
    private JjgLqsLjxhntlmMapper jjgLqsLjxhntlmMapper;

    @Override
    public void importlqs(MultipartFile file, String projectname) throws IOException {
        ExcelReader excelReader = EasyExcel.read(file.getInputStream()).build();
        ReadSheet sheetql = EasyExcel.readSheet(0).head(QlVo.class).registerReadListener(new ExcelHandler<QlVo>(QlVo.class) {
            @Override
            public void handle(List<QlVo> dataList) {
                for(QlVo ql: dataList)
                {
                    JjgLqsQl jjgLqsQl = new JjgLqsQl();
                    BeanUtils.copyProperties(ql,jjgLqsQl);
                    jjgLqsQl.setZhq(Double.valueOf(ql.getZhq()));
                    jjgLqsQl.setZhz(Double.valueOf(ql.getZhz()));
                    jjgLqsQl.setProname(projectname);
                    jjgLqsQl.setCreateTime(new Date());
                    jjgLqsQlMapper.insert(jjgLqsQl);
                }
            }
        }).build();
        ReadSheet sheetsd = EasyExcel.readSheet(1).head(SdVo.class).registerReadListener(new ExcelHandler<SdVo>(SdVo.class) {
            @Override
            public void handle(List<SdVo> dataList) {
                for(SdVo sd: dataList)
                {
                    JjgLqsSd jjgLqsSd = new JjgLqsSd();
                    BeanUtils.copyProperties(sd,jjgLqsSd);
                    jjgLqsSd.setZhq(Double.valueOf(sd.getZhq()));
                    jjgLqsSd.setZhz(Double.valueOf(sd.getZhz()));
                    jjgLqsSd.setProname(projectname);
                    jjgLqsSd.setCreateTime(new Date());
                    jjgLqsSdMapper.insert(jjgLqsSd);
                }
            }
        }).build();
        ReadSheet sheetfhlm = EasyExcel.readSheet(2).head(FhlmVo.class).registerReadListener(new ExcelHandler<FhlmVo>(FhlmVo.class) {
            @Override
            public void handle(List<FhlmVo> dataList) {
                for(FhlmVo fhlm: dataList)
                {
                    JjgLqsFhlm jjgLqsFhlm = new JjgLqsFhlm();
                    BeanUtils.copyProperties(fhlm,jjgLqsFhlm);
                    jjgLqsFhlm.setZhq(Double.valueOf(fhlm.getZhq()));
                    jjgLqsFhlm.setZhz(Double.valueOf(fhlm.getZhz()));
                    jjgLqsFhlm.setProname(projectname);
                    jjgLqsFhlm.setCreateTime(new Date());
                    jjgLqsFhlmMapper.insert(jjgLqsFhlm);
                }
            }
        }).build();
        ReadSheet sheethntlmzd = EasyExcel.readSheet(3).head(HntlmzdVo.class).registerReadListener(new ExcelHandler<HntlmzdVo>(HntlmzdVo.class) {
            @Override
            public void handle(List<HntlmzdVo> dataList) {
                for (HntlmzdVo hntlmzdVo : dataList) {
                    JjgLqsHntlmzd jjgLqsHntlmzd = new JjgLqsHntlmzd();
                    BeanUtils.copyProperties(hntlmzdVo, jjgLqsHntlmzd);
                    jjgLqsHntlmzd.setProname(projectname);
                    jjgLqsHntlmzd.setCreateTime(new Date());
                    jjgLqsHntlmzdMapper.insert(jjgLqsHntlmzd);
                }
            }
        }).build();
        ReadSheet sheetsfz = EasyExcel.readSheet(4).head(SfzVo.class).registerReadListener(new ExcelHandler<SfzVo>(SfzVo.class) {
            @Override
            public void handle(List<SfzVo> dataList) {
                for(SfzVo sfz: dataList)
                {
                    JjgSfz jjgSfz = new JjgSfz();
                    BeanUtils.copyProperties(sfz,jjgSfz);
                    jjgSfz.setProname(projectname);
                    jjgSfz.setCreateTime(new Date());
                    jjgLqsSfzMapper.insert(jjgSfz);
                }
            }
        }).build();
        ReadSheet sheetljx = EasyExcel.readSheet(5).head(LjxVo.class).registerReadListener(new ExcelHandler<LjxVo>(LjxVo.class) {
            @Override
            public void handle(List<LjxVo> dataList) {
                for(LjxVo ljx: dataList)
                {
                    JjgLjx jjgLjx = new JjgLjx();
                    BeanUtils.copyProperties(ljx,jjgLjx);
                    jjgLjx.setProname(projectname);
                    jjgLjx.setCreateTime(new Date());
                    jjgLqsLjxMapper.insert(jjgLjx);
                }
            }
        }).build();
        ReadSheet sheetljxql = EasyExcel.readSheet(6).head(LjxqlVo.class).registerReadListener(new ExcelHandler<LjxqlVo>(LjxqlVo.class) {
            @Override
            public void handle(List<LjxqlVo> dataList) {
                for(LjxqlVo ljxqlVo: dataList)
                {
                    JjgLjxQl jjgLjxQl = new JjgLjxQl();
                    BeanUtils.copyProperties(ljxqlVo,jjgLjxQl);
                    jjgLjxQl.setZhq(Double.valueOf(ljxqlVo.getZhq()));
                    jjgLjxQl.setZhz(Double.valueOf(ljxqlVo.getZhz()));
                    jjgLjxQl.setProname(projectname);
                    jjgLjxQl.setCreateTime(new Date());
                    jjgLqsLjxqlMapper.insert(jjgLjxQl);
                }
            }
        }).build();
        ReadSheet sheetljxsd = EasyExcel.readSheet(7).head(LjxsdVo.class).registerReadListener(new ExcelHandler<LjxsdVo>(LjxsdVo.class) {
            @Override
            public void handle(List<LjxsdVo> dataList) {
                for(LjxsdVo ljxsdVo: dataList)
                {
                    JjgLjxSd jjgLjxSd = new JjgLjxSd();
                    BeanUtils.copyProperties(ljxsdVo,jjgLjxSd);
                    jjgLjxSd.setZhq(Double.valueOf(ljxsdVo.getZhq()));
                    jjgLjxSd.setZhz(Double.valueOf(ljxsdVo.getZhz()));
                    jjgLjxSd.setProname(projectname);
                    jjgLjxSd.setCreateTime(new Date());
                    jjgLqsLjxsdMapper.insert(jjgLjxSd);
                }
            }
        }).build();
        ReadSheet sheetljxhntlm = EasyExcel.readSheet(8).head(LjxhntlmVo.class).registerReadListener(new ExcelHandler<LjxhntlmVo>(LjxhntlmVo.class) {
            @Override
            public void handle(List<LjxhntlmVo> dataList) {
                for(LjxhntlmVo ljxhntlmVo: dataList)
                {
                    JjgLjxhntlm jjgLjxhntlm = new JjgLjxhntlm();
                    BeanUtils.copyProperties(ljxhntlmVo,jjgLjxhntlm);
                    jjgLjxhntlm.setZhq(Double.valueOf(ljxhntlmVo.getZhq()));
                    jjgLjxhntlm.setZhz(Double.valueOf(ljxhntlmVo.getZhz()));
                    jjgLjxhntlm.setProname(projectname);
                    jjgLjxhntlm.setCreateTime(new Date());
                    jjgLqsLjxhntlmMapper.insert(jjgLjxhntlm);
                }
            }
        }).build();
        excelReader.read(sheetql,sheetsd,sheetfhlm,sheethntlmzd,sheetsfz,sheetljx,sheetljxql,sheetljxsd,sheetljxhntlm);
        excelReader.finish();

    }

}
