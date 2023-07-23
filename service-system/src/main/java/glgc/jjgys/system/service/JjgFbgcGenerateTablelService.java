package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;

import java.io.IOException;


/**
 * <p>
 *  服务类
 * </p>
 *
 * @author wq
 * @since 2023-03-01
 */
public interface JjgFbgcGenerateTablelService extends IService<Object> {


    void generatePdb(CommonInfoVo commonInfoVo) throws IOException;

    void generateJSZLPdb(CommonInfoVo commonInfoVo) throws IOException;

    void generateBGZBG(String proname) throws IOException;
}
