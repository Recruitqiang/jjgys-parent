package glgc.jjgys.system.service;

import com.baomidou.mybatisplus.extension.service.IService;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;


/**
 * <p>
 *  服务类
 * </p>
 *
 * @author wq
 * @since 2023-03-01
 */
public interface JjgFbgcGenerateWordService extends IService<Object> {


    void generateword(String proname) throws IOException, InvalidFormatException;
}
