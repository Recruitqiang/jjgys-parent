package glgc.jjgys.system.controller;


import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcGenerateTablelService;
import glgc.jjgys.system.service.JjgFbgcGenerateWordService;
import glgc.jjgys.system.service.JjgHtdService;
import io.swagger.annotations.ApiOperation;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;



/**
 * <p>
 *  前端控制器
 * </p>
 *
 * @author wq
 * @since 2023-03-01
 */
@RestController
@RequestMapping("/jjg/fbgc/generate/word")
public class JjgFbgcGenerateWordController {

    @Autowired
    private JjgFbgcGenerateWordService jjgFbgcGenerateWordService;


    @ApiOperation("生成word")
    @PostMapping("generateword")
    public Result generateword(String proname) throws IOException, InvalidFormatException {
        jjgFbgcGenerateWordService.generateword(proname);
        return Result.ok();
    }
}

