package glgc.jjgys.system.controller;


import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcGenerateTablelService;
import glgc.jjgys.system.service.JjgHtdService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.*;

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
@RequestMapping("/jjg/fbgc/generate/table")
public class JjgFbgcGenerateTableController {

    @Autowired
    private JjgHtdService jjgHtdService;

    @Autowired
    private JjgFbgcGenerateTablelService jjgFbgcGenerateTablelService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;

    @ApiOperation("生成评定表")
    @PostMapping("generatePdb")
    public void generatePdb(@RequestBody CommonInfoVo commonInfoVo) throws IOException {

        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        jjgFbgcGenerateTablelService.generatePdb(commonInfoVo);

    }

    @ApiOperation("生成建设项目质量评定表")
    @PostMapping("generateJSZLPdb")
    public void generateJSZLPdb(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        String htd = commonInfoVo.getHtd();
        jjgFbgcGenerateTablelService.generateJSZLPdb(commonInfoVo);

    }

    @ApiOperation("报告中表格")
    @PostMapping("generateBGZBG")
    public void generateBGZBG(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        String proname = commonInfoVo.getProname();
        jjgFbgcGenerateTablelService.generateBGZBG(proname);

    }




}

