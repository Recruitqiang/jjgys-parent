package glgc.jjgys.system.controller;

import glgc.jjgys.system.service.JjgLqsService;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.UnsupportedEncodingException;

@Api(tags = "路桥隧")
@RestController
@RequestMapping("/project/info/lqs")
@CrossOrigin
public class JjgLqsController {

    @Autowired
    private JjgLqsService jjgLqsService;

    //路桥隧文件导出
    @ApiOperation("路桥隧文件导出")
    @GetMapping("exportLqs/{projectname}")
    public void exportLqs(HttpServletResponse response,
                          @PathVariable String projectname){
        jjgLqsService.exportLqs(response,projectname);
    }

    //路桥隧文件导入
    @ApiOperation("路桥隧文件导入")
    @PostMapping("importlqs")
    public void importlqs(@RequestParam("file") MultipartFile file,
                          @RequestParam String projectname) throws IOException {
        jjgLqsService.importlqs(file,projectname);
    }


}
