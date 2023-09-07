package glgc.jjgys.system.controller;

import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcJagcService;
import glgc.jjgys.system.service.OperLogService;
import glgc.jjgys.system.utils.JjgFbgcUtils;
import io.swagger.annotations.ApiOperation;
import net.lingala.zip4j.core.ZipFile;
import net.lingala.zip4j.exception.ZipException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;

@RestController
@RequestMapping("/jjg/fbgc/jagc")
@CrossOrigin
public class JjgFbgcJagcController {
    @Autowired
    private JjgFbgcJagcService jjgFbgcJagcService;

    @Autowired
    private OperLogService operLogService;


    @Value(value = "${jjgys.path.filepath}")
    private String filespath;


    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname, String htd) throws IOException {
        String fileName = "交安工程鉴定表.zip";

        String workpath=filespath+ File.separator+proname+File.separator+htd;
        jjgFbgcJagcService.download(response,fileName,proname,htd,filespath+ File.separator+proname+File.separator+htd);
        try {
            fileName=URLEncoder.encode(fileName, "UTF-8");
        } catch (UnsupportedEncodingException e) {
            throw new RuntimeException(e);
        }
        response.setHeader("Content-disposition", "attachment; filename=" +fileName);
        response.setContentType("application/zip;charset=utf-8");
        //response.setCharacterEncoding("utf-8");
        try {
            JjgFbgcUtils.zipFile(workpath+"/交安工程",response.getOutputStream());
        } catch (ZipException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        JjgFbgcUtils.deleteDirAndFiles(new File(workpath+"/交安工程"));

    }
    @ApiOperation("生成交安工程鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcJagcService.generateJdb(commonInfoVo);

    }
    @ApiOperation("交安工程模板文件导出")
    @GetMapping("exportjagc")
    public void exportjagc(HttpServletResponse response) {


        String filepath = System.getProperty("user.dir");
        jjgFbgcJagcService.exportjagc(response,filepath);
        String zipName = "交安工程指标模板文件";
        String downloadName = null;

        try {
            downloadName = URLEncoder.encode(zipName + ".zip", "UTF-8");
        } catch (UnsupportedEncodingException e) {
            throw new RuntimeException(e);
        }


        response.setHeader("Content-disposition", "attachment; filename=" + downloadName);
        response.setContentType("application/zip;charset=utf-8");
        //response.setCharacterEncoding("utf-8");
        try {
            JjgFbgcUtils.zipFile(filepath+"/交安工程",response.getOutputStream());
        } catch (ZipException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        JjgFbgcUtils.deleteDirAndFiles(new File(filepath+"/交安工程"));

    }
    @ApiOperation("交安工程模板数据文件导入")
    @PostMapping("importjagc")
    public Result importjagc(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        String filepath = System.getProperty("user.dir");
        File file1=JjgFbgcUtils.multipartFileToFile(file);
        ZipFile zipFile= null;
        try {
            zipFile = new ZipFile(file1);
            zipFile.setFileNameCharset("GBK");
            JjgFbgcUtils.createDirectory("暂存", filepath);
            zipFile.extractAll(filepath + "/暂存");
        } catch (ZipException e) {
            throw new RuntimeException(e);
        }

        jjgFbgcJagcService.importjagc(commonInfoVo,filepath+"/暂存");
        file1.delete();
        return Result.ok();

    }
}
