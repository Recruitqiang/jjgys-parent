package glgc.jjgys.system.controller;

import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.system.service.JjgFbgcLjgcService;
import glgc.jjgys.system.service.JjgFbgcQlgcService;
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

/**
 * <p>
 *  前端控制器
 * </p>
 *
 * @author lhy
 * @since 2023-07-12
 */
@RestController
@RequestMapping("/jjg/fbgc/qlgc")
@CrossOrigin
public class JjgFbgcQlgcController {

    @Autowired
    private JjgFbgcQlgcService jjgFbgcQlgcService;

    @Autowired
    private OperLogService operLogService;


    @Value(value = "${jjgys.path.filepath}")
    private String filespath;


    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletResponse response, String proname, String htd) throws IOException {
        String fileName = "桥梁工程鉴定表.zip";

        String workpath =filespath+ File.separator+proname+File.separator+htd;
        jjgFbgcQlgcService.download(response,fileName,proname,htd,filespath+ File.separator+proname+File.separator+htd);
        try {
            fileName=URLEncoder.encode(fileName, "UTF-8");
        } catch (UnsupportedEncodingException e) {
            throw new RuntimeException(e);
        }
        response.setHeader("Content-disposition", "attachment; filename=" +fileName);
        response.setContentType("application/zip;charset=utf-8");
        //response.setCharacterEncoding("utf-8");
        try {
            JjgFbgcUtils.zipFile(workpath+"/桥梁工程",response.getOutputStream());
        } catch (ZipException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        JjgFbgcUtils.deleteDirAndFiles(new File(workpath+"/桥梁工程"));

    }
    @ApiOperation("生成桥梁工程鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcQlgcService.generateJdb(commonInfoVo);

    }
    @ApiOperation("桥梁工程模板文件导出")
    @GetMapping("exportqlgc")
    public void exportqlgc(HttpServletResponse response) {


        String filepath = System.getProperty("user.dir");
        jjgFbgcQlgcService.exportqlgc(response,filepath);
        String zipName = "桥梁工程指标模板文件";
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
            JjgFbgcUtils.zipFile(filepath+"/桥梁工程",response.getOutputStream());
        } catch (ZipException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        JjgFbgcUtils.deleteDirAndFiles(new File(filepath+"/桥梁工程"));

    }
    @ApiOperation("桥梁工程模板数据文件导入")
    @PostMapping("importqlgc")
    public Result importqlgc(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
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

        jjgFbgcQlgcService.importqlgc(commonInfoVo,filepath+"/暂存");
        file1.delete();
        return Result.ok();

    }
}
