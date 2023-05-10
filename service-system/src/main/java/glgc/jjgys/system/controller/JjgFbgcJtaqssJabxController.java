package glgc.jjgys.system.controller;


import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.common.utils.IpUtil;
import glgc.jjgys.common.utils.JwtHelper;
import glgc.jjgys.model.project.JjgFbgcJtaqssJabx;
import glgc.jjgys.model.projectvo.ljgc.CommonInfoVo;
import glgc.jjgys.model.system.SysOperLog;
import glgc.jjgys.system.annotation.Log;
import glgc.jjgys.system.enums.BusinessType;
import glgc.jjgys.system.service.JjgFbgcJtaqssJabxService;
import glgc.jjgys.system.service.OperLogService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.context.request.RequestAttributes;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;


/**
 * <p>
 *  前端控制器
 * </p>
 *
 * @author wq
 * @since 2023-03-01
 */
@RestController
@RequestMapping("/jjg/fbgc/jtaqss/jabx")
public class JjgFbgcJtaqssJabxController {

    @Autowired
    private JjgFbgcJtaqssJabxService jjgFbgcJtaqssJabxService;

    @Autowired
    private OperLogService operLogService;

    @Value(value = "${jjgys.path.filepath}")
    private String filespath;


    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void downloadExport(HttpServletRequest request, HttpServletResponse response, String proname, String htd) throws IOException {
        String fpath= filespath+File.separator+proname+File.separator+htd;
        String fileName1 = "57交安标线厚度.xlsx";
        String fileName2 = "57交安标线白线逆反射系数.xlsx";
        String fileName3 = "57交安标线黄线逆反射系数.xlsx";
        List list = new ArrayList();
        list.add(fileName1);
        list.add(fileName2);
        list.add(fileName3);
        //设置响应头信息csc
        /*response.reset();
        response.setCharacterEncoding("utf-8");
        response.setContentType("multipart/form-data");*/
        //设置压缩包的名字
        String downloadName = URLEncoder.encode("57交安标线.zip", "UTF-8");
        //String userAgent = request.getHeader("User-Agent");
        response.reset();
        response.setHeader("Content-disposition", "attachment; filename=" + downloadName);
        response.setContentType("application/zip;charset=utf-8");
        response.setCharacterEncoding("utf-8");
        //返回客户端浏览器的版本号、类型
        /*String agent = request.getHeader("USER-AGENT");
        try {
            //针对IE或者以IE为内核的浏览器：
            if (agent.contains("MSIE") || agent.contains("Trident")) {
                downloadName = java.net.URLEncoder.encode(downloadName, "UTF-8");
            } else {
                //非IE浏览器的处理：
                downloadName = new String(downloadName.getBytes(StandardCharsets.UTF_8), StandardCharsets.ISO_8859_1);
            }
            System.out.println(downloadName);
            System.out.println("111");
        } catch (Exception e) {
            e.printStackTrace();
        }*/
        //response.setHeader("Content-Disposition", "attachment;fileName=\"" + downloadName + "\"");
        //设置压缩流：直接写入response，实现边压缩边下载
        ZipOutputStream zipOs = null;
        //循环将文件写入压缩流
        DataOutputStream os = null;
        //文件
        File file;
        try {
            zipOs = new ZipOutputStream(new BufferedOutputStream(response.getOutputStream()));
            //设置压缩方法
            zipOs.setMethod(ZipOutputStream.DEFLATED);
            //遍历文件信息（主要获取文件名/文件路径等）
            for (int i=0;i<list.size();i++) {
                String name = list.get(i).toString();
                String path = fpath+File.separator+name;
                file = new File(path);
                if (!file.exists()) {
                    break;
                }
                //添加ZipEntry，并将ZipEntry写入文件流
                zipOs.putNextEntry(new ZipEntry(name));
                os = new DataOutputStream(zipOs);
                FileInputStream fs = new FileInputStream(file);
                byte[] b = new byte[100];
                int length;
                //读入需要下载的文件的内容，打包到zip文件
                while ((length = fs.read(b)) != -1) {
                    os.write(b, 0, length);
                }
                //关闭流
                fs.close();
                zipOs.closeEntry();

            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            //关闭流
            try {
                if (os != null) {
                    os.flush();
                    os.close();
                }
                if (zipOs != null) {
                    zipOs.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

    @ApiOperation("生成交安标线鉴定表")
    @PostMapping("generateJdb")
    public void generateJdb(@RequestBody CommonInfoVo commonInfoVo) throws Exception {
        jjgFbgcJtaqssJabxService.generateJdb(commonInfoVo);

    }

    @ApiOperation("查看交安标线鉴定结果")
    @PostMapping("lookJdbjg")
    public Result lookJdbjg(@RequestBody CommonInfoVo commonInfoVo) throws IOException {
        List<Map<String,Object>> jdjg = jjgFbgcJtaqssJabxService.lookJdbjg(commonInfoVo);
        return Result.ok(jdjg);

    }

    @ApiOperation("交安标线模板文件导出")
    @GetMapping("exportjabx")
    public void exportjabx(HttpServletResponse response){
        jjgFbgcJtaqssJabxService.exportjabx(response);
    }

    @ApiOperation(value = "交安标线数据文件导入")
    @PostMapping("importjabx")
    public Result importjabx(@RequestParam("file") MultipartFile file, CommonInfoVo commonInfoVo) {
        jjgFbgcJtaqssJabxService.importjabx(file,commonInfoVo);
        return Result.ok();
    }

    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgFbgcJtaqssJabx jjgFbgcJtaqssJabx) {
        //创建page对象
        Page<JjgFbgcJtaqssJabx> pageParam = new Page<>(current, limit);
        if (jjgFbgcJtaqssJabx != null) {
            QueryWrapper<JjgFbgcJtaqssJabx> wrapper = new QueryWrapper<>();
            wrapper.like("proname", jjgFbgcJtaqssJabx.getProname());
            wrapper.like("htd", jjgFbgcJtaqssJabx.getHtd());
            wrapper.like("fbgc", jjgFbgcJtaqssJabx.getFbgc());
            Date jcsj = jjgFbgcJtaqssJabx.getJcsj();
            if (!StringUtils.isEmpty(jcsj)) {
                wrapper.like("jcsj", jcsj);
            }
            //调用方法分页查询
            IPage<JjgFbgcJtaqssJabx> pageModel = jjgFbgcJtaqssJabxService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);
        }
        return Result.ok().message("无数据");
    }


    @ApiOperation("批量删除交安标线数据")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean hd = jjgFbgcJtaqssJabxService.removeByIds(idList);
        if(hd){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }

    @ApiOperation("根据id查询")
    @GetMapping("getJabx/{id}")
    public Result getJabx(@PathVariable String id) {
        JjgFbgcJtaqssJabx user = jjgFbgcJtaqssJabxService.getById(id);
        return Result.ok(user);
    }

    //@Log(title = "交安标线数据",businessType = BusinessType.UPDATE)
    @ApiOperation("修改交安标线数据")
    @PostMapping("update")
    public Result update(@RequestBody JjgFbgcJtaqssJabx user) {
        RequestAttributes ra = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes sra = (ServletRequestAttributes) ra;
        HttpServletRequest request = sra.getRequest();
        boolean is_Success = jjgFbgcJtaqssJabxService.updateById(user);
        if(is_Success) {
            SysOperLog sysOperLog = new SysOperLog();
            sysOperLog.setProname(user.getProname());
            sysOperLog.setHtd(user.getHtd());
            sysOperLog.setFbgc(user.getFbgc());
            sysOperLog.setTitle("交安标线数据");
            sysOperLog.setBusinessType("修改");
            sysOperLog.setOperName(JwtHelper.getUsername(request.getHeader("token")));
            sysOperLog.setOperIp(IpUtil.getIpAddress(request));
            sysOperLog.setOperTime(new Date());
            operLogService.saveSysLog(sysOperLog);
            return Result.ok();
        } else {
            return Result.fail();
        }
    }

}

