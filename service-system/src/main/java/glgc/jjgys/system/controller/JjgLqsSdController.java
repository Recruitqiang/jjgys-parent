package glgc.jjgys.system.controller;

import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.metadata.IPage;
import com.baomidou.mybatisplus.extension.plugins.pagination.Page;
import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.project.JjgLqsSd;
import glgc.jjgys.system.service.JjgLqsSdService;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.util.List;

/**
 * <p>
 *  隧道清单
 * </p>
 *
 * @author wq
 * @since 2023-01-09
 */
@RestController
@RequestMapping("/project/info/lqs/sd")
@CrossOrigin
public class JjgLqsSdController {

    @Autowired
    private JjgLqsSdService jjgLqsSdService;

    /**
     * 隧道清单模板文件导出
     * @param response
     */
    @ApiOperation("隧道清单文件导出")
    @GetMapping("exportSD/{projectname}")
    public void exportSD(HttpServletResponse response,
                         @PathVariable String projectname){
        jjgLqsSdService.exportSD(response,projectname);
    }

    @ApiOperation(value = "隧道信息导入")
    @PostMapping("importSD")
    public Result importQL(@RequestParam("file") MultipartFile file, @RequestParam String proname) {
        jjgLqsSdService.importSD(file,proname);
        return Result.ok();
    }

    /**
     * 分页查询
     * @param current
     * @param limit
     * @param jjgLqsSd
     * @return
     */
    @PostMapping("findQueryPage/{current}/{limit}")
    public Result findQueryPage(@PathVariable long current,
                                @PathVariable long limit,
                                @RequestBody JjgLqsSd jjgLqsSd){
        //创建page对象
        Page<JjgLqsSd> pageParam=new Page<>(current,limit);
        //判断projectQueryVo对象是否为空，直接查全部
        if(jjgLqsSd == null){
            IPage<JjgLqsSd> pageModel = jjgLqsSdService.page(pageParam,null);
            return Result.ok(pageModel);
        }else {
            //获取条件值，进行非空判断，条件封装
            String sdname = jjgLqsSd.getSdname();
            QueryWrapper<JjgLqsSd> wrapper=new QueryWrapper<>();
            if (!StringUtils.isEmpty(sdname)){
                wrapper.like("name",sdname);
            }
            wrapper.like("proname",jjgLqsSd.getProname());
            wrapper.orderByDesc("create_time");
            //调用方法分页查询
            IPage<JjgLqsSd> pageModel = jjgLqsSdService.page(pageParam, wrapper);
            //返回
            return Result.ok(pageModel);

        }
    }

    @ApiOperation("批量删除隧道清单信息")
    @DeleteMapping("removeBatch")
    public Result removeBeatch(@RequestBody List<String> idList){
        boolean ql = jjgLqsSdService.removeByIds(idList);
        if(ql){
            return Result.ok();
        } else {
            return Result.fail().message("删除失败！");
        }

    }


}
