package glgc.jjgys.system.controller;

import glgc.jjgys.common.result.Result;
import glgc.jjgys.model.system.SysDept;
import glgc.jjgys.system.service.SysDepartService;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@Api(tags = "部门接口")
@RestController
@RequestMapping(value = "/dept")
@CrossOrigin
public class SysDepartController {
    @Autowired
    private SysDepartService sysDepartService;

    @ApiOperation("查询部门名称")
    @GetMapping("selectdept")
    public Result selectdept(){
        List<SysDept> list = sysDepartService.list();
        return Result.ok(list);
    }
}
