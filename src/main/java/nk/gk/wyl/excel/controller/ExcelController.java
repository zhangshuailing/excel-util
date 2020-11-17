package nk.gk.wyl.excel.controller;

import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import nk.gk.wyl.excel.api.ExcelService;
import nk.gk.wyl.excel.entity.result.Response;
import nk.gk.wyl.excel.entity.search.ExcelModel;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.Map;

@RestController
@RequestMapping("api/v1/excel")
@Api(tags = "Excel接口")
public class ExcelController {
    @Autowired
    private ExcelService excelService;


    @PostMapping(value = "uploadTxt",consumes = "multipart/*",headers = "content-type=multipart/form-data")
    @ApiOperation(value = "word文件上传返回文本")
    public @ResponseBody Response uploadTxt(MultipartFile file,  @RequestParam("key_row_line") int key_row_line){
        try {
            return new Response().success(excelService.getDataByExcelStandard(file,key_row_line));
        } catch (Exception e) {
            e.printStackTrace();
            return new Response().error(e.getMessage());
        }
    }

}
