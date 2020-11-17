package nk.gk.wyl.excel.impl;

import nk.gk.wyl.excel.api.ExcelService;
import nk.gk.wyl.excel.util.CommonValidator;
import nk.gk.wyl.excel.util.ExcelUtil;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;
import java.util.Map;

@Service
public class ExcelServiceImpl implements ExcelService {


    /**
     * 标准的Excel读取（行列没有合并的）
     *
     * @param file        文件
     * @param key_row_num excel表中那一列作为key值
     * @return 返回数据
     */
    @Override
    public Map<String,List<Map<String,String>>> getDataByExcelStandard(MultipartFile file,int key_row_num) throws Exception {
        if(key_row_num <= 0){
           key_row_num = 1;
        }
        // 文件后缀
        String suffix = CommonValidator.checkFile(file);
        // 获取数据
        Map<String,List<Map<String,String>>> data = ExcelUtil.readExcel(file.getInputStream(),suffix,key_row_num);
        return data;
    }
}
