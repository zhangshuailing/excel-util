package nk.gk.wyl.excel.api;

import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * excel 接口
 */
public interface ExcelService {

    /**
     * 标准的Excel读取（行列没有合并的）
     * @param file 文件
     * @param key_row_num excel表中那一列作为key值
     * @return 返回数据
     */
    public Map<String,List<Map<String,String>>> getDataByExcelStandard(MultipartFile file,int key_row_num) throws Exception;

}
