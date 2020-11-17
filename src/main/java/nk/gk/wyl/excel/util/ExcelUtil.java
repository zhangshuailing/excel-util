package nk.gk.wyl.excel.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Excel 工具类
 */
public class ExcelUtil {

    private static Logger logger = LoggerFactory.getLogger(ExcelUtil.class);
    private static String XLXS = "xlsx";
    private static String XLS = "xls";

    public static String getXLXS() {
        return XLXS;
    }

    public static String getXLS() {
        return XLS;
    }

    public static void main(String[] args) {
        String str = "F:\\cjg功能点及进度_后端.xlsx";
        System.out.println(readExcel(str));
    }

    /**
     * 读取excel
     * @param path
     * @return
     */
    public static Map<String,List<Map<String,Map<String,String>>>> readExcel(String path){
        // 文件后缀
        String suffix = path.substring(path.lastIndexOf(".")+1,path.length());
        // 文件流
        InputStream inputStream = null;
        try {
            inputStream = new FileInputStream(path);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return readExcel(inputStream,suffix);
    }



    /**
     * 获取 Workbook
     * @param inputStream 文件流
     * @param suffix 文件后缀
     * @return 返回 Workbook
     */
    public static Workbook getWorkbook(InputStream inputStream,String suffix) throws IOException {
        suffix = suffix.toLowerCase();
        Workbook workbook = null;
        // 判断
        if(suffix.equals(XLS)){
            workbook = new HSSFWorkbook(inputStream);
        }else if(suffix.equals(XLXS)){
            workbook = new XSSFWorkbook(inputStream);
        }
        return workbook;
    }

    /**
     * 根据文件流读取excel【标准的Excel读取（行列没有合并的）】
     * @param inputStream
     * @param suffix
     * @return
     */
    public static Map<String,List<Map<String,Map<String,String>>>> readExcel(InputStream inputStream,String suffix){
        // 工作簿
        Workbook workbook = null;
        try {
            workbook = getWorkbook(inputStream,suffix);
        } catch (IOException e) {
            e.printStackTrace();
        }
        if(inputStream!=null){
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        if(workbook == null){
            return null;
        }
        return geData(workbook);
    }

    /**
     * 获取数据
     * @param workbook
     * @return
     */
    public static Map<String,List<Map<String,Map<String,String>>>> geData(Workbook workbook){
        // 获取sheet 数量
        int num = workbook.getNumberOfSheets();
        // 定义数据
        Map<String,List<Map<String,Map<String,String>>>> data = new HashMap<>();
        // 循环
        for (int i = 0; i < num; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            String sheetName = sheet.getSheetName();
            //获取最大行数
            // 定义每一个sheet的数据集合
            List<Map<String,Map<String,String>>> list = new ArrayList<>();
            int rownum = sheet.getPhysicalNumberOfRows();
            for (int j = 0; j < rownum; j++) {
                // 定义没一行的数据对象
                Map<String,Map<String,String>> row_map = new HashMap<>();
                Row row = sheet.getRow(j);
                int cells = row.getLastCellNum();
                // 每一行的数据
                Map<String,String> map1 = new HashMap<>();
                for (int k = 0; k < cells; k++) {
                    map1.put((k+1)+"",row.getCell(k)==null?"":row.getCell(k).toString());
                }
                row_map.put((j+1)+"",map1);
                list.add(row_map);
            }
            data.put(sheetName+"_"+(i+1),list);
        }
        return data;
    }

    /**
     * 根据文件流读取excel【标准的Excel读取（行列没有合并的）】
     * @param inputStream
     * @param suffix
     * @param row_line 第几行作为key
     * @return
     */
    public static Map<String,List<Map<String,String>>> readExcel(InputStream inputStream,
                                                                             String suffix,
                                                                             int row_line) throws Exception {
        // 工作簿
        Workbook workbook = null;
        try {
            workbook = getWorkbook(inputStream,suffix);
        } catch (IOException e) {
            throw new Exception("根据文件流获取Workbook失败："+e.getMessage());
        }
        if(inputStream!=null){
            try {
                inputStream.close();
            } catch (IOException e) {
                throw new Exception("文件流关闭失败："+e.getMessage());
            }
        }
        if(workbook == null){
            return null;
        }
        if(row_line<=0){
            row_line = 1;
        }
        Map<String,List<Map<String,String>>> result = new HashMap<>();
        // 获取sheet 数量
        int num = workbook.getNumberOfSheets();
        // 循环
        for (int i = 0; i < num; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            String sheetName = sheet.getSheetName();
            result.put(sheetName,getDataBySheet(sheet,row_line));
        }
        return result;
    }

    /**
     * 获取每一个sheet数据
     * @param sheet
     * @return
     */
    public static List<Map<String,String>> getDataBySheet(Sheet sheet,int row_line){
        // 定义每一个sheet的数据集合
        List<Map<String,String>> result = new ArrayList<>();
        int rownum = sheet.getPhysicalNumberOfRows();
        // 获取第一次循环的数据，其值是key
        Map<String,String> map_key = new HashMap<>();
        for (int j = (row_line-1); j < rownum; j++) {
            Row row = sheet.getRow(j);
            int cells = row.getLastCellNum();
            // 每一行的数据
            Map<String,String> row_map = new HashMap<>();
            if(j==row_line-1){
                for (int k = 0; k < cells; k++) {
                    map_key.put((k+1)+"",row.getCell(k)==null?"":row.getCell(k).toString());
                }
                logger.info(sheet.getSheetName()+"的key值："+map_key.toString());
            }else{
                for (int k = 0; k < cells; k++) {
                    String key = map_key.get((k+1)+"") == null?"":map_key.get((k+1)+"").toString();
                    if(!"".equals(key)){
                        row_map.put(key,row.getCell(k)==null?"":row.getCell(k).toString());
                    }
                }
                result.add(row_map);
            }
        }
        logger.info(sheet.getSheetName()+"的数据量："+result.size());
        return result;
    }
}
