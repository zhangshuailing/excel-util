package nk.gk.wyl.excel.config;

import org.springframework.context.annotation.Bean;
import org.springframework.stereotype.Component;

/**
 * 文件初始化数据
 */
@Component
public class ContextConfig {


    @Bean
    public int initStatic(){

        return 0;
    }

    /**
     * 获取项目路径
     * @return
     */
    public String getContextPath(){
        String path = System.getProperty("user.dir");
        return path;
    }
}
