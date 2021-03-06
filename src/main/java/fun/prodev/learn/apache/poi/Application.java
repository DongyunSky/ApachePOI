package fun.prodev.learn.apache.poi;

import org.springframework.boot.Banner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

/**
 * @author prodev
 * @date 2019/5/12 20:31
 * @description
 */
@SpringBootApplication
public class Application {

    public static void main(String[] args) {
        // SpringApplication.run(fun.prodev.learn.apache.poi.Application.class, args);
        SpringApplication application = new SpringApplication(Application.class);
        application.setBannerMode(Banner.Mode.OFF); // 关闭启动banner
        application.run(args);
    }

}
