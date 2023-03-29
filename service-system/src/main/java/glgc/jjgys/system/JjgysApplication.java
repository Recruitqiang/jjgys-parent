package glgc.jjgys.system;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
@MapperScan("glgc.jjgys.system.mapper")
public class JjgysApplication {

    public static void main(String[] args) {
        SpringApplication.run(JjgysApplication.class, args);
    }
}
