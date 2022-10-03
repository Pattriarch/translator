package com.example.translator;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.context.annotation.SessionScope;

@Configuration
public class Config {
//    @Bean
//    @SessionScope
//    public ExcelReader sessionScopedBean() {
//        return new ExcelReader();
//    }

    @Bean
    @SessionScope
    public ExcelWriter sessionScopedBean2() {
        return new ExcelWriter();
    }
}
