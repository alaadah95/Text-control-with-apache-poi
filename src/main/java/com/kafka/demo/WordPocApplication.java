package com.kafka.demo;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.web.servlet.config.annotation.CorsRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;

@SpringBootApplication
public class WordPocApplication implements WebMvcConfigurer {

	public static void main(String[] args) {
		SpringApplication.run(WordPocApplication.class, args);
	}
	@Override // we can use it here if there is no security .
	public void addCorsMappings(CorsRegistry registry) {
//		registry.addMapping("/**").allowedMethods("*").allowedOrigins("*").allowedHeaders("*").allowCredentials(true);
		registry.addMapping("/**").allowedMethods("*").allowedOrigins("*");
	}
}
