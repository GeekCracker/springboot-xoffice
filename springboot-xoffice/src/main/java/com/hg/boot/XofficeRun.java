package com.hg.boot;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.EnableAutoConfiguration;
import org.springframework.context.annotation.ComponentScan;
import org.springframework.context.annotation.PropertySource;
import org.springframework.scheduling.annotation.EnableScheduling;

@ComponentScan(basePackages="com.hg")
@EnableAutoConfiguration
@EnableScheduling
@PropertySource({ "classpath:config.properties"})
public class XofficeRun {

	public static void main(String[] args) {
		SpringApplication application = new SpringApplication(XofficeRun.class);
		application.run(args);
	}
}