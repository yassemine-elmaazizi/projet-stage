package com.example.powerpointtemplate;

import com.example.powerpointtemplate.services.PowerPointService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;

@SpringBootApplication
public class PowerpointtemplateApplication {

	public static void main(String[] args) {
			SpringApplication.run(PowerpointtemplateApplication.class, args);
		}
	@Bean
	public CommandLineRunner demo(PowerPointService powerPointService) {
		return (args) -> powerPointService.processPresentation();
	}

}
