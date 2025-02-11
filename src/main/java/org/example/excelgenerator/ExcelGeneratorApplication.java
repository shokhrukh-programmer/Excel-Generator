package org.example.excelgenerator;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.autoconfigure.jdbc.DataSourceAutoConfiguration;

@SpringBootApplication(exclude = {DataSourceAutoConfiguration.class})
public class ExcelGeneratorApplication {

	public static void main(String[] args) {
		SpringApplication.run(ExcelGeneratorApplication.class, args);
	}

}
