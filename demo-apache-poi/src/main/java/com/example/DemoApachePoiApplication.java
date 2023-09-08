package com.example;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.util.HashMap;
import java.util.Map;

@SpringBootApplication
public class DemoApachePoiApplication {

	public static void main(String[] args) {
		SpringApplication.run(DemoApachePoiApplication.class, args);

		Map<String, String> map = new HashMap<>();
		map.put("", "");
	}

}
