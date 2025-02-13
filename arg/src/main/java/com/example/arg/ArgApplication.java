package com.example.arg;


import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.awt.Desktop;
import java.net.URI;

@SpringBootApplication
public class ArgApplication {

    //USING JAVA 22 and JDK 22
    public static void main(String[] args) {
        SpringApplication.run(ArgApplication.class, args);

        openBrowser("http://localhost:8080");

	}

    private static void openBrowser(String url) {
        try {
            new ProcessBuilder("cmd", "/c", "start", url).start();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
        
}

