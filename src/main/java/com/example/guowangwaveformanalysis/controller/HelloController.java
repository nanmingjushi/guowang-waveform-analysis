package com.example.guowangwaveformanalysis.controller;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

/**
 * @author nan chao
 * @since 2025/4/8 11:08
 */

@RestController
public class HelloController {
    @RequestMapping("/hello")
    public String hello(){
        return "hello springboot";
    }
}
