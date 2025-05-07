package com.example.guowangwaveformanalysis.controller;

import com.example.guowangwaveformanalysis.pojo.Result;
import com.example.guowangwaveformanalysis.service.XlsService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.UUID;
import org.springframework.core.io.FileSystemResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.*;
import java.io.IOException;
import java.nio.file.StandardCopyOption;


/**
 * @author nan chao
 * @since 2025/4/8 11:18
 */


@Slf4j
@RestController
public class XlsController {


    @Autowired
    private XlsService xlsService;

    @PostMapping("/upload")
    public Result upload(@RequestParam("file") MultipartFile file) throws IOException {

        try {
            String downloadUrl = xlsService.processExcelFile(file);
            return Result.success(downloadUrl);
        } catch (Exception e) {
            return Result.error("文件处理失败：" + e.getMessage());
        }

    }



}




