package com.example.guowangwaveformanalysis.controller;

import com.example.guowangwaveformanalysis.pojo.Result;
import com.example.guowangwaveformanalysis.service.XlsService;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

@Slf4j
@RestController
public class XlsController {

    @Autowired
    private XlsService xlsService;

    @PostMapping("/upload")
    public Map<String, Object> upload(@RequestParam("file") MultipartFile file) {
        Map<String, Object> result = new HashMap<>();
        try {
            String outputPath = xlsService.processExcelFile(file);
            // 假设 outputPath = "outputs/output.docx" 或 "outputs\\output.docx"
            String fileName = outputPath.substring(outputPath.lastIndexOf(File.separator) + 1);
            String downloadUrl = "/download/" + fileName;
            result.put("downloadUrl", downloadUrl);

            result.put("code", 0);
            result.put("msg", "ok");
        } catch (Exception e) {
            result.put("code", 1);
            result.put("msg", "文件处理失败：" + e.getMessage());
        }
        return result;
    }

    @GetMapping("/download/{fileName:.+}")
    public void downloadFile(@PathVariable String fileName, HttpServletResponse response) throws IOException {
        java.io.File file = new java.io.File("outputs", fileName);
        if (!file.exists()) {
            response.setStatus(HttpServletResponse.SC_NOT_FOUND);
            return;
        }
        response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\"");
        try (java.io.FileInputStream fis = new java.io.FileInputStream(file);
             java.io.OutputStream os = response.getOutputStream()) {
            byte[] buffer = new byte[8192];
            int bytesRead;
            while ((bytesRead = fis.read(buffer)) != -1) {
                os.write(buffer, 0, bytesRead);
            }
        }
    }


}
