package com.example.guowangwaveformanalysis.controller;

import com.example.guowangwaveformanalysis.service.XlsService;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.File;
import java.io.IOException;
import java.util.*;

@Slf4j
@RestController
public class XlsController {

    @Autowired
    private XlsService xlsService;

    @PostMapping("/upload")
    public Map<String, Object> upload(
            @RequestParam("file") MultipartFile file,
            @RequestParam("templateFile") MultipartFile templateFile,
            @RequestParam(value = "images", required = false) MultipartFile[] images,

            // 固定字段
            @RequestParam("reportNo") String reportNo,
            @RequestParam("client") String client,
            @RequestParam("addressOfClient") String addressOfClient,
            @RequestParam("applicant") String applicant,
            @RequestParam("addressOfApplicant") String addressOfApplicant,
            @RequestParam("testSite") String testSite,
            @RequestParam("voltage") String voltage,
            @RequestParam("spot") String spot,
            @RequestParam("environmentTemperature") String environmentTemperature,
            @RequestParam("relativeHumidity") String relativeHumidity,

            // 时间
            @RequestParam("startYear") String startYear,
            @RequestParam("startMonth") String startMonth,
            @RequestParam("startDay") String startDay,
            @RequestParam("startHour") String startHour,
            @RequestParam("startMinute") String startMinute,
            @RequestParam("endYear") String endYear,
            @RequestParam("endMonth") String endMonth,
            @RequestParam("endDay") String endDay,
            @RequestParam("endHour") String endHour,
            @RequestParam("endMinute") String endMinute,

            // 动态仪器，前端需用 JSON.stringify(fields.measurements)
            @RequestParam(value = "measurements", required = false) String measurementsJson
    ) {
        Map<String, Object> result = new HashMap<>();
        try {
            Map<String, String> replaceMap = new HashMap<>();
            replaceMap.put("reportNo", reportNo);
            replaceMap.put("client", client);
            replaceMap.put("addressOfClient", addressOfClient);
            replaceMap.put("applicant", applicant);
            replaceMap.put("addressOfApplicant", addressOfApplicant);
            replaceMap.put("testSite", testSite);
            replaceMap.put("voltage", voltage);
            replaceMap.put("spot", spot);
            replaceMap.put("environmentTemperature", environmentTemperature);
            replaceMap.put("relativeHumidity", relativeHumidity);
            replaceMap.put("startYear", startYear);
            replaceMap.put("startMonth", startMonth);
            replaceMap.put("startDay", startDay);
            replaceMap.put("startHour", startHour);
            replaceMap.put("startMinute", startMinute);
            replaceMap.put("endYear", endYear);
            replaceMap.put("endMonth", endMonth);
            replaceMap.put("endDay", endDay);
            replaceMap.put("endHour", endHour);
            replaceMap.put("endMinute", endMinute);

            // 1. 解析仪器参数（字符串转List<Map>）
            List<Map<String, String>> measurementList = new ArrayList<>();
            if (measurementsJson != null && !measurementsJson.isEmpty()) {
                ObjectMapper objectMapper = new ObjectMapper();
                measurementList = objectMapper.readValue(measurementsJson, new TypeReference<List<Map<String, String>>>() {});
            }

            // 2. 拼接仪器字符串（自动换行，每个用两个空格隔开字段，支持多行）
            StringBuilder sb = new StringBuilder();
            for (Map<String, String> item : measurementList) {
                sb.append(item.getOrDefault("measurement", ""))
                        .append("  ")
                        .append(item.getOrDefault("certificateNo", ""))
                        .append("  ")
                        .append(item.getOrDefault("certificateDate", ""))
                        .append("\n");
            }
            if (sb.length() > 0) {
                replaceMap.put("measurement", sb.toString().trim());
            } else {
                replaceMap.put("measurement", "");
            }

            // 3. 传递到 service 层
            String outputPath = xlsService.processExcelFile(file, templateFile, images, replaceMap,measurementList);
            String fileName = outputPath.substring(outputPath.lastIndexOf(File.separator) + 1);
            String downloadUrl = "/download/" + fileName;
            result.put("downloadUrl", downloadUrl);
            result.put("code", 0);
            result.put("msg", "ok");
        } catch (Exception e) {
            log.error("上传失败", e);
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
