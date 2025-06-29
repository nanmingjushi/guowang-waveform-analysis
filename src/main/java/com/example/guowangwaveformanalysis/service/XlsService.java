package com.example.guowangwaveformanalysis.service;

import org.springframework.web.multipart.MultipartFile;
import java.util.Map;
import java.util.List;

public interface XlsService {
    /**
     * 生成报告
     * @param file           Excel文件
     * @param templateFile   Word模板
     * @param images         图片数组
     * @param replaceMap     需要替换的基本字段（String-String）
     * @param measurementList 仪器列表（每个Map可包含 measurement、certificateNo、certificateDate）
     * @return 输出Word路径
     * @throws Exception 异常
     */
    String processExcelFile(
            MultipartFile file,
            MultipartFile templateFile,
            MultipartFile[] images,
            Map<String, String> replaceMap,
            List<Map<String, String>> measurementList
    ) throws Exception;
}
