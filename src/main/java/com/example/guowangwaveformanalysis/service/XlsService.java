package com.example.guowangwaveformanalysis.service;

import org.springframework.web.multipart.MultipartFile;

/**
 * @author nan chao
 * @since 2025/4/8 11:16
 */


public interface XlsService {
    String processExcelFile(MultipartFile file, MultipartFile templateFile) throws Exception;

}
