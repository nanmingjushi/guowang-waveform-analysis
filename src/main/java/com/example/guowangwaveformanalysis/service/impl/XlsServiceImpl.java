package com.example.guowangwaveformanalysis.service.impl;

import com.example.guowangwaveformanalysis.service.XlsService;
import lombok.Getter;
import org.apache.xmlbeans.XmlCursor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.util.Units;

import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

@Service
public class XlsServiceImpl implements XlsService {

    //日志对象
    private static Logger log = LoggerFactory.getLogger(XlsServiceImpl.class);
    //输出目录和文件名
    private static final String OUTPUT_DIR = "outputs";
    private static final String OUTPUT_FILE_NAME = "output.docx";

    @Override
    public String processExcelFile(MultipartFile excelFile, MultipartFile templateFile, MultipartFile[] images, Map<String, String> replaceMap, List<Map<String, String>> measurementList) throws Exception {
        try (InputStream excelStream = excelFile.getInputStream();
             InputStream templateStream = templateFile.getInputStream()) {
            ExcelSheetData data = parseExcelFromStream(excelStream);
            String outputPath = generateWordDocument(data, templateStream, images, replaceMap, measurementList);
            log.info("Word文档已生成：{}", outputPath);
            return outputPath;
        } catch (Exception e) {
            log.error("处理文件时发生异常: {}", e.getMessage(), e);
            throw e;
        }
    }


    //从Excel输入流中解析需要的数据，封装到ExcelSheetData对象
    private ExcelSheetData parseExcelFromStream(InputStream excelStream) {
        ExcelSheetData data = new ExcelSheetData();
        ExcelReader reader = ExcelUtil.getReader(excelStream);
        try {
            reader.setSheet("电压谐波");
            List<List<Object>> voltageData = reader.read();
            processMergedCells(voltageData, reader.getSheet());
            data.getVoltageHarmonicData().addAll(voltageData);

            reader.setSheet("电流谐波");
            List<List<Object>> currentData = reader.read();
            processMergedCells(currentData, reader.getSheet());
            data.getCurrentHarmonicData().addAll(currentData);

            reader.setSheet("功率");
            List<List<Object>> powerData = reader.read();
            processMergedCells(powerData, reader.getSheet());
            data.getPowerData().addAll(powerData);
        } finally {
            reader.close();
        }
        return data;
    }

    //处理Excel中的合并单元格，将其拆分或填充成适于后续处理的标准二维表数据
    private void processMergedCells(List<List<Object>> sheetData, Sheet sheet) {
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for (CellRangeAddress region : mergedRegions) {
            int firstRow = region.getFirstRow();
            int lastRow = region.getLastRow();
            int firstCol = region.getFirstColumn();
            int lastCol = region.getLastColumn();
            Object value = sheetData.get(firstRow).get(firstCol);
            for (int row = firstRow; row <= lastRow; row++) {
                for (int col = firstCol; col <= lastCol; col++) {
                    if (row != firstRow || col != firstCol) {
                        if (row < sheetData.size() && col < sheetData.get(row).size()) {
                            sheetData.get(row).set(col, value);
                        }
                    }
                }
            }
        }
    }

    // ======== 全局字段批量替换 =======
    private void replacePlaceholders(XWPFDocument doc, Map<String, String> replaceMap) {
        // 段落
        for (XWPFParagraph para : doc.getParagraphs()) {
            for (XWPFRun run : para.getRuns()) {
                String text = run.getText(0);
                if (text != null) {
                    for (Map.Entry<String, String> entry : replaceMap.entrySet()) {
                        if (text.contains("{{" + entry.getKey() + "}}")) {
                            text = text.replace("{{" + entry.getKey() + "}}", entry.getValue());
                        }
                    }
                    run.setText(text, 0);
                }
            }
        }
        // 表格
        for (XWPFTable table : doc.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph para : cell.getParagraphs()) {
                        for (XWPFRun run : para.getRuns()) {
                            String text = run.getText(0);
                            if (text != null) {
                                for (Map.Entry<String, String> entry : replaceMap.entrySet()) {
                                    if (text.contains("{{" + entry.getKey() + "}}")) {
                                        text = text.replace("{{" + entry.getKey() + "}}", entry.getValue());
                                    }
                                }
                                run.setText(text, 0);
                            }
                        }
                    }
                }
            }
        }
    }

    //根据解析后的数据 data 和 Word 模板 templateStream，生成最终输出的 Word 报告文档。
    private String generateWordDocument(
            ExcelSheetData data,
            InputStream templateStream,
            MultipartFile[] images,
            Map<String, String> replaceMap,
            List<Map<String, String>> measurementList
    ) throws IOException {
        XWPFDocument doc = new XWPFDocument(templateStream);

        // 拼接仪器参数，形如“仪器1 证书1 日期1\n仪器2 证书2 日期2”
        if (measurementList != null && !measurementList.isEmpty()) {
            StringBuilder sb = new StringBuilder();
            for (Map<String, String> item : measurementList) {
                String measurement = item.getOrDefault("measurement", "");
                String certNo = item.getOrDefault("certificateNo", "");
                String certDate = item.getOrDefault("certificateDate", "");
                sb.append(measurement).append("  ");
                sb.append(certNo).append("  ");
                sb.append(certDate).append("\n");
            }
            replaceMap.put("measurement", sb.toString().trim());
        } else {
            replaceMap.put("measurement", ""); // 防止占位符未被替换
        }

        // 替换所有 {{xxx}} 字段
        if (replaceMap != null && !replaceMap.isEmpty()) {
            System.out.println("替换字段：" + replaceMap);
            replacePlaceholders(doc, replaceMap);
        }

        // 获取监测位置
        String rawMonitorPosition = data.getVoltageHarmonicData().get(1).get(0).toString();
        String monitorPosition = rawMonitorPosition.replaceAll(".*[：:]", "").trim();

        // 假设模板中第一个表格是谐波电压表格，第二个表格是谐波电流表格
        List<XWPFTable> tables = doc.getTables();
        if (tables.size() >= 4) {

            // 谐波电压
            setTableTitle(doc, tables.get(0), "表1.1  " + monitorPosition + "谐波电压统计表");
            fillVoltageHarmonicTable(tables.get(0), data.getVoltageHarmonicData(),1.1,doc);
            // 谐波电流
            setTableTitle(doc, tables.get(1), "表1.2  " + monitorPosition + "谐波电流统计表");
            fillCurrentHarmonicTable(tables.get(1), data.getCurrentHarmonicData());
            //频率偏差、三相电压不平衡度及长时间闪变
            setTableTitle(doc, tables.get(2), "表1.3  " + monitorPosition + "频率偏差、三相电压不平衡度及长时间闪变统计表");
            fillFrequencyDeviationAndVoltageUnbalanceAndLongTermFlickerTable(tables.get(2), data.voltageHarmonicData,data.getPowerData());
            //电压偏差
            setTableTitle(doc, tables.get(3), "表1.4  " + monitorPosition + "电压偏差统计表");
            fillVoltageDeviationTable(tables.get(3), data.voltageHarmonicData, replaceMap);
            System.out.println("maxVoltageDeviation: " + replaceMap.get("maxVoltageDeviation"));

        }

        // 插入图片
        if (images != null) {
            for (int i = 0; i < images.length; i++) {
                String placeholder = "{{image" + (i + 1) + "}}";
                try (InputStream imgStream = images[i].getInputStream()) {
                    insertImageAndModifyCaption(
                            doc,
                            placeholder,
                            imgStream,
                            detectImageType(images[i].getOriginalFilename()),
                            400, 250,    // 这里宽高就是400*250像素
                            monitorPosition
                    );
                } catch (Exception e) {
                    log.warn("图片插入失败: {}", e.getMessage());
                }
            }
        }

        File outputDir = new File(OUTPUT_DIR);
        if (!outputDir.exists()) {
            outputDir.mkdirs();
        }
        String outputPath = OUTPUT_DIR + File.separator + OUTPUT_FILE_NAME;
        try (FileOutputStream fos = new FileOutputStream(outputPath)) {
            doc.write(fos);
        }
        return outputPath;
    }

    //谐波电压
    private void fillVoltageHarmonicTable(XWPFTable table, List<List<Object>> data,double tableNumber, XWPFDocument doc) {

        // 谐波电压数据从Excel读取
        double AB_average_fundamental = getDoubleValue(data.get(9).get(3));
        List<Double> AB_average_hruh_list = readColumnRange(data, 3, 10, 33);
        double AB_average_THD = getDoubleValue(data.get(59).get(3));

        double AB_95_fundamental = getDoubleValue(data.get(9).get(5));
        List<Double> AB_95_hruh_list = readColumnRange(data, 5, 10, 33);
        double AB_95_THD = getDoubleValue(data.get(59).get(5));

        double BC_average_fundamental = getDoubleValue(data.get(9).get(8));
        List<Double> BC_average_hruh_list = readColumnRange(data, 8, 10, 33);
        double BC_average_THD = getDoubleValue(data.get(59).get(8));

        double BC_95_fundamental = getDoubleValue(data.get(9).get(10));
        List<Double> BC_95_hruh_list = readColumnRange(data, 10, 10, 33);
        double BC_95_THD = getDoubleValue(data.get(59).get(10));

        double CA_average_fundamental = getDoubleValue(data.get(9).get(13));
        List<Double> CA_average_hruh_list = readColumnRange(data, 13, 10, 33);
        double CA_average_THD = getDoubleValue(data.get(59).get(13));

        double CA_95_fundamental = getDoubleValue(data.get(9).get(15));
        List<Double> CA_95_hruh_list = readColumnRange(data, 15, 10, 33);
        double CA_95_THD = getDoubleValue(data.get(59).get(15));

        List<Double> limit_hruh_list = readColumnRange(data, 17, 10, 33);
        double limit_THD = getDoubleValue(data.get(59).get(17));


        // 往docx填充数据
        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 2) {
            setCellText(table.getRows().get(2).getCell(1),formatDouble(AB_average_fundamental / 1000, 2));
        }
        for (int i = 0; i < AB_average_hruh_list.size(); i++) {
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 2) {
                setCellText(table.getRows().get(i + 3).getCell(2),formatDouble(AB_average_hruh_list.get(i), 2));
            }
        }
        if (table.getRows().size() > 27 && table.getRows().get(27).getTableCells().size() > 2) {
            setCellText(table.getRows().get(27).getCell(1),formatDouble(AB_average_THD, 2));
        }

        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 3) {
            setCellText(table.getRows().get(2).getCell(2),formatDouble(AB_95_fundamental / 1000, 2));
        }
        for (int i = 0; i < AB_95_hruh_list.size(); i++) {
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 3) {
                setCellText(table.getRows().get(i + 3).getCell(3),formatDouble(AB_95_hruh_list.get(i), 2));
            }
        }
        if (table.getRows().size() > 27 && table.getRows().get(27).getTableCells().size() > 3) {
            setCellText(table.getRows().get(27).getCell(2),formatDouble(AB_95_THD, 2));
        }

        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 4) {
            setCellText(table.getRows().get(2).getCell(3),formatDouble(BC_average_fundamental / 1000, 2));
        }
        for (int i = 0; i < BC_average_hruh_list.size(); i++) {
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 4) {
                setCellText(table.getRows().get(i + 3).getCell(4),formatDouble(BC_average_hruh_list.get(i), 2));
            }
        }
        if (table.getRows().size() > 27 && table.getRows().get(27).getTableCells().size() > 4) {
            setCellText(table.getRows().get(27).getCell(3),formatDouble(BC_average_THD, 2));
        }

        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 5) {
            setCellText(table.getRows().get(2).getCell(4),formatDouble(BC_95_fundamental / 1000, 2));
        }
        for (int i = 0; i < BC_95_hruh_list.size(); i++) {
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 5) {
                setCellText(table.getRows().get(i + 3).getCell(5),formatDouble(BC_95_hruh_list.get(i), 2));
            }
        }
        if (table.getRows().size() > 27 && table.getRows().get(27).getTableCells().size() > 5) {
            setCellText(table.getRows().get(27).getCell(4),formatDouble(BC_95_THD, 2));
        }

        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 6) {
            setCellText(table.getRows().get(2).getCell(5),formatDouble(CA_average_fundamental / 1000, 2));
        }
        for (int i = 0; i < CA_average_hruh_list.size(); i++) {
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 6) {
                setCellText(table.getRows().get(i + 3).getCell(6),formatDouble(CA_average_hruh_list.get(i), 2));
            }
        }
        if (table.getRows().size() > 27 && table.getRows().get(27).getTableCells().size() > 6) {
            setCellText(table.getRows().get(27).getCell(5),formatDouble(CA_average_THD, 2));
        }

        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 7) {
            setCellText(table.getRows().get(2).getCell(6),formatDouble(CA_95_fundamental / 1000, 2));
        }
        for (int i = 0; i < CA_95_hruh_list.size(); i++) {
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 7) {
                setCellText(table.getRows().get(i + 3).getCell(7),formatDouble(CA_95_hruh_list.get(i), 2));
            }
        }
        if (table.getRows().size() > 27 && table.getRows().get(27).getTableCells().size() > 7) {
            setCellText(table.getRows().get(27).getCell(6),formatDouble(CA_95_THD, 2));
        }

        // hruh 限值
        for (int i = 0; i < limit_hruh_list.size(); i++) {
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 8) {
                setCellText(table.getRows().get(i + 3).getCell(8),formatDouble(limit_hruh_list.get(i), 2));
            }
        }
        //  基波电压限值
        setCellText(table.getRows().get(2).getCell(7),("—"));

        // thd 限值
        setCellText(table.getRows().get(27).getCell(7),formatDouble(limit_THD, 2));
    }

    //谐波电流
    private void fillCurrentHarmonicTable(XWPFTable table, List<List<Object>> data) {

        // 1. 数据提取
        double A_average_fundamental = getDoubleValue(data.get(9).get(3));
        List<Double> A_average_hruh_list = readColumnRange(data, 3, 10, 33);
        double A_average_THD = getDoubleValue(data.get(59).get(3));

        double A_95_fundamental = getDoubleValue(data.get(9).get(5));
        List<Double> A_95_hruh_list = readColumnRange(data, 5, 10, 33);
        double A_95_THD = getDoubleValue(data.get(59).get(5));

        double B_average_fundamental = getDoubleValue(data.get(9).get(8));
        List<Double> B_average_hruh_list = readColumnRange(data, 8, 10, 33);
        double B_average_THD = getDoubleValue(data.get(59).get(8));

        double B_95_fundamental = getDoubleValue(data.get(9).get(10));
        List<Double> B_95_hruh_list = readColumnRange(data, 10, 10, 33);
        double B_95_THD = getDoubleValue(data.get(59).get(10));

        double C_average_fundamental = getDoubleValue(data.get(9).get(13));
        List<Double> C_average_hruh_list = readColumnRange(data, 13, 10, 33);
        double C_average_THD = getDoubleValue(data.get(59).get(13));

        double C_95_fundamental = getDoubleValue(data.get(9).get(15));
        List<Double> C_95_hruh_list = readColumnRange(data, 15, 10, 33);
        double C_95_THD = getDoubleValue(data.get(59).get(15));

        List<Double> limit_hruh_list = readColumnRange(data, 17, 10, 33);
        double limit_THD = getDoubleValue(data.get(59).get(17));

        // 2.填充数据
        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 2) {
            setCellText(table.getRows().get(2).getCell(1),formatDouble(A_average_fundamental, 2));
        }
        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 3) {
            setCellText(table.getRows().get(2).getCell(2),formatDouble(A_95_fundamental, 2));
        }
        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 4) {
            setCellText(table.getRows().get(2).getCell(3),formatDouble(B_average_fundamental, 2));
        }
        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 5) {
            setCellText(table.getRows().get(2).getCell(4),formatDouble(B_95_fundamental, 2));
        }
        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 6) {
            setCellText(table.getRows().get(2).getCell(5),formatDouble(C_average_fundamental, 2));
        }
        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 7) {
            setCellText(table.getRows().get(2).getCell(6),formatDouble(C_95_fundamental, 2));
        }

        for (int i = 0; i < 24; i++) {
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 2) {
                setCellText(table.getRows().get(i + 3).getCell(2),formatDouble(A_average_hruh_list.get(i), 2));
            }
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 3) {
                setCellText(table.getRows().get(i + 3).getCell(3),formatDouble(A_95_hruh_list.get(i), 2));
            }
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 4) {
                setCellText(table.getRows().get(i + 3).getCell(4),formatDouble(B_average_hruh_list.get(i), 2));
            }
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 5) {
                setCellText(table.getRows().get(i + 3).getCell(5),formatDouble(B_95_hruh_list.get(i), 2));
            }
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 6) {
                setCellText(table.getRows().get(i + 3).getCell(6),formatDouble(C_average_hruh_list.get(i), 2));
            }
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 7) {
                setCellText(table.getRows().get(i + 3).getCell(7),formatDouble(C_95_hruh_list.get(i), 2));
            }
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 8) {
                setCellText(table.getRows().get(i + 3).getCell(8),formatDouble(limit_hruh_list.get(i), 2));
            }
        }
        //  基波电流限值
        setCellText(table.getRows().get(2).getCell(7),("—"));
    }

    //频率偏差、三相电压不平衡度及长时间闪变
    private void fillFrequencyDeviationAndVoltageUnbalanceAndLongTermFlickerTable(XWPFTable table, List<List<Object>> voltageHarmonicData, List<List<Object>> powerData) {

        // 1. 数据提取
        double frequency_max = getDoubleValue(powerData.get(15).get(2));
        double frequency_average = getDoubleValue(powerData.get(15).get(3));
        double frequency_min = getDoubleValue(powerData.get(15).get(4));
        double frequency_95 = getDoubleValue(powerData.get(15).get(5));
        String frequency_limit = powerData.get(15).get(17).toString();

        double voltage_unbalance_max = getDoubleValue(powerData.get(16).get(2));
        double voltage_unbalance_average = getDoubleValue(powerData.get(16).get(3));
        double voltage_unbalance_min = getDoubleValue(powerData.get(16).get(4));
        double voltage_unbalance_95 = getDoubleValue(powerData.get(16).get(5));
        double voltage_unbalance_limit = getDoubleValue(powerData.get(16).get(17));

        double long_term_flicker_AB_max = getDoubleValue(voltageHarmonicData.get(61).get(2));
        double long_term_flicker_AB_average = getDoubleValue(voltageHarmonicData.get(61).get(3));
        double long_term_flicker_AB_min = getDoubleValue(voltageHarmonicData.get(61).get(4));
        double long_term_flicker_AB_95 = getDoubleValue(voltageHarmonicData.get(61).get(5));

        double long_term_flicker_BC_max = getDoubleValue(voltageHarmonicData.get(61).get(7));
        double long_term_flicker_BC_average = getDoubleValue(voltageHarmonicData.get(61).get(8));
        double long_term_flicker_BC_min = getDoubleValue(voltageHarmonicData.get(61).get(9));
        double long_term_flicker_BC_95 = getDoubleValue(voltageHarmonicData.get(61).get(10));

        double long_term_flicker_AC_max = getDoubleValue(voltageHarmonicData.get(61).get(12));
        double long_term_flicker_AC_average = getDoubleValue(voltageHarmonicData.get(61).get(13));
        double long_term_flicker_AC_min = getDoubleValue(voltageHarmonicData.get(61).get(14));
        double long_term_flicker_AC_95 = getDoubleValue(voltageHarmonicData.get(61).get(15));

        double long_term_flicker_limit = getDoubleValue(voltageHarmonicData.get(61).get(17));

        // 2. 填充数据
        //2.1 填充频率偏差数据
        setCellText(table.getRow(1).getCell(1),formatDouble(frequency_max-50, 2));
        setCellText(table.getRow(1).getCell(2),formatDouble(frequency_average-50, 2));
        setCellText(table.getRow(1).getCell(3),formatDouble(frequency_min-50, 2));
        setCellText(table.getRow(1).getCell(4),formatDouble(frequency_95-50, 2));
        setCellText(table.getRow(1).getCell(5),frequency_limit);
        //2.2 填充三相电压不平衡度数据
        setCellText(table.getRow(2).getCell(1),formatDouble(voltage_unbalance_max, 2));
        setCellText(table.getRow(2).getCell(2),formatDouble(voltage_unbalance_average, 2));
        setCellText(table.getRow(2).getCell(3),formatDouble(voltage_unbalance_min, 2));
        setCellText(table.getRow(2).getCell(4),formatDouble(voltage_unbalance_95, 2));
        setCellText(table.getRow(2).getCell(5),formatDouble(voltage_unbalance_limit, 2));

        //2.3 填充长时间闪变数据
        // AB相闪变数据
        setCellText(table.getRow(3).getCell(2),formatDouble(long_term_flicker_AB_max, 2));
        setCellText(table.getRow(3).getCell(3),formatDouble(long_term_flicker_AB_average, 2));
        setCellText(table.getRow(3).getCell(4),formatDouble(long_term_flicker_AB_min, 2));
        setCellText(table.getRow(3).getCell(5),formatDouble(long_term_flicker_AB_95, 2));
        setCellText(table.getRow(3).getCell(6),formatDouble(long_term_flicker_limit, 2));

        // BC相闪变数据
        setCellText(table.getRow(4).getCell(2),formatDouble(long_term_flicker_BC_max, 2));
        setCellText(table.getRow(4).getCell(3),formatDouble(long_term_flicker_BC_average, 2));
        setCellText(table.getRow(4).getCell(4),formatDouble(long_term_flicker_BC_min, 2));
        setCellText(table.getRow(4).getCell(5),formatDouble(long_term_flicker_BC_95, 2));
        setCellText(table.getRow(4).getCell(6),formatDouble(long_term_flicker_limit, 2));

        // AC相闪变数据
        setCellText(table.getRow(5).getCell(2),formatDouble(long_term_flicker_AC_max, 2));
        setCellText(table.getRow(5).getCell(3),formatDouble(long_term_flicker_AC_average, 2));
        setCellText(table.getRow(5).getCell(4),formatDouble(long_term_flicker_AC_min, 2));
        setCellText(table.getRow(5).getCell(5),formatDouble(long_term_flicker_AC_95, 2));
        setCellText(table.getRow(5).getCell(6),formatDouble(long_term_flicker_limit, 2));
    }

    //电压偏差
    private void fillVoltageDeviationTable(XWPFTable table, List<List<Object>> voltageHarmonicData, Map<String, String> replaceMap) {

        // 1.1 数据提取（上偏差）
        double voltage_deviation_up_AB_max = getDoubleValue(voltageHarmonicData.get(63).get(2));
        double voltage_deviation_up_AB_min = getDoubleValue(voltageHarmonicData.get(63).get(4));
        double voltage_deviation_up_BC_max = getDoubleValue(voltageHarmonicData.get(63).get(7));
        double voltage_deviation_up_BC_min = getDoubleValue(voltageHarmonicData.get(63).get(9));
        double voltage_deviation_up_AC_max = getDoubleValue(voltageHarmonicData.get(63).get(12));
        double voltage_deviation_up_AC_min = getDoubleValue(voltageHarmonicData.get(63).get(14));
        double voltage_deviation_up_limit = getDoubleValue(voltageHarmonicData.get(63).get(17));

        // 1.2 数据提取（下偏差）
        double voltage_deviation_down_AB_max = getDoubleValue(voltageHarmonicData.get(64).get(2));
        double voltage_deviation_down_AB_min = getDoubleValue(voltageHarmonicData.get(64).get(4));
        double voltage_deviation_down_BC_max = getDoubleValue(voltageHarmonicData.get(64).get(7));
        double voltage_deviation_down_BC_min = getDoubleValue(voltageHarmonicData.get(64).get(9));
        double voltage_deviation_down_AC_max = getDoubleValue(voltageHarmonicData.get(64).get(12));
        double voltage_deviation_down_AC_min = getDoubleValue(voltageHarmonicData.get(64).get(14));
        double voltage_deviation_down_limit = Double.parseDouble("-"+voltageHarmonicData.get(64).get(17).toString());

        // 2. 填充数据
        // 上偏差数据
        setCellText(table.getRow(2).getCell(1),formatDouble(voltage_deviation_up_AB_max, 2));
        setCellText(table.getRow(2).getCell(2),formatDouble(voltage_deviation_up_AB_min, 2));
        setCellText(table.getRow(2).getCell(3),formatDouble(voltage_deviation_up_BC_max, 2));
        setCellText(table.getRow(2).getCell(4),formatDouble(voltage_deviation_up_BC_min, 2));
        setCellText(table.getRow(2).getCell(5),formatDouble(voltage_deviation_up_AC_max, 2));
        setCellText(table.getRow(2).getCell(6),formatDouble(voltage_deviation_up_AC_min, 2));
        setCellText(table.getRow(2).getCell(7),formatDouble(voltage_deviation_up_limit, 2));

        // 下偏差数据
        setCellText(table.getRow(3).getCell(1),formatDouble(voltage_deviation_down_AB_max, 2));
        setCellText(table.getRow(3).getCell(2),formatDouble(voltage_deviation_down_AB_min, 2));
        setCellText(table.getRow(3).getCell(3),formatDouble(voltage_deviation_down_BC_max, 2));
        setCellText(table.getRow(3).getCell(4),formatDouble(voltage_deviation_down_BC_min, 2));
        setCellText(table.getRow(3).getCell(5),formatDouble(voltage_deviation_down_AC_max, 2));
        setCellText(table.getRow(3).getCell(6),formatDouble(voltage_deviation_down_AC_min, 2));
        setCellText(table.getRow(3).getCell(7),formatDouble(voltage_deviation_down_limit, 2));

        // 3. 找出上偏差中 AB/BC/AC 最大者
        double max_voltage_deviation = Math.max(voltage_deviation_up_AB_max, Math.max(voltage_deviation_up_BC_max, voltage_deviation_up_AC_max));
        replaceMap.put("maxVoltageDeviation", formatDouble(max_voltage_deviation,2));


    }


    @Getter
    private static class ExcelSheetData {
        private final List<List<Object>> voltageHarmonicData = new ArrayList<>();
        private final List<List<Object>> currentHarmonicData = new ArrayList<>();
        private final List<List<Object>> powerData = new ArrayList<>();
    }

    // 读取 Excel 列范围数据
    private List<Double> readColumnRange(List<List<Object>> data, int colIndex, int startRow, int endRow) {
        List<Double> result = new ArrayList<>();
        for (int i = startRow; i <= endRow; i++) {
            result.add(getDoubleValue(data.get(i).get(colIndex)));
        }
        return result;
    }

    // 获取对象的 double 值
    private double getDoubleValue(Object obj) {
        if (obj == null) {
            return 0.0;
        }
        try {
            return Double.parseDouble(obj.toString());
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }

    // 格式化 double 类型数据
    private String formatDouble(double value, int scale) {
        BigDecimal bd = new BigDecimal(value);
        bd = bd.setScale(scale, BigDecimal.ROUND_HALF_UP);
        return bd.toString();
    }

    //给docx中的表格添加标题
    private void setTableTitle(XWPFDocument doc, XWPFTable table, String title) {
        List<IBodyElement> bodyElements = doc.getBodyElements();
        for (int i = 0; i < bodyElements.size(); i++) {
            if (bodyElements.get(i).getElementType() == BodyElementType.TABLE && table == (XWPFTable)bodyElements.get(i)) {
                if (i > 0 && bodyElements.get(i-1).getElementType() == BodyElementType.PARAGRAPH) {
                    XWPFParagraph para = (XWPFParagraph) bodyElements.get(i-1);
                    // 安全移除所有 runs
                    int runCount = para.getRuns().size();
                    for (int j = runCount - 1; j >= 0; j--) {
                        para.removeRun(j);
                    }
                    XWPFRun run = para.createRun();
                    run.setText(title);
                    run.setFontFamily("SimSun"); // 宋体
                    run.setFontSize(12);         // 小四（12磅）
                } else {
                    // 没有标题段落，需新插入一个（用XmlCursor）
                    XmlCursor cursor = table.getCTTbl().newCursor();
                    XWPFParagraph newPara = doc.insertNewParagraph(cursor);
                    XWPFRun run = newPara.createRun();
                    run.setText(title);
                    run.setFontFamily("SimSun");
                    run.setFontSize(12);
                }
                break;
            }
        }
    }

    // 设置docx单元格文本格式
    private void setCellText(XWPFTableCell cell, String text) {
        // 清空所有段落
        while (cell.getParagraphs().size() > 0) {
            cell.removeParagraph(0);
        }
        XWPFParagraph para = cell.addParagraph();
        para.setAlignment(ParagraphAlignment.CENTER); // 居中对齐
        XWPFRun run = para.createRun();
        run.setText(text);
        run.setFontFamily("Times New Roman"); // 设置为 Times New Roman
        run.setFontSize(10);         // 小五
    }

    // 自动判断图片类型
    private int detectImageType(String filename) {
        if (filename == null) return XWPFDocument.PICTURE_TYPE_PNG;
        String lower = filename.toLowerCase();
        if (lower.endsWith(".png")) return XWPFDocument.PICTURE_TYPE_PNG;
        if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return XWPFDocument.PICTURE_TYPE_JPEG;
        if (lower.endsWith(".bmp")) return XWPFDocument.PICTURE_TYPE_BMP;
        if (lower.endsWith(".gif")) return XWPFDocument.PICTURE_TYPE_GIF;
        return XWPFDocument.PICTURE_TYPE_PNG; // 默认
    }

    // 替换占位符为图片
    private void insertImageAndModifyCaption(
            XWPFDocument doc,
            String placeholder,
            InputStream imageStream,
            int imageType,
            int widthPx,
            int heightPx,
            String monitorPosition
    ) throws Exception {
        List<XWPFParagraph> paragraphs = doc.getParagraphs();
        for (int i = 0; i < paragraphs.size(); i++) {
            XWPFParagraph para = paragraphs.get(i);
            String text = para.getText();
            if (text != null && text.contains(placeholder)) {
                // 清空占位符段落并插入图片
                int runCount = para.getRuns().size();
                for (int j = runCount - 1; j >= 0; j--) para.removeRun(j);
                XWPFRun run = para.createRun();
                run.addPicture(imageStream, imageType, "image",
                        Units.toEMU(widthPx), Units.toEMU(heightPx));

                // 自动替换下一个段落（图名）
                if (i + 1 < paragraphs.size()) {
                    XWPFParagraph captionPara = paragraphs.get(i + 1);
                    String origCaption = captionPara.getText();
                    String newCaption = origCaption;
                    if (origCaption != null) {
                        // 只判断开头为“图1.1”，其余均加监测位置
                        if (origCaption.trim().startsWith("图1.1")) {
                            newCaption = origCaption;
                        } else {
                            int firstSpace = origCaption.indexOf(' ');
                            if (firstSpace != -1 && firstSpace + 1 < origCaption.length()) {
                                newCaption = origCaption.substring(0, firstSpace + 1)
                                        + " " + monitorPosition
                                        + origCaption.substring(firstSpace + 1);
                            } else {
                                newCaption = origCaption + " " + monitorPosition;
                            }
                        }
                    }
                    int capRunCount = captionPara.getRuns().size();
                    for (int j = capRunCount - 1; j >= 0; j--) captionPara.removeRun(j);
                    XWPFRun capRun = captionPara.createRun();
                    capRun.setText(newCaption);
                    capRun.setFontFamily("SimSun");
                    capRun.setFontSize(12); // 小四
                }

                return;
            }
        }
    }

}