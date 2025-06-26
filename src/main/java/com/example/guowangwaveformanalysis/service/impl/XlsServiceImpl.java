package com.example.guowangwaveformanalysis.service.impl;

import com.example.guowangwaveformanalysis.service.XlsService;
import lombok.Getter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

@Service
public class XlsServiceImpl implements XlsService {

    private static Logger log = LoggerFactory.getLogger(XlsServiceImpl.class);

    private static final String OUTPUT_DIR = "outputs";
    private static final String OUTPUT_FILE_NAME = "output.docx";

    @Override
    public String processExcelFile(MultipartFile excelFile, MultipartFile templateFile) throws Exception {
        try (InputStream excelStream = excelFile.getInputStream();
             InputStream templateStream = templateFile.getInputStream()) {
            ExcelSheetData data = parseExcelFromStream(excelStream);
            String outputPath = generateWordDocument(data, templateStream);
            log.info("Word文档已生成：{}", outputPath);
            return outputPath;
        } catch (Exception e) {
            log.error("处理文件时发生异常: {}", e.getMessage(), e);
            throw e;
        }
    }

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

    private String generateWordDocument(ExcelSheetData data, InputStream templateStream) throws IOException {
        XWPFDocument doc = new XWPFDocument(templateStream);

        // 获取监测位置
        String monitorPosition = data.getVoltageHarmonicData().get(1).get(0).toString();

        // 假设模板中第一个表格是谐波电压表格，第二个表格是谐波电流表格
        List<XWPFTable> tables = doc.getTables();
        if (tables.size() >= 2) {

            // 填充谐波电压表格
            fillVoltageHarmonicTable(tables.get(0), data.getVoltageHarmonicData(),1.1,doc);
            // 填充谐波电流表格
            fillCurrentHarmonicTable(tables.get(1), data.getCurrentHarmonicData());
            //填充频率偏差、三相电压不平衡度及长时间闪变
            fillFrequencyDeviationAndVoltageUnbalanceAndLongTermFlickerTable(tables.get(2), data.voltageHarmonicData,data.getPowerData());
            //填充电压偏差
            fillVoltageDeviationTable(tables.get(3), data.voltageHarmonicData);
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
        // 监测位置
        String monitorPosition = data.get(1).get(0).toString();


        // 谐波电压表格数据读取
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


        // 填充数据
        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 2) {
            table.getRows().get(2).getCell(1).setText(formatDouble(AB_average_fundamental / 1000, 2));
        }
        for (int i = 0; i < AB_average_hruh_list.size(); i++) {
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 2) {
                table.getRows().get(i + 3).getCell(2).setText(formatDouble(AB_average_hruh_list.get(i), 2));
            }
        }
        if (table.getRows().size() > 27 && table.getRows().get(27).getTableCells().size() > 2) {
            table.getRows().get(27).getCell(1).setText(formatDouble(AB_average_THD, 2));
        }

        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 3) {
            table.getRows().get(2).getCell(2).setText(formatDouble(AB_95_fundamental / 1000, 2));
        }
        for (int i = 0; i < AB_95_hruh_list.size(); i++) {
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 3) {
                table.getRows().get(i + 3).getCell(3).setText(formatDouble(AB_95_hruh_list.get(i), 2));
            }
        }
        if (table.getRows().size() > 27 && table.getRows().get(27).getTableCells().size() > 3) {
            table.getRows().get(27).getCell(2).setText(formatDouble(AB_95_THD, 2));
        }

        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 4) {
            table.getRows().get(2).getCell(3).setText(formatDouble(BC_average_fundamental / 1000, 2));
        }
        for (int i = 0; i < BC_average_hruh_list.size(); i++) {
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 4) {
                table.getRows().get(i + 3).getCell(4).setText(formatDouble(BC_average_hruh_list.get(i), 2));
            }
        }
        if (table.getRows().size() > 27 && table.getRows().get(27).getTableCells().size() > 4) {
            table.getRows().get(27).getCell(3).setText(formatDouble(BC_average_THD, 2));
        }

        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 5) {
            table.getRows().get(2).getCell(4).setText(formatDouble(BC_95_fundamental / 1000, 2));
        }
        for (int i = 0; i < BC_95_hruh_list.size(); i++) {
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 5) {
                table.getRows().get(i + 3).getCell(5).setText(formatDouble(BC_95_hruh_list.get(i), 2));
            }
        }
        if (table.getRows().size() > 27 && table.getRows().get(27).getTableCells().size() > 5) {
            table.getRows().get(27).getCell(4).setText(formatDouble(BC_95_THD, 2));
        }

        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 6) {
            table.getRows().get(2).getCell(5).setText(formatDouble(CA_average_fundamental / 1000, 2));
        }
        for (int i = 0; i < CA_average_hruh_list.size(); i++) {
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 6) {
                table.getRows().get(i + 3).getCell(6).setText(formatDouble(CA_average_hruh_list.get(i), 2));
            }
        }
        if (table.getRows().size() > 27 && table.getRows().get(27).getTableCells().size() > 6) {
            table.getRows().get(27).getCell(5).setText(formatDouble(CA_average_THD, 2));
        }

        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 7) {
            table.getRows().get(2).getCell(6).setText(formatDouble(CA_95_fundamental / 1000, 2));
        }
        for (int i = 0; i < CA_95_hruh_list.size(); i++) {
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 7) {
                table.getRows().get(i + 3).getCell(7).setText(formatDouble(CA_95_hruh_list.get(i), 2));
            }
        }
        if (table.getRows().size() > 27 && table.getRows().get(27).getTableCells().size() > 7) {
            table.getRows().get(27).getCell(6).setText(formatDouble(CA_95_THD, 2));
        }

        // hruh 限值
        for (int i = 0; i < limit_hruh_list.size(); i++) {
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 8) {
                table.getRows().get(i + 3).getCell(8).setText(formatDouble(limit_hruh_list.get(i), 2));
            }
        }
        //  基波电压限值
        table.getRows().get(2).getCell(7).setText("—");

        // thd 限值
        table.getRows().get(27).getCell(7).setText(formatDouble(limit_THD, 2));
    }

    //谐波电流
    private void fillCurrentHarmonicTable(XWPFTable table, List<List<Object>> data) {
        // 1. 基础数据准备
        String monitorPosition = data.get(1).get(0).toString();

        // 2. 数据提取
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

        // 填充数据
        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 2) {
            table.getRows().get(2).getCell(1).setText(formatDouble(A_average_fundamental, 2));
        }
        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 3) {
            table.getRows().get(2).getCell(2).setText(formatDouble(A_95_fundamental, 2));
        }
        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 4) {
            table.getRows().get(2).getCell(3).setText(formatDouble(B_average_fundamental, 2));
        }
        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 5) {
            table.getRows().get(2).getCell(4).setText(formatDouble(B_95_fundamental, 2));
        }
        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 6) {
            table.getRows().get(2).getCell(5).setText(formatDouble(C_average_fundamental, 2));
        }
        if (table.getRows().size() > 2 && table.getRows().get(2).getTableCells().size() > 7) {
            table.getRows().get(2).getCell(6).setText(formatDouble(C_95_fundamental, 2));
        }

        for (int i = 0; i < 24; i++) {
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 2) {
                table.getRows().get(i + 3).getCell(2).setText(formatDouble(A_average_hruh_list.get(i), 2));
            }
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 3) {
                table.getRows().get(i + 3).getCell(3).setText(formatDouble(A_95_hruh_list.get(i), 2));
            }
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 4) {
                table.getRows().get(i + 3).getCell(4).setText(formatDouble(B_average_hruh_list.get(i), 2));
            }
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 5) {
                table.getRows().get(i + 3).getCell(5).setText(formatDouble(B_95_hruh_list.get(i), 2));
            }
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 6) {
                table.getRows().get(i + 3).getCell(6).setText(formatDouble(C_average_hruh_list.get(i), 2));
            }
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 7) {
                table.getRows().get(i + 3).getCell(7).setText(formatDouble(C_95_hruh_list.get(i), 2));
            }
            if (table.getRows().size() > i + 3 && table.getRows().get(i + 3).getTableCells().size() > 8) {
                table.getRows().get(i + 3).getCell(8).setText(formatDouble(limit_hruh_list.get(i), 2));
            }
        }
        //  基波电流限值
        table.getRows().get(2).getCell(7).setText("—");
    }

    //频率偏差、三相电压不平衡度及长时间闪变
    private void fillFrequencyDeviationAndVoltageUnbalanceAndLongTermFlickerTable(XWPFTable table, List<List<Object>> voltageHarmonicData, List<List<Object>> powerData) {
        // 1. 基础数据准备
        String monitorPosition = voltageHarmonicData.get(1).get(0).toString();

        // 2. 数据提取
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

        // 3. 填充数据
        //3.1 填充频率偏差数据
        table.getRow(1).getCell(1).setText(formatDouble(frequency_max, 2));
        table.getRow(1).getCell(2).setText(formatDouble(frequency_average, 2));
        table.getRow(1).getCell(3).setText(formatDouble(frequency_min, 2));
        table.getRow(1).getCell(4).setText(formatDouble(frequency_95, 2));
        table.getRow(1).getCell(5).setText(frequency_limit);
        //3.2 填充三相电压不平衡度数据
        table.getRow(2).getCell(1).setText(formatDouble(voltage_unbalance_max, 2));
        table.getRow(2).getCell(2).setText(formatDouble(voltage_unbalance_average, 2));
        table.getRow(2).getCell(3).setText(formatDouble(voltage_unbalance_min, 2));
        table.getRow(2).getCell(4).setText(formatDouble(voltage_unbalance_95, 2));
        table.getRow(2).getCell(5).setText(formatDouble(voltage_unbalance_limit, 2));

        //3.3 填充长时间闪变数据
        // AB相闪变数据
        table.getRow(3).getCell(2).setText(formatDouble(long_term_flicker_AB_max, 2));
        table.getRow(3).getCell(3).setText(formatDouble(long_term_flicker_AB_average, 2));
        table.getRow(3).getCell(4).setText(formatDouble(long_term_flicker_AB_min, 2));
        table.getRow(3).getCell(5).setText(formatDouble(long_term_flicker_AB_95, 2));
        table.getRow(3).getCell(6).setText(formatDouble(long_term_flicker_limit, 2));

        // BC相闪变数据
        table.getRow(4).getCell(2).setText(formatDouble(long_term_flicker_BC_max, 2));
        table.getRow(4).getCell(3).setText(formatDouble(long_term_flicker_BC_average, 2));
        table.getRow(4).getCell(4).setText(formatDouble(long_term_flicker_BC_min, 2));
        table.getRow(4).getCell(5).setText(formatDouble(long_term_flicker_BC_95, 2));
        table.getRow(4).getCell(6).setText(formatDouble(long_term_flicker_limit, 2));

        // AC相闪变数据
        table.getRow(5).getCell(2).setText(formatDouble(long_term_flicker_AC_max, 2));
        table.getRow(5).getCell(3).setText(formatDouble(long_term_flicker_AC_average, 2));
        table.getRow(5).getCell(4).setText(formatDouble(long_term_flicker_AC_min, 2));
        table.getRow(5).getCell(5).setText(formatDouble(long_term_flicker_AC_95, 2));
        table.getRow(5).getCell(6).setText(formatDouble(long_term_flicker_limit, 2));
    }

    //电压偏差
    private void fillVoltageDeviationTable(XWPFTable table, List<List<Object>> voltageHarmonicData) {
        // 1. 基础数据准备
        String monitorPosition = voltageHarmonicData.get(1).get(0).toString();

        // 2.1 数据提取（上偏差）
        double voltage_deviation_up_AB_max = getDoubleValue(voltageHarmonicData.get(63).get(2));
        double voltage_deviation_up_AB_min = getDoubleValue(voltageHarmonicData.get(63).get(4));
        double voltage_deviation_up_BC_max = getDoubleValue(voltageHarmonicData.get(63).get(7));
        double voltage_deviation_up_BC_min = getDoubleValue(voltageHarmonicData.get(63).get(9));
        double voltage_deviation_up_AC_max = getDoubleValue(voltageHarmonicData.get(63).get(12));
        double voltage_deviation_up_AC_min = getDoubleValue(voltageHarmonicData.get(63).get(14));
        double voltage_deviation_up_limit = getDoubleValue(voltageHarmonicData.get(63).get(17));

        // 2.2 数据提取（下偏差）
        double voltage_deviation_down_AB_max = getDoubleValue(voltageHarmonicData.get(64).get(2));
        double voltage_deviation_down_AB_min = getDoubleValue(voltageHarmonicData.get(64).get(4));
        double voltage_deviation_down_BC_max = getDoubleValue(voltageHarmonicData.get(64).get(7));
        double voltage_deviation_down_BC_min = getDoubleValue(voltageHarmonicData.get(64).get(9));
        double voltage_deviation_down_AC_max = getDoubleValue(voltageHarmonicData.get(64).get(12));
        double voltage_deviation_down_AC_min = getDoubleValue(voltageHarmonicData.get(64).get(14));
        double voltage_deviation_down_limit = Double.parseDouble("-"+voltageHarmonicData.get(64).get(17).toString());

        // 3. 填充数据
        // 上偏差数据
        table.getRow(2).getCell(1).setText(formatDouble(voltage_deviation_up_AB_max, 2));
        table.getRow(2).getCell(2).setText(formatDouble(voltage_deviation_up_AB_min, 2));
        table.getRow(2).getCell(3).setText(formatDouble(voltage_deviation_up_BC_max, 2));
        table.getRow(2).getCell(4).setText(formatDouble(voltage_deviation_up_BC_min, 2));
        table.getRow(2).getCell(5).setText(formatDouble(voltage_deviation_up_AC_max, 2));
        table.getRow(2).getCell(6).setText(formatDouble(voltage_deviation_up_AC_min, 2));
        table.getRow(2).getCell(7).setText(formatDouble(voltage_deviation_up_limit, 2));

        // 下偏差数据
        table.getRow(3).getCell(1).setText(formatDouble(voltage_deviation_down_AB_max, 2));
        table.getRow(3).getCell(2).setText(formatDouble(voltage_deviation_down_AB_min, 2));
        table.getRow(3).getCell(3).setText(formatDouble(voltage_deviation_down_BC_max, 2));
        table.getRow(3).getCell(4).setText(formatDouble(voltage_deviation_down_BC_min, 2));
        table.getRow(3).getCell(5).setText(formatDouble(voltage_deviation_down_AC_max, 2));
        table.getRow(3).getCell(6).setText(formatDouble(voltage_deviation_down_AC_min, 2));
        table.getRow(3).getCell(7).setText(formatDouble(voltage_deviation_down_limit, 2));

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
}