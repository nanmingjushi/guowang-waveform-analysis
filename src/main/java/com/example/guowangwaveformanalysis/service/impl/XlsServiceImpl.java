package com.example.guowangwaveformanalysis.service.impl;

import com.example.guowangwaveformanalysis.service.XlsService;
import lombok.Getter;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

import java.math.BigInteger;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

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
    public String processExcelFile(MultipartFile file) throws Exception {
        try (InputStream excelStream = file.getInputStream()) {
            // 直接从流中读取Excel数据
            ExcelSheetData data = parseExcelFromStream(excelStream);

            // 生成Word文档
            String outputPath = generateWordDocument(data);
            log.info("Word文档已生成：{}", outputPath);

            return outputPath;
        } catch (Exception e) {
            log.error("处理文件时发生异常: {}", e.getMessage(), e);
            throw e;
        }
    }

    /* 从流中解析Excel数据（使用Hutool）
     */
    private ExcelSheetData parseExcelFromStream(InputStream excelStream) {
        ExcelSheetData data = new ExcelSheetData();


        // 使用Hutool读取Excel
        ExcelReader reader = ExcelUtil.getReader(excelStream);

        try {
            // 读取电压谐波Sheet
            reader.setSheet("电压谐波");
            List<List<Object>> voltageData = reader.read();
            processMergedCells(voltageData, reader.getSheet());
            data.getVoltageHarmonicData().addAll(voltageData);

            // 读取电流谐波Sheet
            reader.setSheet("电流谐波");
            List<List<Object>> currentData = reader.read();
            processMergedCells(currentData, reader.getSheet());
            data.getCurrentHarmonicData().addAll(currentData);

            // 读取功率Sheet
            reader.setSheet("功率");
            List<List<Object>> powerData = reader.read();
            processMergedCells(powerData, reader.getSheet());
            data.getPowerData().addAll(powerData);

            /*// 打印前20行数据用于调试
            System.out.println("===== 电压谐波数据前20行 =====");
            for (int i = 0; i < Math.min(20, data.getVoltageHarmonicData().size()); i++) {
                System.out.printf("行 %2d: %s%n", i, data.getVoltageHarmonicData().get(i));
            }*/

        } finally {
            reader.close();
        }

        return data;
    }

    /*处理合并单元格，填充合并区域的值
     */
    private void processMergedCells(List<List<Object>> sheetData, Sheet sheet) {
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();

        for (CellRangeAddress region : mergedRegions) {
            int firstRow = region.getFirstRow();
            int lastRow = region.getLastRow();
            int firstCol = region.getFirstColumn();
            int lastCol = region.getLastColumn();

            // 获取合并单元格的值
            Object value = sheetData.get(firstRow).get(firstCol);

            // 填充合并区域的其他单元格
            for (int row = firstRow; row <= lastRow; row++) {
                for (int col = firstCol; col <= lastCol; col++) {
                    if (row != firstRow || col != firstCol) {
                        // 确保行和列在数据范围内
                        if (row < sheetData.size() && col < sheetData.get(row).size()) {
                            sheetData.get(row).set(col, value);
                        }
                    }
                }
            }
        }
    }


    private String generateWordDocument(ExcelSheetData data) throws IOException {
        XWPFDocument doc = new XWPFDocument();

        // 添加各表格
        addVoltageHarmonicTable(doc, data.getVoltageHarmonicData());
        addCurrentHarmonicTable(doc, data.getCurrentHarmonicData());
        addFrequencyDeviationAndVoltageUnbalanceAndLongTermFlicker(
                doc, data.getPowerData(), data.getVoltageHarmonicData());
        addVoltageDeviation(doc, data.getVoltageHarmonicData());

        // 保存Word文档
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

    // 数据存储类（保持不变）
    @Getter
    private static class ExcelSheetData {
        private final List<List<Object>> voltageHarmonicData = new ArrayList<>();
        private final List<List<Object>> currentHarmonicData = new ArrayList<>();
        private final List<List<Object>> powerData = new ArrayList<>();
    }


    // 添加电压谐波表格数据
    private void addVoltageHarmonicTable(XWPFDocument doc, List<List<Object>> data) {
        /*// 1. 打印原始数据前20行用于调试
        System.out.println("===== 原始数据前20行 =====");
        for (int i = 0; i < Math.min(20, data.size()); i++) {
            System.out.printf("行 %2d: %s%n", i, data.get(i));
        }*/

        // 监测位置
        String monitorPosition = data.get(1).get(0).toString();
        /*System.out.println("\n===== 关键数据验证 =====");
        System.out.printf("监测位置 (行1列0): %s%n", monitorPosition);*/

        // 谐波电压表格数据读取
        double AB_average_fundamental = getDoubleValue(data.get(9).get(3));
        List<Double> AB_average_hruh_list = readColumnRange(data, 3, 10, 33);
        double AB_average_THD = getDoubleValue(data.get(59).get(3));
        /*System.out.println("AB_average_fundamental: " + AB_average_fundamental);
        System.out.println("AB_average_hruh_list: " + AB_average_hruh_list);
        System.out.println("AB_average_THD: " + AB_average_THD);*/


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

        // 表名
        XWPFParagraph title = doc.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = title.createRun();
        run.setFontSize(12);
        run.setFontFamily("宋体");
        run.setText(monitorPosition.substring(5) + "谐波电压统计表");
        // 创建一个新的表格（先创建空表，再动态添加行和列）
        XWPFTable table = doc.createTable();

        // 预先创建28行
        for (int i = 0; i < 27; i++) {
            XWPFTableRow row = table.createRow();
            // 每行创建9列
            for (int j = 0; j < 8; j++) {
                row.createCell();
            }
        }

        setTableStyle(table);


        // 参数
        // 使用安全方法设置单元格内容
        safeSetCellText(table, 0, 0, "参数");
        mergeCells(table, 0, 0, 1, 1);

        // AB
        safeSetCellText(table, 0, 2, "AB");
        safeSetCellText(table, 1, 2, "平均值");
        safeSetCellText(table, 1, 3, "95%值");
        mergeCells(table, 0, 2, 0, 3);

        // BC
        safeSetCellText(table, 0, 4, "BC");
        safeSetCellText(table, 1, 4, "平均值");
        safeSetCellText(table, 1, 5, "95%值");
        mergeCells(table, 0, 4, 0, 5);

        // AC
        safeSetCellText(table, 0, 6, "AC");
        safeSetCellText(table, 1, 6, "平均值");
        safeSetCellText(table, 1, 7, "95%值");
        mergeCells(table, 0, 6, 0, 7);

        // 限值
        safeSetCellText(table, 0, 8, "限值");
        safeSetCellText(table, 2, 8, "—");
        mergeCells(table, 0, 8, 1, 8);

        // 基波电压(kV)
        safeSetCellText(table, 2, 0, "基波电压(kV)");
        mergeCells(table, 2, 0, 2, 1);

        // 2至25次谐波电压含有率(%)
        safeSetCellText(table, 3, 0, "2至25次谐波电压含有率(%)");
        mergeCells(table, 3, 0, 26, 0);

        // 2 ~ 25次
        for (int i = 3; i < 27; i++) {
            safeSetCellText(table, i, 1, String.valueOf(i - 1));
        }

        // 电压总畸变率(%)
        safeSetCellText(table, 27, 0, "电压总畸变率(%)");
        mergeCells(table, 27, 0, 27, 1);

        // 填充数据
        safeSetCellText(table, 2, 2, formatDouble(AB_average_fundamental / 1000, 2));

        /*System.out.println("AB_average_hruh_list 大小: " + AB_average_hruh_list.size());
        System.out.println("AB_average_hruh_list 内容: " + AB_average_hruh_list);*/

        for (int i = 0; i < AB_average_hruh_list.size(); i++) {
            safeSetCellText(table, i + 3, 2, formatDouble(AB_average_hruh_list.get(i), 2));
        }
        safeSetCellText(table, 27, 2, formatDouble(AB_average_THD, 2));

        safeSetCellText(table, 2, 3, formatDouble(AB_95_fundamental / 1000, 2));
        for (int i = 0; i < AB_95_hruh_list.size(); i++) {
            safeSetCellText(table, i + 3, 3, formatDouble(AB_95_hruh_list.get(i), 2));
        }
        safeSetCellText(table, 27, 3, formatDouble(AB_95_THD, 2));

        safeSetCellText(table, 2, 4, formatDouble(BC_average_fundamental / 1000, 2));
        for (int i = 0; i < BC_average_hruh_list.size(); i++) {
            safeSetCellText(table, i + 3, 4, formatDouble(BC_average_hruh_list.get(i), 2));
        }
        safeSetCellText(table, 27, 4, formatDouble(BC_average_THD, 2));

        safeSetCellText(table, 2, 5, formatDouble(BC_95_fundamental / 1000, 2));
        for (int i = 0; i < BC_95_hruh_list.size(); i++) {
            safeSetCellText(table, i + 3, 5, formatDouble(BC_95_hruh_list.get(i), 2));
        }
        safeSetCellText(table, 27, 5, formatDouble(BC_95_THD, 2));

        safeSetCellText(table, 2, 6, formatDouble(CA_average_fundamental / 1000, 2));
        for (int i = 0; i < CA_average_hruh_list.size(); i++) {
            safeSetCellText(table, i + 3, 6, formatDouble(CA_average_hruh_list.get(i), 2));
        }
        safeSetCellText(table, 27, 6, formatDouble(CA_average_THD, 2));

        safeSetCellText(table, 2, 7, formatDouble(CA_95_fundamental / 1000, 2));
        for (int i = 0; i < CA_95_hruh_list.size(); i++) {
            safeSetCellText(table, i + 3, 7, formatDouble(CA_95_hruh_list.get(i), 2));
        }
        safeSetCellText(table, 27, 7, formatDouble(CA_95_THD, 2));

        // hruh 限值
        for (int i = 0; i < limit_hruh_list.size(); i++) {
            safeSetCellText(table, i + 3, 8, formatDouble(limit_hruh_list.get(i), 2));
        }
        // thd 限值
        safeSetCellText(table, 27, 8, formatDouble(limit_THD, 2));

        setTableStyle(table);

    }


    // 添加电流谐波表格数据
    private void addCurrentHarmonicTable(XWPFDocument doc, List<List<Object>> data) {
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

        // 3. 创建表格结构
        doc.createParagraph().createRun().addBreak(BreakType.PAGE);

        XWPFParagraph title = doc.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = title.createRun();
        run.setFontSize(12);
        run.setFontFamily("宋体");
        run.setText(monitorPosition.substring(5) + "谐波电流统计表");

        // 4. 表格初始化（28行11列）
        XWPFTable table = doc.createTable(27, 11);


        // 5. 表头设置
        mergeCells(table, 0, 0, 1, 1);
        table.getRow(0).getCell(0).setText("参数");

        mergeCells(table, 0, 2, 0, 3);
        table.getRow(0).getCell(2).setText("A相");
        table.getRow(1).getCell(2).setText("平均值");
        table.getRow(1).getCell(3).setText("95%值");

        mergeCells(table, 0, 4, 0, 5);
        table.getRow(0).getCell(4).setText("B相");
        table.getRow(1).getCell(4).setText("平均值");
        table.getRow(1).getCell(5).setText("95%值");

        mergeCells(table, 0, 6, 0, 7);
        table.getRow(0).getCell(6).setText("C相");
        table.getRow(1).getCell(6).setText("平均值");
        table.getRow(1).getCell(7).setText("95%值");

        mergeCells(table, 0, 8, 1, 8);
        table.getRow(0).getCell(8).setText("限值");
        safeSetCellText(table, 2, 8, "—");

        mergeCells(table, 0, 9, 1, 9);
        table.getRow(0).getCell(9).setText("限值(0.66MVA)");
        safeSetCellText(table, 2, 9, "—");

        mergeCells(table, 0, 10, 1, 10);
        table.getRow(0).getCell(10).setText("限值(150MVA)");
        safeSetCellText(table, 2, 10, "—");

        // 6. 填充数据
        // 基波电流
        mergeCells(table, 2, 0, 2, 1);
        table.getRow(2).getCell(0).setText("基波电流(A)");
        table.getRow(2).getCell(2).setText(formatDouble(A_average_fundamental, 2));
        table.getRow(2).getCell(3).setText(formatDouble(A_95_fundamental, 2));
        table.getRow(2).getCell(4).setText(formatDouble(B_average_fundamental, 2));
        table.getRow(2).getCell(5).setText(formatDouble(B_95_fundamental, 2));
        table.getRow(2).getCell(6).setText(formatDouble(C_average_fundamental, 2));
        table.getRow(2).getCell(7).setText(formatDouble(C_95_fundamental, 2));

        // 谐波电流
        mergeCells(table, 3, 0, 26, 0);
        table.getRow(3).getCell(0).setText("2至25次谐波电流含有率(%)");

        for (int i = 3; i < 27; i++) {
            table.getRow(i).getCell(1).setText(String.valueOf(i - 1));
        }

        // 填充各相数据
        for (int i = 0; i < 24; i++) {
            table.getRow(i + 3).getCell(2).setText(formatDouble(A_average_hruh_list.get(i), 2));
            table.getRow(i + 3).getCell(3).setText(formatDouble(A_95_hruh_list.get(i), 2));
            table.getRow(i + 3).getCell(4).setText(formatDouble(B_average_hruh_list.get(i), 2));
            table.getRow(i + 3).getCell(5).setText(formatDouble(B_95_hruh_list.get(i), 2));
            table.getRow(i + 3).getCell(6).setText(formatDouble(C_average_hruh_list.get(i), 2));
            table.getRow(i + 3).getCell(7).setText(formatDouble(C_95_hruh_list.get(i), 2));
            table.getRow(i + 3).getCell(8).setText(formatDouble(limit_hruh_list.get(i), 2));
        }


        setTableStyle(table);

    }
    // 添加频率偏差、三相电压不平衡度及长时间闪变表格数据
    private void addFrequencyDeviationAndVoltageUnbalanceAndLongTermFlicker(
            XWPFDocument doc, List<List<Object>> powerData, List<List<Object>> voltageHarmonicData) {

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

        // 3. 创建表格结构
        doc.createParagraph().createRun().addBreak(BreakType.PAGE);

        XWPFParagraph title = doc.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = title.createRun();
        run.setFontSize(12);
        run.setFontFamily("宋体");
        run.setText(monitorPosition.substring(5) + "频率偏差、三相电压不平衡度及长时间闪变统计表");

        // 4. 表格初始化（6行7列）
        XWPFTable table = doc.createTable(6, 7);


        // 5. 表头设置
        mergeCells(table, 0, 0, 0, 1);
        table.getRow(0).getCell(0).setText("参数");
        table.getRow(0).getCell(2).setText("最大值");
        table.getRow(0).getCell(3).setText("平均值");
        table.getRow(0).getCell(4).setText("最小值");
        table.getRow(0).getCell(5).setText("95%值");
        table.getRow(0).getCell(6).setText("限值");

        // 6. 填充频率数据
        mergeCells(table, 1, 0, 1, 1);
        table.getRow(1).getCell(0).setText("频率(Hz)");
        table.getRow(1).getCell(2).setText(formatDouble(frequency_max, 2));
        table.getRow(1).getCell(3).setText(formatDouble(frequency_average, 2));
        table.getRow(1).getCell(4).setText(formatDouble(frequency_min, 2));
        table.getRow(1).getCell(5).setText(formatDouble(frequency_95, 2));
        table.getRow(1).getCell(6).setText(frequency_limit);

        // 7. 填充电压不平衡度数据
        mergeCells(table, 2, 0, 2, 1);
        table.getRow(2).getCell(0).setText("三相电压不平衡度(%)");
        table.getRow(2).getCell(2).setText(formatDouble(voltage_unbalance_max, 2));
        table.getRow(2).getCell(3).setText(formatDouble(voltage_unbalance_average, 2));
        table.getRow(2).getCell(4).setText(formatDouble(voltage_unbalance_min, 2));
        table.getRow(2).getCell(5).setText(formatDouble(voltage_unbalance_95, 2));
        table.getRow(2).getCell(6).setText(formatDouble(voltage_unbalance_limit, 2));

        // 8. 填充长时间闪变数据
        mergeCells(table, 3, 0, 5, 0);
        table.getRow(3).getCell(0).setText("长时间闪变(Plt)");
        table.getRow(3).getCell(1).setText("AB");
        table.getRow(4).getCell(1).setText("BC");
        table.getRow(5).getCell(1).setText("AC");

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

        setTableStyle(table);

    }
    // 添加电压偏差表格数据
    private void addVoltageDeviation(XWPFDocument doc, List<List<Object>> voltageHarmonicData) {
        // 1. 基础数据准备
        String monitorPosition = voltageHarmonicData.get(1).get(0).toString();

        // 2. 数据提取（上偏差）
        double voltage_deviation_up_AB_max = getDoubleValue(voltageHarmonicData.get(63).get(2));
        double voltage_deviation_up_AB_min = getDoubleValue(voltageHarmonicData.get(63).get(4));
        double voltage_deviation_up_BC_max = getDoubleValue(voltageHarmonicData.get(63).get(7));
        double voltage_deviation_up_BC_min = getDoubleValue(voltageHarmonicData.get(63).get(9));
        double voltage_deviation_up_AC_max = getDoubleValue(voltageHarmonicData.get(63).get(12));
        double voltage_deviation_up_AC_min = getDoubleValue(voltageHarmonicData.get(63).get(14));
        double voltage_deviation_up_limit = getDoubleValue(voltageHarmonicData.get(63).get(17));

        // 3. 数据提取（下偏差）
        double voltage_deviation_down_AB_max = getDoubleValue(voltageHarmonicData.get(64).get(2));
        double voltage_deviation_down_AB_min = getDoubleValue(voltageHarmonicData.get(64).get(4));
        double voltage_deviation_down_BC_max = getDoubleValue(voltageHarmonicData.get(64).get(7));
        double voltage_deviation_down_BC_min = getDoubleValue(voltageHarmonicData.get(64).get(9));
        double voltage_deviation_down_AC_max = getDoubleValue(voltageHarmonicData.get(64).get(12));
        double voltage_deviation_down_AC_min = getDoubleValue(voltageHarmonicData.get(64).get(14));
        double voltage_deviation_down_limit = Double.parseDouble("-"+voltageHarmonicData.get(64).get(17).toString());

        // 4. 创建表格结构
        doc.createParagraph().createRun().addBreak(BreakType.PAGE);

        XWPFParagraph title = doc.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = title.createRun();
        run.setFontSize(12);
        run.setFontFamily("宋体");
        run.setText(monitorPosition.substring(5) + "电压偏差统计表");

        // 5. 表格初始化（4行8列）
        XWPFTable table = doc.createTable(4, 8);


        // 6. 表头设置
        // 参数列
        mergeCells(table, 0, 0, 1, 0);
        table.getRow(0).getCell(0).setText("参数");

        // AB相标题
        mergeCells(table, 0, 1, 0, 2);
        table.getRow(0).getCell(1).setText("AB");
        table.getRow(1).getCell(1).setText("最大值");
        table.getRow(1).getCell(2).setText("最小值");

        // BC相标题
        mergeCells(table, 0, 3, 0, 4);
        table.getRow(0).getCell(3).setText("BC");
        table.getRow(1).getCell(3).setText("最大值");
        table.getRow(1).getCell(4).setText("最小值");

        // AC相标题
        mergeCells(table, 0, 5, 0, 6);
        table.getRow(0).getCell(5).setText("AC");
        table.getRow(1).getCell(5).setText("最大值");
        table.getRow(1).getCell(6).setText("最小值");

        // 限值标题
        mergeCells(table, 0, 7, 1, 7);
        table.getRow(0).getCell(7).setText("限值");

        // 7. 填充数据
        // 上偏差数据
        table.getRow(2).getCell(0).setText("上偏差(%)");
        table.getRow(2).getCell(1).setText(formatDouble(voltage_deviation_up_AB_max, 2));
        table.getRow(2).getCell(2).setText(formatDouble(voltage_deviation_up_AB_min, 2));
        table.getRow(2).getCell(3).setText(formatDouble(voltage_deviation_up_BC_max, 2));
        table.getRow(2).getCell(4).setText(formatDouble(voltage_deviation_up_BC_min, 2));
        table.getRow(2).getCell(5).setText(formatDouble(voltage_deviation_up_AC_max, 2));
        table.getRow(2).getCell(6).setText(formatDouble(voltage_deviation_up_AC_min, 2));
        table.getRow(2).getCell(7).setText(formatDouble(voltage_deviation_up_limit, 2));

        // 下偏差数据
        table.getRow(3).getCell(0).setText("下偏差(%)");
        table.getRow(3).getCell(1).setText(formatDouble(voltage_deviation_down_AB_max, 2));
        table.getRow(3).getCell(2).setText(formatDouble(voltage_deviation_down_AB_min, 2));
        table.getRow(3).getCell(3).setText(formatDouble(voltage_deviation_down_BC_max, 2));
        table.getRow(3).getCell(4).setText(formatDouble(voltage_deviation_down_BC_min, 2));
        table.getRow(3).getCell(5).setText(formatDouble(voltage_deviation_down_AC_max, 2));
        table.getRow(3).getCell(6).setText(formatDouble(voltage_deviation_down_AC_min, 2));
        table.getRow(3).getCell(7).setText(formatDouble(voltage_deviation_down_limit, 2));

        setTableStyle(table);

    }

    // 合并单元格方法
    private void mergeCells(XWPFTable table, int startRow, int startCol, int endRow, int endCol) {
        // 确保表格有足够的行和列
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            // 确保行存在
            XWPFTableRow row = table.getRow(rowIndex);
            if (row == null) {
                row = table.createRow();
            }

            // 确保列存在
            for (int colIndex = startCol; colIndex <= endCol; colIndex++) {
                if (row.getCell(colIndex) == null) {
                    row.createCell();
                }
            }
        }

        // 执行合并操作
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            XWPFTableRow row = table.getRow(rowIndex);
            for (int colIndex = startCol; colIndex <= endCol; colIndex++) {
                XWPFTableCell cell = row.getCell(colIndex);
                CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();

                // 水平合并
                if (colIndex > startCol) {
                    if (!tcPr.isSetHMerge()) {
                        tcPr.addNewHMerge().setVal(STMerge.CONTINUE);
                    }
                } else {
                    if (!tcPr.isSetHMerge()) {
                        tcPr.addNewHMerge().setVal(STMerge.RESTART);
                    }
                }

                // 垂直合并
                if (rowIndex > startRow) {
                    if (!tcPr.isSetVMerge()) {
                        tcPr.addNewVMerge().setVal(STMerge.CONTINUE);
                    }
                } else {
                    if (!tcPr.isSetVMerge()) {
                        tcPr.addNewVMerge().setVal(STMerge.RESTART);
                    }
                }
            }
        }
    }

    //设置表格基础样式
    private void setTableStyle(XWPFTable table) {
        // 1. 表格整体设置
        // 设置表格宽度为100%，使其充满容器
        table.setWidth("100%");
        // 设置表格居中对齐
        table.setTableAlignment(TableRowAlign.CENTER);

        // 强制固定表格布局（避免自动调整）
//        table.getCTTbl().addNewTblPr().addNewTblLayout().setType(STTblLayoutType.FIXED);

        // 2. 单元格统一设置
        for (XWPFTableRow row : table.getRows()) {
            if (row == null) continue;

            // 设置行高（强制固定）
            row.setHeight(300);
            row.getCtRow().addNewTrPr().addNewTrHeight().setVal(BigInteger.valueOf(300));

            for (XWPFTableCell cell : row.getTableCells()) {
                if (cell == null) continue;

                // 3. 单元格宽度设置
                CTTcPr tcPr = cell.getCTTc().isSetTcPr()
                        ? cell.getCTTc().getTcPr()
                        : cell.getCTTc().addNewTcPr();

                // 设置默认宽度
                if (!tcPr.isSetTcW()) {
                    tcPr.addNewTcW().setW(BigInteger.valueOf(600));
                }

                // 4. 垂直居中设置
                if (tcPr.isSetVAlign()) {
                    tcPr.getVAlign().setVal(STVerticalJc.CENTER);
                } else {
                    tcPr.addNewVAlign().setVal(STVerticalJc.CENTER);
                }

                // 5. 确保单元格有段落
                if (cell.getParagraphs().isEmpty()) {
                    cell.addParagraph();
                }

                // 6. 设置段落水平居中
                XWPFParagraph para = cell.getParagraphs().get(0);
                para.setAlignment(ParagraphAlignment.CENTER);

                // 7. 设置默认字体（宋体12号）
                for (XWPFRun run : para.getRuns()) {
                    run.setFontSize(9);
                    CTRPr rPr = run.getCTR().getRPr();
                    if (rPr == null) {
                        rPr = run.getCTR().addNewRPr();
                    }

                    CTFonts fonts = rPr.addNewRFonts();

                    // 设置中文字体为宋体
                    fonts.setEastAsia("宋体");
                    // 设置英文字体和数字字体为 Times New Roman
                    fonts.setAscii("Times New Roman");
                    fonts.setHAnsi("Times New Roman");
                }
            }
        }
    }

    private void safeSetCellText(XWPFTable table, int rowNum, int colNum, String text) {
        // 确保行存在
        while (table.getNumberOfRows() <= rowNum) {
            table.createRow();
        }
        XWPFTableRow row = table.getRow(rowNum);

        // 确保列存在
        while (row.getTableCells().size() <= colNum) {
            row.createCell();
        }

        // 设置文本
        XWPFTableCell cell = row.getCell(colNum);
        if (cell.getParagraphs().isEmpty()) {
            cell.addParagraph();
        }
        XWPFParagraph para = cell.getParagraphs().get(0);
        if (para.getRuns().isEmpty()) {
            para.createRun();
        }
        para.getRuns().get(0).setText(text);
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
