package com.igalia.ooxmlexporter;

import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFPieChartData;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.data.general.DefaultPieDataset;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ChartConverter {
    private String inputFile;

    public ChartConverter(String inputFile) {
        this.inputFile = inputFile;
    }

    public void convert() throws IOException {
        Workbook wb = WorkbookFactory.create(new FileInputStream(inputFile));
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
            for (XSSFChart chart : drawing.getCharts()) {
                for (XDDFChartData series: chart.getChartSeries()) {
                    if (series instanceof XDDFPieChartData) {
                        convertPieChart((XDDFPieChartData) series);
                    }
                }
            }
        }
    }

    private void convertPieChart(XDDFPieChartData chart) throws IOException {
        if (chart.getSeriesCount() > 0) {
            XDDFChartData.Series series = chart.getSeries(0);
            XDDFDataSource<?> categories = series.getCategoryData();
            XDDFNumericalDataSource<? extends Number> values = series.getValuesData();

            DefaultPieDataset outputChartData = new DefaultPieDataset();
            for (int i = 0; i < categories.getPointCount(); i++) {
                String category = categories.getPointAt(i).toString();
                Number value = values.getPointAt(i);
                outputChartData.setValue(category, value);
            }
            JFreeChart outputPieChart = ChartFactory.createPieChart("Excel Pie Chart Java Example", outputChartData, true, true, false);
            ChartUtils.saveChartAsJPEG(new File(inputFile + ".jpg"), outputPieChart, 640, 480);
        }
    }
}
