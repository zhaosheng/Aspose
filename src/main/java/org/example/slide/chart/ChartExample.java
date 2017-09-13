package org.example.slide.chart;

import com.aspose.slides.ChartTypeCharacterizer;
import com.aspose.slides.IBaseSlide;
import com.aspose.slides.IChart;
import com.aspose.slides.IChartCategoryCollection;
import com.aspose.slides.IChartDataCell;
import com.aspose.slides.IChartDataPointCollection;
import com.aspose.slides.IChartDataWorkbook;
import com.aspose.slides.IChartSeries;
import com.aspose.slides.IChartSeriesCollection;
import com.aspose.slides.Presentation;
import com.aspose.slides.SaveFormat;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


public class ChartExample {
  private static class ChartData {
    List<Map<String, String>> chartData;

    ChartData() {
      this.chartData = new ArrayList<Map<String, String>>();
    }

    List<Map<String, String>> getChartData() {
      return chartData;
    }

    ChartData addRow(String quarter, String revenue, String yoy) {
      HashMap<String, String> row = new HashMap<String, String>();
      row.put("Quarter", quarter);
      row.put("Revenue", revenue);
      row.put("YoY", yoy);
      this.chartData.add(row);
      return this;
    }
  }

  private ChartData loadData() {
    return new ChartData()
        .addRow("2016-Q1", "11", "0.3")
        .addRow("2016-Q2", "18", "0.21")
        .addRow("2016-Q3", "21", "0.19")
        .addRow("2016-Q4", "21", "0.53")
        .addRow("2017-Q1", "23", "0.13")
        .addRow("2017-Q2", "31", "0.23")
        .addRow("2017-Q3", "16", "0.48")
        .addRow("2017-Q4", "19", "0.74")
        ;
  }

  void processChart(IChart chart) {
    ChartData chartData = loadData();
    IChartCategoryCollection categories = chart.getChartData().getCategories();
    IChartSeriesCollection series = chart.getChartData().getSeries();
    IChartDataWorkbook wb = chart.getChartData().getChartDataWorkbook();

    String categoryName = wb.getCell(0, 0, 0).getValue().toString();
    IChartSeries currentSeries;
    IChartDataPointCollection dataPoints;
    int wbIndex = 0;
    List<Map<String, String>> data = chartData.getChartData();
    Map<String, String> element;
    for (int i = 0; i < data.size(); i++) {
      element = data.get(i);
      int row = i + 1;
      categories.add(wb.getCell(wbIndex, row, 0, element.get(categoryName)));
      for (int j = 0; j < series.size(); j++) {
        currentSeries = series.get_Item(j);
        dataPoints = currentSeries.getDataPoints();

        String val = element.get(currentSeries.getName().toString());
        if (val == null) {
          String msg =
              String.format("Chart requests series named %s but not found from given data!", currentSeries.getName());
          throw new RuntimeException(msg);
        }
        Object parseVal;
        try {
          parseVal = Double.parseDouble(val);
        } catch (NumberFormatException e) {
          // cannot parse into double number, fallback to string value.
          parseVal = val;
        }
        IChartDataCell cell = wb.getCell(wbIndex, row, j + 1, parseVal);
        System.out.println("ChartType: " + chart.getType());
        if (ChartTypeCharacterizer.isChartTypeLine(chart.getType())) {        //  || chart.getType() == 73
          dataPoints.addDataPointForLineSeries(cell);
        } else if (ChartTypeCharacterizer.isChartTypeBar(chart.getType())) {
          dataPoints.addDataPointForBarSeries(cell);
        } else if (ChartTypeCharacterizer.isChartTypeArea(chart.getType())) {
          dataPoints.addDataPointForAreaSeries(cell);
        } else if (ChartTypeCharacterizer.isChartTypeBubble(chart.getType())) {
          throw new ChartTypeNotImplementedException("bubble");
        } else if (ChartTypeCharacterizer.isChartTypeColumn(chart.getType())) {
          dataPoints.addDataPointForBarSeries(cell);
        } else if (ChartTypeCharacterizer.isChartTypeDoughnut(chart.getType())) {
          dataPoints.addDataPointForDoughnutSeries(cell);
        } else if (ChartTypeCharacterizer.isChartTypePie(chart.getType())) {
          dataPoints.addDataPointForPieSeries(cell);
        } else if (ChartTypeCharacterizer.isChartTypeRadar(chart.getType())) {
          dataPoints.addDataPointForRadarSeries(cell);
        } else if (ChartTypeCharacterizer.isChartTypeScatter(chart.getType())) {
          throw new ChartTypeNotImplementedException("scatter");
        } else if (ChartTypeCharacterizer.isChartTypeStock(chart.getType())) {
          throw new ChartTypeNotImplementedException("stock");
        } else if (ChartTypeCharacterizer.isChartTypeSurface(chart.getType())) {
          dataPoints.addDataPointForSurfaceSeries(cell);
        } else {
          throw new ChartTypeNotImplementedException("unknown chart type: " + chart.getType());
        }
      }
    }
  }

  void loadPresentation(String name) {
    InputStream resourceAsStream = this.getClass().getClassLoader().getResourceAsStream(name);
    Presentation presentation = new Presentation(resourceAsStream);
    IBaseSlide slideById = presentation.getSlides().get_Item(0);
    IChart item = ((IChart) slideById.getShapes().get_Item(0));
    processChart(item);
    presentation.save("output.pptx", SaveFormat.Pptx);
  }

  public static void main(String[] args) {
    String presentationName = args.length > 0 ? args[0] : "mixChart.pptx";
    ChartExample example = new ChartExample();
    example.loadPresentation(presentationName);
  }

  private class ChartTypeNotImplementedException extends RuntimeException {
    ChartTypeNotImplementedException(String msg) {
    }
  }
}
