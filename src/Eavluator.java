import java.awt.Font;
import java.io.*;

import jxl.*;
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.format.VerticalAlignment;
import jxl.write.*;
import jxl.write.Label;
import jxl.write.Number;
import javax.swing.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.awt.*;
import java.awt.event.*;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.*;
import org.jfree.chart.labels.StandardXYItemLabelGenerator;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.renderer.category.BarRenderer;
import org.jfree.chart.title.LegendTitle;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.renderer.category.LineAndShapeRenderer;
import org.jfree.ui.RectangleEdge;

public class Eavluator {
    private static SimpleDateFormat myFormat = new SimpleDateFormat("yyyyMMdd");

    private static Queue<String> buyerTracked;
    private static Performance[] buyerPerformance;
    private static Performance totalPerformance = new Performance("Total");
    public static String todayDate;
    private static String threeDaysBef;
    private static String startPlotDate;
    private static String endPlotDate;
    private static String actualTodayDate;

    private static WritableCellFormat titleFormat;
    private static WritableCellFormat goodFormat;
    private static WritableCellFormat expiredFormat;
    private static WritableCellFormat noneFormat;
    private static WritableCellFormat normalFormat;

    private static Queue<String> plotTrackingBuyers;
    private static Queue<String> plotTrackingDate;
    private static PerformancePlotData[] dataPerformance;
    private static Queue<String> dateQueue;
    private static Queue<String> datePlotQueue;
    private static String[] dateArr;
    private static String[] buyersArr;
    private static ArrayList<Integer> removeIdx = new ArrayList<>();

    private static int count = 0;

    public static void main(String[] args){
        actualTodayDate = getTodayDate();
        initializeFormat();
        createGUI();
    }

    private static void generateTable() throws Exception{
        copyFile("../SourceFiles/OpenPO" + todayDate + ".xls", "../ProcessedFiles/OpenPO" + todayDate + ".xls");
        String filePath = "../ProcessedFiles/OpenPO" + todayDate + ".xls";
        File openPoData = new File(filePath);
        buyerTracked = new LinkedList<>();
        buyerPerformance = new Performance[2];
        initializeFormat();
        setBorders(false);
        getBuyerPerformance(openPoData);
        outputData();
    }

    private static void copyFile(String oldPath, String newPath) throws IOException {
        File oldFile = new File(oldPath);
        File file = new File(newPath);
        FileInputStream in = new FileInputStream(oldFile);
        FileOutputStream out = new FileOutputStream(file);
        byte[] buffer=new byte[2097152];
        while((in.read(buffer)) != -1){
            out.write(buffer);
        }
    }

    private static String getTodayDate(){
        Calendar myCalendar = Calendar.getInstance();
        return myFormat.format(myCalendar.getTime());
    }

    private static String dateAddition(String date, int num){
        Date dateToCal;
        try {
            dateToCal = myFormat.parse(date);
        }
        catch (Exception e){
            System.out.println("Format error");
            return "";
        }
        long finalDate=(dateToCal.getTime()/1000) + 60*60*24*num;
        dateToCal.setTime(finalDate*1000);
        return myFormat.format(dateToCal);
    }

    private static int findIdx(String obj, boolean isDate){
        String[] searchArr;
        if(isDate){
            searchArr = dateArr;
        }
        else{
            searchArr = buyersArr;
        }
        for(int i = 0; i < searchArr.length;i++){
            if(obj.equals(searchArr[i])){
                return i;
            }
        }
        return -1;
    }

    public static void generatePlot() throws Exception{
        dateQueue = new LinkedList<>();
        datePlotQueue = new LinkedList<>();
        dataPerformance = new PerformancePlotData[2];
        initializeFormat();
        setBorders(true);
        getAllDate();
        retainData();
        DefaultCategoryDataset dataSet = new DefaultCategoryDataset();
        DefaultCategoryDataset expireDataSet = new DefaultCategoryDataset();
        DefaultCategoryDataset missedDataSet = new DefaultCategoryDataset();
        String outputFilePath = "../PerformanceOutput/Plots/Comparison" + startPlotDate + "_" + endPlotDate + ".xls";
        WritableWorkbook outputFile = Workbook.createWorkbook(new File(outputFilePath));
        WritableSheet mySheet = outputFile.createSheet("Comparison Page 1", 0);
        WritableSheet mySheet_2 = outputFile.createSheet("Comparison Page 2", 1);
        jxl.write.Label buyerLabel = new jxl.write.Label(0, 1, "Buyer", titleFormat);
        mySheet.addCell(buyerLabel);
        mySheet.setColumnView(0, 15);
        mySheet_2.setColumnView(0, 15);
        mySheet_2.setColumnView(1, 21);
        Calendar calendar = Calendar.getInstance();
        calendar.setFirstDayOfWeek(Calendar.SUNDAY);
        int count = 0;
        String[] tempDateArr = new String[plotTrackingDate.size()];
        while(!plotTrackingDate.isEmpty()){
            String currDate = plotTrackingDate.poll();
            Date currWeek = myFormat.parse(currDate);
            calendar.setTime(currWeek);
            jxl.write.Label weekLabel = new jxl.write.Label((count*6+1), 0, "Week " + String.valueOf(calendar.get(Calendar.WEEK_OF_YEAR)));
            mySheet.addCell(weekLabel);
            datePlotQueue.add(String.valueOf(calendar.get(Calendar.WEEK_OF_YEAR)));
            jxl.write.Label day = new jxl.write.Label(count*6+2, 0, "Day " + currDate);
            mySheet.addCell(day);
            jxl.write.Label sheet2Date = new jxl.write.Label(count+1, 1, "Day " + currDate);
            mySheet_2.addCell(sheet2Date);
            tempDateArr[count] = currDate;
            count++;
        }
        dateArr = datePlotQueue.toArray(new String[0]);
        buyersArr = plotTrackingBuyers.toArray(new String[0]);
        int cnt = 0;
        jxl.write.Label expireTitle = new jxl.write.Label(1, 0, "Promise Date Expired", titleFormat);
        mySheet_2.addCell(expireTitle);
        for(int i = 0; i < 2; i++) {
            cnt++;
            jxl.write.Label buyerTitle = new jxl.write.Label(0, cnt, "Buyer", titleFormat);
            mySheet_2.addCell(buyerTitle);
            for(int j = 0; j < buyersArr.length; j++) {
                cnt++;
                jxl.write.Label buyerName = new jxl.write.Label(0, cnt,buyersArr[j], normalFormat);
                mySheet_2.addCell(buyerName);
            }
            cnt++;
            jxl.write.Label totalName = new jxl.write.Label(0, cnt, "total", titleFormat);
            mySheet_2.addCell(totalName);
            cnt += 2;
            if(i == 0) {
                jxl.write.Label missedTitle = new jxl.write.Label(1, cnt, "Promise Date Missed", titleFormat);
                mySheet_2.addCell(missedTitle);
                for(int k = 0; k < tempDateArr.length; k++){
                    jxl.write.Label dateLabel = new jxl.write.Label(k+1, cnt+1, "Day " + tempDateArr[k]);
                    mySheet_2.addCell(dateLabel);
                    mySheet_2.setColumnView(k+1, 21);
                }
            }
        }
        int expireRowNum;
        int colNum;
        int missedRowNum;
        int[] expiredNumTotal = new int[dateArr.length];
        int[] missedNumTotal = new int[dateArr.length];
        for(int i = 0; i < plotTrackingBuyers.size(); i++){
            PerformancePlotData currData = dataPerformance[i];
            jxl.write.Label currBuyer = new jxl.write.Label(0, i+2, currData.getName(), normalFormat);
            mySheet.addCell(currBuyer);
            SheetSettings sheetSettings =  mySheet.getSettings();
            sheetSettings.setHorizontalFreeze(1);
            while(!currData.isEmpty()){
                Performance currPlotData = currData.poll();
                Date currWeek = myFormat.parse(currPlotData.getDate());
                calendar.setTime(currWeek);
                int currCol = findIdx(String.valueOf(calendar.get(Calendar.WEEK_OF_YEAR)), true)*6+1;
                WritableCellFormat currFormat;
                int goodNum = currPlotData.getGoodPromiseDate();
                int expiredNum = currPlotData.getExpiredPromiseDate();
                int missedNum = currPlotData.getNonePromiseDate();
                int totalNum = goodNum + expiredNum + missedNum;
                int goodPercent = (int) (currPlotData.getGoodPercent()*100);
                jxl.write.Label goodLabel = new jxl.write.Label(currCol, 1, "Promise Date OK", titleFormat);
                mySheet.addCell(goodLabel);
                mySheet.setColumnView(currCol, 16);
                currFormat = normalFormat;
                if(goodNum != 0) {
                    Number goodPromise = new Number(currCol, i+2, goodNum, currFormat);
                    mySheet.addCell(goodPromise);
                }
                else{
                    jxl.write.Label emptyLabel = new jxl.write.Label(currCol, i+2, "", currFormat);
                    mySheet.addCell(emptyLabel);
                }
                currCol++;
                jxl.write.Label expiredLabel = new jxl.write.Label(currCol, 1, "Promise Date Expired", titleFormat);
                mySheet.addCell(expiredLabel);
                mySheet.setColumnView(currCol, 21);
                currFormat = expiredFormat;
                expireRowNum = findIdx(currData.getName(), false) + 2;
                colNum = findIdx(String.valueOf(calendar.get(Calendar.WEEK_OF_YEAR)), true) + 1;
                if(expiredNum != 0) {
                    Number expiredPromise = new Number(currCol, i + 2, expiredNum, currFormat);
                    mySheet.addCell(expiredPromise);
                    Number expiredNumLabel = new Number(colNum, expireRowNum, expiredNum, currFormat);
                    mySheet_2.addCell(expiredNumLabel);
                }
                else{
                    jxl.write.Label emptyLabel = new jxl.write.Label(currCol, i+2, "", currFormat);
                    mySheet.addCell(emptyLabel);
                    jxl.write.Label sheet2EmptyLabel = new jxl.write.Label(colNum, expireRowNum, "", currFormat);
                    mySheet_2.addCell(sheet2EmptyLabel);
                }
                expiredNumTotal[findIdx(String.valueOf(calendar.get(Calendar.WEEK_OF_YEAR)), true)] += expiredNum;
                currCol++;
                jxl.write.Label missedLabel = new jxl.write.Label(currCol, 1, "Promise Date Missed", titleFormat);
                mySheet.addCell(missedLabel);
                mySheet.setColumnView(currCol, 20);
                currFormat = noneFormat;
                missedRowNum = findIdx(currData.getName(), false) + buyersArr.length + 6;
                if(missedNum != 0) {
                    Number missedPromise = new Number(currCol, i + 2, missedNum, currFormat);
                    mySheet.addCell(missedPromise);
                    Number missedNumLabel = new Number(colNum, missedRowNum, missedNum, currFormat);
                    mySheet_2.addCell(missedNumLabel);
                }
                else{
                    jxl.write.Label emptyLabel = new jxl.write.Label(currCol, i+2, "", currFormat);
                    mySheet.addCell(emptyLabel);
                    jxl.write.Label sheet2EmptyLabel = new jxl.write.Label(colNum, missedRowNum, "", currFormat);
                    mySheet_2.addCell(sheet2EmptyLabel);
                }
                missedNumTotal[findIdx(String.valueOf(calendar.get(Calendar.WEEK_OF_YEAR)), true)] += missedNum;
                currCol++;
                jxl.write.Label totalLabel = new jxl.write.Label(currCol, 1, "Total", titleFormat);
                mySheet.addCell(totalLabel);
                mySheet.setColumnView(currCol, 10);
                Number totalPromise = new Number(currCol, i+2, totalNum, normalFormat);
                mySheet.addCell(totalPromise);
                currCol++;
                jxl.write.Label percentLabel = new jxl.write.Label(currCol, 1, "Performance Percent", titleFormat);
                mySheet.addCell(percentLabel);
                mySheet.setColumnView(currCol, 20);
                if(goodPercent > 80){
                    currFormat = normalFormat;
                }
                else if(goodPercent > 60){
                    currFormat = noneFormat;
                }
                else{
                    currFormat = expiredFormat;
                }
                jxl.write.Label goodPromisePercent = new jxl.write.Label(currCol, i+2, ""+goodPercent+"%", currFormat);
                mySheet.addCell(goodPromisePercent);
                dataSet.setValue(currPlotData.getGoodPercent(), currData.getName(), "W " + String.valueOf(calendar.get(Calendar.WEEK_OF_YEAR)));
                String formattedName = currData.getName().substring(currData.getName().indexOf(' ')+1, currData.getName().indexOf(' ')+2).toUpperCase()+currData.getName().substring(currData.getName().indexOf(' ')+2).toLowerCase();
                expireDataSet.setValue(expiredNum, currPlotData.getDate(), formattedName);
                missedDataSet.setValue(missedNum, currPlotData.getDate(), formattedName);
            }
        }
        for(int i = 0; i < dateArr.length; i++){
            Number expireTotalLabel = new Number(i+1, buyersArr.length+2, expiredNumTotal[i], expiredFormat);
            Number missedTotalLabel = new Number(i+1, buyersArr.length*2+6, missedNumTotal[i], noneFormat);
            expireDataSet.setValue(expiredNumTotal[i], tempDateArr[i], "Total");
            missedDataSet.setValue(missedNumTotal[i], tempDateArr[i], "Total");
            mySheet_2.addCell(expireTotalLabel);
            mySheet_2.addCell(missedTotalLabel);
        }
        JFreeChart expiredChart = ChartFactory.createBarChart("Promise Date Expired", "", "", expireDataSet,
                PlotOrientation.VERTICAL, true, false, false);
        JFreeChart missedChart = ChartFactory.createBarChart("Promise Date Missed", "", "", missedDataSet,
                PlotOrientation.VERTICAL, true, false, false);
        JFreeChart percentChart = ChartFactory.createLineChart("Delivery Performance", "",
                "", dataSet, PlotOrientation.VERTICAL, true, false, false);
        setPlotFormat(percentChart, 0.05, true);
        setPlotFormat(expiredChart, 10, false);
        setPlotFormat(missedChart, 10, false);
        OutputStream os = new FileOutputStream("../PerformanceOutput/Plots/PercentChange" + startPlotDate + "_" + endPlotDate + ".jpg");
        ChartUtilities.writeChartAsJPEG(os, percentChart, 1250, 750);
        os.close();
        OutputStream os1 = new FileOutputStream("../PerformanceOutput/Plots/ExpiredChange" + startPlotDate + "_" + endPlotDate + ".jpg");
        ChartUtilities.writeChartAsJPEG(os1, expiredChart, 1250, 750);
        os1.close();
        OutputStream os2 = new FileOutputStream("../PerformanceOutput/Plots/MissedChange" + startPlotDate + "_" + endPlotDate + ".jpg");
        ChartUtilities.writeChartAsJPEG(os2, missedChart, 1250, 750);
        os2.close();
        outputFile.write();
        outputFile.close();
    }

    private static void setPlotFormat(JFreeChart myChart, double yAxisInt, boolean isPercent){
        myChart.removeLegend();
        CategoryPlot plot = (CategoryPlot) myChart.getPlot();
        plot.setBackgroundAlpha(0.5f);
        plot.setForegroundAlpha(0.5f);
        CategoryAxis domainAxis = plot.getDomainAxis();
        ValueAxis rAxis = plot.getRangeAxis();
        LegendTitle legendTitle = new LegendTitle(plot);
        legendTitle.setPosition(RectangleEdge.TOP);
        legendTitle.setItemFont(new Font("Arial", Font.BOLD, 16));
        myChart.addLegend(legendTitle);
        if(isPercent) {
            LineAndShapeRenderer renderer = (LineAndShapeRenderer) plot.getRenderer();
            renderer.setBaseShapesVisible(true);
            renderer.setBaseLinesVisible(true);
            renderer.setUseSeriesOffset(true);
        }
        else{
            plot.setBackgroundAlpha(0.0f);
            plot.setForegroundAlpha(1.0f);
            BarRenderer render = (BarRenderer) plot.getRenderer();
            render.setItemMargin(0.0);
            render.setSeriesPaint(0, Color.decode("#8DB6CD"));
            render.setSeriesPaint(1, Color.decode("#8C8C8C"));
            render.setSeriesPaint(2, Color.decode("#8B636C"));
            render.setSeriesPaint(3, Color.decode("#9BCD9B"));
            render.setSeriesPaint(4, Color.decode("#CD950C"));
            //domainAxis.setCategoryLabelPositions(CategoryLabelPositions.UP_45);
            plot.setRangeGridlinePaint(Color.black);
            plot.setRangeGridlinesVisible(true);
        }
        domainAxis.setTickLabelFont(new Font("Arial", Font.BOLD, 15));
        rAxis.setTickLabelFont(new Font("Arial", Font.PLAIN, 18));
//        renderer.setBaseItemLabelGenerator(new StandardXYItemLabelGenerator("{2}",NumberFormat.getPercentInstance(),
//                                NumberFormat.getPercentInstance()));
        NumberAxis numAxis = (NumberAxis) plot.getRangeAxis();
        numAxis.setTickUnit(new NumberTickUnit(yAxisInt));
        if(isPercent) {
            DecimalFormat percentFormat = new DecimalFormat("##.##%");
            numAxis.setNumberFormatOverride(percentFormat);
        }
    }

    private static void setBorders(boolean isSet) throws Exception{
        if(isSet) {
            titleFormat.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
            normalFormat.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
            goodFormat.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
            expiredFormat.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
            noneFormat.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
        }
        else{
            titleFormat.setBorder(jxl.format.Border.NONE, jxl.format.BorderLineStyle.THIN);
            normalFormat.setBorder(jxl.format.Border.NONE, jxl.format.BorderLineStyle.THIN);
            goodFormat.setBorder(jxl.format.Border.NONE, jxl.format.BorderLineStyle.THIN);
            expiredFormat.setBorder(jxl.format.Border.NONE, jxl.format.BorderLineStyle.THIN);
            noneFormat.setBorder(jxl.format.Border.NONE, jxl.format.BorderLineStyle.THIN);
        }
    }

    private static void initializeFormat(){
        try{
            WritableFont titleFont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD,false);
            titleFormat = new WritableCellFormat(titleFont);
            titleFormat.setAlignment(Alignment.CENTRE);
            titleFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
            WritableFont myFont = new WritableFont(WritableFont.ARIAL,10, WritableFont.NO_BOLD, false);
            goodFormat = new WritableCellFormat(myFont);
            expiredFormat = new WritableCellFormat(myFont);
            noneFormat = new WritableCellFormat(myFont);
            normalFormat = new WritableCellFormat(myFont);
            normalFormat.setAlignment(Alignment.CENTRE);
            normalFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
            goodFormat.setBackground(Colour.LIGHT_GREEN);
            goodFormat.setAlignment(Alignment.CENTRE);
            goodFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
            expiredFormat.setBackground(Colour.RED);
            expiredFormat.setAlignment(Alignment.CENTRE);
            expiredFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
            noneFormat.setBackground(Colour.YELLOW);
            noneFormat.setAlignment(Alignment.CENTRE);
            noneFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
        }
        catch(WriteException e){
            System.out.println("Err: 5, Initialize Error.");
        }
    }

    private static JFrame createFrame(int x, int y, int width, int height, java.awt.Color colourUse, String title, LayoutManager layoutUse){
        JFrame resultFrame = new JFrame(title);
        resultFrame.setBounds(x, y, width,height);
        resultFrame.setBackground(colourUse);
        resultFrame.setResizable(false);
        resultFrame.setLayout(layoutUse);
        resultFrame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        return resultFrame;
    }

    private static JButton createButton(String textUse, int x, int y, int width, int height, Font fontUse, java.awt.Color colourUse){
        JButton resultButton = new JButton(textUse);
        resultButton.setBounds(x, y, width, height);
        resultButton.setFont(fontUse);
        if(colourUse == null) return resultButton;
        resultButton.setBackground(colourUse);
        return resultButton;
    }

    private static void createGUI(){
        JFrame mainFrame = createFrame(400, 100, 500, 350, Color.LIGHT_GRAY,
                "Performance Evaluator Version 2.0", null);
        JButton generateTableB = createButton("Generate Table", 50, 100, 175, 75,
                new Font("Arial", Font.BOLD, 18), null);
        JButton generatePlotB = createButton("Generate Plot", 275, 100, 175, 75,
                new Font("Arial", Font.BOLD, 18), null);
        JButton exitB = createButton("Exit!", 50, 200, 400, 75,
                new Font("Arial", Font.BOLD, 30), Color.RED);
        JLabel dateLabel = new JLabel("Date (YYYYMMDD): ");
        JTextField dateText = new JTextField(10);
        JPanel datePanel = new JPanel();
        dateLabel.setFont(new Font("Arial", Font.BOLD, 18));
        datePanel.add(dateLabel);
        datePanel.add(dateText);
        datePanel.setLayout(new GridLayout(1, 2));
        datePanel.setBounds(50, 50, 400, 25);
        mainFrame.add(datePanel);
        mainFrame.add(generateTableB);
        mainFrame.add(generatePlotB);
        mainFrame.add(exitB);
        mainFrame.addWindowListener(new MyWin());
        mainFrame.setVisible(true);
        exitB.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                System.out.println("Exit Program");
                System.exit(0);
            }
        });
        generateTableB.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                todayDate = dateText.getText();
                if(todayDate.isEmpty()){
                    JOptionPane.showMessageDialog(null,"Please enter date, format: YYYYMMDD",
                            "Warning",JOptionPane.WARNING_MESSAGE);
                }
                else if(todayDate.charAt(0) == '%'){
                    Runnable tableGeneration = new MyRunnable();
                    Thread thread1 = new Thread(tableGeneration);
                    thread1.start();
                }
                else if(todayDate.length() != 8 ||
                        (Integer.parseInt(todayDate.substring(0, 3)) > 2010 && Integer.parseInt(todayDate.substring(0, 3)) < 2015) ||
                        (Integer.parseInt(todayDate.substring(4, 5)) > 12 && Integer.parseInt(todayDate.substring(4, 5)) < 1) ||
                        Integer.parseInt(todayDate.substring(6, 7)) > 31 && Integer.parseInt(todayDate.substring(6, 7)) < 1){
                    JOptionPane.showMessageDialog(null,"Date format wrong, format: YYYYMMDD",
                            "Warning",JOptionPane.WARNING_MESSAGE);
                }
                else{
                    try {
                        threeDaysBef = dateAddition(todayDate, -3);
                        generateTable();
                        JOptionPane.showMessageDialog(null,"Table generated successfully","Progress",
                                JOptionPane.WARNING_MESSAGE);
                    }
                    catch (Exception e1){
                        JOptionPane.showMessageDialog(null,"Date format wrong, format: YYYYMMDD or file not found",
                                "Warning",JOptionPane.WARNING_MESSAGE);
                    }
                }
            }
        });
        generatePlotB.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String datePeriod = JOptionPane.showInputDialog("Please enter the period for the plot\n(Start Date - End Date: YYYYMMDD-YYYYMMDD):");
                while(true) {
                    startPlotDate = datePeriod.substring(0, 8);
                    endPlotDate = datePeriod.substring(9, 17);
                    System.out.println("StartDate: " + startPlotDate);
                    System.out.println("endDate: " + endPlotDate);
                    if(Long.parseLong(startPlotDate) < Long.parseLong(endPlotDate) && Long.parseLong(endPlotDate) <= Long.parseLong(actualTodayDate)){
                        break;
                    }
                    else{
                        datePeriod = JOptionPane.showInputDialog("Date period format wrong, please enter again\n(YYYYMMDD-YYYYMMDD)");
                    }
                }
                Runnable generatePlotThread = new MyRunnable2();
                Thread thread1 = new Thread(generatePlotThread);
                thread1.start();
            }
        });
    }

    public static void startMultipleGeneration(String startDate, String endDate){
        todayDate = startDate;
        while(!todayDate.equals(dateAddition(endDate, 1))){
            try {
                threeDaysBef = dateAddition(todayDate, -3);
                generateTable();
                buyerTracked = new LinkedList<>();
                buyerPerformance = new Performance[2];
                totalPerformance = new Performance("Total");
                todayDate = dateAddition(todayDate, 1);
            }
            catch (Exception e){
                todayDate = dateAddition(todayDate, 1);
            }
        }
    }

    private static void getAllDate(){
        String currDate = startPlotDate;
        while(true){
            dateQueue.add(currDate);
            currDate = dateAddition(currDate, 7);
            if(currDate.equals(endPlotDate)|| Long.parseLong(currDate) > Long.parseLong(endPlotDate)){
                dateQueue.add(currDate);
                break;
            }
        }

    }

    private static void retainData(){
//        Date startDate = myFormat.parse(startPlotDate);
//        Date endDate = myFormat.parse(endPlotDate);
//        long startDateL = Long.parseLong(startPlotDate);
//        long endDateL = Long.parseLong(endPlotDate);
        String currDate = dateQueue.poll();
        String filePathInit = "../PerformanceOutput/Performance";
        plotTrackingBuyers = new LinkedList<>();
        plotTrackingDate = new LinkedList<>();
        while(Long.parseLong(currDate) <= Long.parseLong(endPlotDate)){
            String filePath = filePathInit + currDate + ".xls";
            File performanceData = new File(filePath);
            try {
                InputStream is = new FileInputStream(performanceData.getAbsolutePath());
                Workbook wb = Workbook.getWorkbook(is);
                plotTrackingDate.add(currDate);
                retainDataHelper(wb, currDate);
                if(dateQueue.isEmpty()){
                    return;
                }
                currDate = dateQueue.poll();
                if(currDate.equals("")){
                    return;
                }
            }
            catch (FileNotFoundException e){
                System.out.println("File not found");
                System.out.println(currDate);
                currDate = dateAddition(currDate, 1);
                if(currDate.equals(dateQueue.peek())){
                    currDate = dateQueue.poll();
                }
            }
            catch (Exception e){
                System.out.println("Other error: " + e.toString());
            }
        }
    }

    private static void retainDataHelper(Workbook data, String currDate){
        Sheet dataSheet = data.getSheet(0);
        int rowNum = dataSheet.getRows();
        for(int i = 1; i < rowNum-1; i++){
            String currBuyer = dataSheet.getCell(0, i).getContents();
            int currGoodNum;
            int currExpireNum;
            int currMissNum;
            if(dataSheet.getCell(1, i).getContents().equals("")){
                currGoodNum = 0;
            }
            else{
                currGoodNum = Integer.parseInt(dataSheet.getCell(1, i).getContents());
            }
            if(dataSheet.getCell(2, i).getContents().equals("")){
                currExpireNum = 0;
            }
            else{
                currExpireNum = Integer.parseInt(dataSheet.getCell(2, i).getContents());
            }
            if(dataSheet.getCell(3, i).getContents().equals("")){
                currMissNum = 0;
            }
            else{
                currMissNum = Integer.parseInt(dataSheet.getCell(3, i).getContents());
            }
            if(plotTrackingBuyers.contains(currBuyer)){
                for(int j = 0; j < plotTrackingBuyers.size(); j++){
                    if(dataPerformance[j].getName().equals(currBuyer)){
                        dataPerformance[j].add(currBuyer, currGoodNum, currExpireNum, currMissNum, currDate);
                        break;
                    }
                }
            }
            else{
                copyPlotData();
                plotTrackingBuyers.add(currBuyer);
                PerformancePlotData currData = new PerformancePlotData(currBuyer);
                currData.add(currBuyer, currGoodNum, currExpireNum, currMissNum, currDate);
                dataPerformance[plotTrackingBuyers.size()-1] = currData;
            }
        }
    }

    private static Queue<Sheet> getSheetNum(Workbook wb){
        int sheet_size = wb.getNumberOfSheets();
        Queue<Sheet> results = new LinkedList<>();
        for (int index = 0; index < sheet_size; index++){
            Sheet dataSheet;
            if(wb.getSheet(index).getName().contains("OpenPO")){
                dataSheet = wb.getSheet(index);
                results.add(dataSheet);
            }
        }
        return results;
    }

    private static void getBuyerPerformance(File poData) throws Exception{
        InputStream is = new FileInputStream(poData.getAbsolutePath());
        int orderDateClnIdx;
        int promiseDateClnIdx;
        int buyerClnIdx;
        int vendorClnIdx;
        int remarkClnIdx;
        int currencyClnIdx;
        Workbook wb = Workbook.getWorkbook(is);
        WritableWorkbook wwb = Workbook.createWorkbook(poData, wb);
        Queue<Sheet> dataSheets = getSheetNum(wb);
        int removeLineCnt = 0;
        while(!dataSheets.isEmpty()){
            Sheet currDataSheet = dataSheets.poll();
            WritableSheet currDataSheetW = wwb.getSheet(currDataSheet.getName());
            promiseDateClnIdx = currDataSheet.findCell("Promised Date").getColumn();
            orderDateClnIdx = currDataSheet.findCell("Po Line Creation Date").getColumn();
            buyerClnIdx = currDataSheet.findCell("Buyer").getColumn();
            remarkClnIdx = buyerClnIdx + 1;
            vendorClnIdx = currDataSheet.findCell("Vendor").getColumn();
            currencyClnIdx = currDataSheet.findCell("Currency Code").getColumn();
            currDataSheetW.insertColumn(remarkClnIdx);
            currDataSheetW.setColumnView(remarkClnIdx, 19);
            jxl.write.Label remarkTitle = new jxl.write.Label(remarkClnIdx, 0, "Remark");
            currDataSheetW.addCell(remarkTitle);
            int rowNum = currDataSheet.getRows();
            for(int i = 1; i < rowNum; i++){
                String currVendor = currDataSheet.getCell(vendorClnIdx, i).getContents().toUpperCase();
                if(currVendor.contains("BUC")){
                    removeIdx.add(i);
                    removeLineCnt++;
                    continue;
                }
                String currBuyer = currDataSheet.getCell(buyerClnIdx, i).getContents();
                DateCell currOrderDateCell = (DateCell) currDataSheet.getCell(orderDateClnIdx, i);
                Date currOrderDate_temp = currOrderDateCell.getDate();
                String currOrderDate = myFormat.format(currOrderDate_temp);
                String currPromiseDate = "19900101";
                if(!currDataSheet.getCell(promiseDateClnIdx, i).getContents().equalsIgnoreCase("")){
                    DateCell currPromiseDateCell = (DateCell) currDataSheet.getCell(promiseDateClnIdx, i);
                    Date currPromiseDate_temp = new Date(currPromiseDateCell.getDate().getTime()-8*60*60*1000L);
                    currPromiseDate = myFormat.format(currPromiseDate_temp);
                }
                String currCurrency = currDataSheet.getCell(currencyClnIdx, i).getContents().toUpperCase();
                Performance currBuyerPerformance = null;
                if(buyerTracked.contains(currBuyer)){
                    for(int j = 0; j < buyerTracked.size(); j++){
                        if(buyerPerformance[j].isThisBuyer(currBuyer)){
                            currBuyerPerformance = buyerPerformance[j];
                        }
                    }
                }
                else{
                    copyData();
                    buyerTracked.add(currBuyer);
                    currBuyerPerformance = new Performance(currBuyer);
                    buyerPerformance[buyerTracked.size()-1] = currBuyerPerformance;
                }
                if(currBuyer.contains("Mark") && currPromiseDate.equals("19900101") && !currCurrency.equals("RMB")){
                    if(currBuyerPerformance != null) {
                        currBuyerPerformance.goodPromiseDateAdd();
                    }
                    totalPerformance.goodPromiseDateAdd();
                    jxl.write.Label remarkkCell = new jxl.write.Label(remarkClnIdx, i, "Promise Date OK", goodFormat);
                    currDataSheetW.addCell(remarkkCell);
                    continue;
                }
                if(currVendor.contains("BRANSON") || currVendor.contains("EMERSON") || currPromiseDate.equals("20150909")
                        || currVendor.contains("法埃龙") || currVendor.contains("惠恩") ||
                        (currVendor.contains("必能信") && currVendor.contains("东莞"))){
                    if(currBuyerPerformance != null) {
                        currBuyerPerformance.goodPromiseDateAdd();
                    }
                    totalPerformance.goodPromiseDateAdd();
                    jxl.write.Label remarkCell = new jxl.write.Label(remarkClnIdx, i, "Promise Date OK", goodFormat);
                    currDataSheetW.addCell(remarkCell);
                }
                else if(myFormat.parse(currPromiseDate).getTime() >= myFormat.parse(todayDate).getTime()){
                    if(currBuyerPerformance != null) {
                        currBuyerPerformance.goodPromiseDateAdd();
                    }
                    totalPerformance.goodPromiseDateAdd();
                    jxl.write.Label remarkCell = new jxl.write.Label(remarkClnIdx, i, "Promise Date OK", goodFormat);
                    currDataSheetW.addCell(remarkCell);
                }
                else if(currPromiseDate.equals("19900101")){
                    if(myFormat.parse(currOrderDate).getTime() >= myFormat.parse(threeDaysBef).getTime()){
                        if(currBuyerPerformance != null) {
                            currBuyerPerformance.goodPromiseDateAdd();
                        }
                        totalPerformance.goodPromiseDateAdd();
                        jxl.write.Label remarkCell = new jxl.write.Label(remarkClnIdx, i, "Promise Date OK", goodFormat);
                        currDataSheetW.addCell(remarkCell);
                    }
                    else{
                        if(currBuyerPerformance != null) {
                            currBuyerPerformance.nonePromiseDateAdd();
                        }
                        totalPerformance.nonePromiseDateAdd();
                        jxl.write.Label remarkCell = new jxl.write.Label(remarkClnIdx, i, "Promise Date Missed", noneFormat);
                        currDataSheetW.addCell(remarkCell);
                    }
                }
                else{
                    if(currBuyerPerformance != null) {
                        currBuyerPerformance.expiredPromiseDateAdd();
                    }
                    totalPerformance.expiredPromiseDateAdd();
                    jxl.write.Label remarkCell = new jxl.write.Label(remarkClnIdx, i, "Promise Date Expired", expiredFormat);
                    currDataSheetW.addCell(remarkCell);
                }
            }
            System.out.println("Remove line count: " + removeLineCnt);
            for(int i = 0; i < removeIdx.size(); i++){
                System.out.println(removeIdx.get(i));
                currDataSheetW.removeRow(removeIdx.get(i)-i);
            }
        }
        wwb.write();
        wwb.close();
    }

    private static void outputData(){
        String outputFilePath = "../PerformanceOutput/Performance" + todayDate + ".xls";
        for(int i = 0; i < buyerTracked.size(); i++){
            System.out.println(buyerPerformance[i].toString());
        }
        try{
            WritableWorkbook outputFile = Workbook.createWorkbook(new File(outputFilePath));
            WritableSheet sheet = outputFile.createSheet("Performance" + todayDate, 0);
            jxl.write.Label buyerLabel = new jxl.write.Label(0, 0, "Buyer", titleFormat);
            sheet.addCell(buyerLabel);
            sheet.setColumnView(0, 15);
            jxl.write.Label goodLabel = new jxl.write.Label(1, 0, "Promise Date OK", titleFormat);
            sheet.addCell(goodLabel);
            sheet.setColumnView(1, 16);
            jxl.write.Label expiredLabel = new jxl.write.Label(2, 0, "Promise Date Expired", titleFormat);
            sheet.addCell(expiredLabel);
            sheet.setColumnView(2, 21);
            jxl.write.Label missedLabel = new jxl.write.Label(3, 0, "Promise Date Missed", titleFormat);
            sheet.addCell(missedLabel);
            sheet.setColumnView(3, 20);
            jxl.write.Label totalLabel = new jxl.write.Label(4, 0, "Total", titleFormat);
            sheet.addCell(totalLabel);
            sheet.setColumnView(4, 10);
            jxl.write.Label percentLabel = new jxl.write.Label(5, 0, "Performance Percent", titleFormat);
            sheet.addCell(percentLabel);
            sheet.setColumnView(5, 20);
            int i;
            for(i = 0; i < buyerTracked.size(); i++){
                Performance currPerformance = buyerPerformance[i];
                String currName = currPerformance.getName();
                int goodNum = currPerformance.getGoodPromiseDate();
                int expiredNum = currPerformance.getExpiredPromiseDate();
                int missedNum = currPerformance.getNonePromiseDate();
                int totalNum = goodNum+expiredNum+missedNum;
                int goodPercent = (int) ((double)goodNum/(double)totalNum*100.0);
                WritableCellFormat currFormat;
                jxl.write.Label currBuyer = new jxl.write.Label(0, i+1, currName, normalFormat);
                sheet.addCell(currBuyer);
                currFormat = goodFormat;
                if(goodNum != 0) {
                    //Label goodPromise = new Label(1, i + 1, "" + goodNum, currFormat);
                    Number goodPromise = new Number(1, i+1, goodNum, currFormat);
                    sheet.addCell(goodPromise);
                }
                currFormat = expiredFormat;
                if(expiredNum != 0) {
                    Number expiredPromise = new Number(2, i + 1, expiredNum, currFormat);
                    sheet.addCell(expiredPromise);
                }
                currFormat = noneFormat;
                if(missedNum != 0) {
                    Number missedPromise = new Number(3, i + 1, missedNum, currFormat);
                    sheet.addCell(missedPromise);
                }
                Number totalPromise = new Number(4, i+1, totalNum, normalFormat);
                sheet.addCell(totalPromise);
                if(goodPercent > 80){
                    currFormat = goodFormat;
                }
                else if(goodPercent > 60){
                    currFormat = noneFormat;
                }
                else{
                    currFormat = expiredFormat;
                }
                jxl.write.Label goodPromisePercent = new jxl.write.Label(5, i+1, ""+goodPercent+"%", currFormat);
                sheet.addCell(goodPromisePercent);
            }
            jxl.write.Label buyerTotal = new jxl.write.Label(0, buyerTracked.size()+1, totalPerformance.getName(), titleFormat);
            sheet.addCell(buyerTotal);
            Number goodTotal = new Number(1, buyerTracked.size()+1, totalPerformance.getGoodPromiseDate(), titleFormat);
            sheet.addCell(goodTotal);
            Number expiredTotal = new Number(2, buyerTracked.size()+1, totalPerformance.getExpiredPromiseDate(), titleFormat);
            sheet.addCell(expiredTotal);
            Number missedTotal = new Number(3, buyerTracked.size()+1, totalPerformance.getNonePromiseDate(), titleFormat);
            sheet.addCell(missedTotal);
            int totalNum = totalPerformance.getGoodPromiseDate()+totalPerformance.getExpiredPromiseDate()+totalPerformance.getNonePromiseDate();
            Number totalTotal = new Number(4, buyerTracked.size()+1, totalNum, titleFormat);
            sheet.addCell(totalTotal);
            int totalPercent = (int) ((double)totalPerformance.getGoodPromiseDate()/(double)totalNum*100);
            jxl.write.Label percentTotal = new jxl.write.Label(5, buyerTracked.size()+1, ""+totalPercent+"%", titleFormat);
            sheet.addCell(percentTotal);
            outputFile.write();
            outputFile.close();
            count = 0;
        } catch (Exception e){
            System.out.println(e.toString());
            count++;
            if(count < 5){
                initializeFormat();
                outputData();
            }
        }
    }

    private static void copyData(){
        if(buyerTracked.size() < buyerPerformance.length){
            return;
        }
        Performance[] temp = new Performance[buyerPerformance.length * 2];
        for(int i = 0; i < buyerPerformance.length; i++){
            temp[i] = new Performance(buyerPerformance[i]);
        }
        buyerPerformance = temp;
    }

    private static void copyPlotData(){
        if(plotTrackingBuyers.size() < dataPerformance.length){
            return;
        }
        PerformancePlotData[] temp = new PerformancePlotData[dataPerformance.length * 2];
        for(int i = 0; i < dataPerformance.length; i++){
            temp[i] = new PerformancePlotData(dataPerformance[i]);
        }
        dataPerformance = temp;
    }
}
