import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import jxl.DateCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.format.VerticalAlignment;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.Number;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

import static java.lang.System.exit;

public class Eavluator {
    private final static int orderDateClnIdx = 0;
    private final static int promiseDateClnIdx = 23;
    private final static int buyerClnIdx = 21;
    private final static int vendorClnIdx = 11;
    private final static int remarkClnIdx = 22;
    private static SimpleDateFormat myFormat = new SimpleDateFormat("yyyyMMdd");

    private static Queue<String> buyerTracked = new LinkedList();
    private static Performance[] buyerPerformance = new Performance[2];
    private static Performance totalPerformance = new Performance("Total");
    private static String todayDate;
    private static String threeDaysBef;

    private static WritableFont titleFont;
    private static WritableCellFormat titleFormat;
    private static WritableFont myFont;
    private static WritableCellFormat goodFormat;
    private static WritableCellFormat expiredFormat;
    private static WritableCellFormat noneFormat;
    private static WritableCellFormat normalFormat;

    private static int count = 0;
    public static void main(String[] args){
        todayDate = getTodayDate();       // Uncomment this line when actually use the tool
        String filePath = "../SourceFile/OpenPO" + todayDate + ".xls";
        initializeFormat();
        File openPoData = new File(filePath);
        getBuyerPerformance(openPoData);
        outputData();
        System.out.println("Done");
    }

    public static void initializeFormat(){
        try{
            titleFont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD,false);
            titleFormat = new WritableCellFormat(titleFont);
            titleFormat.setAlignment(Alignment.CENTRE);
            titleFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
            myFont = new WritableFont(WritableFont.ARIAL,10, WritableFont.NO_BOLD, false);
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

    public static void retainData(){
        String filePath = "../DataTracking/trackingData";
        File trackingData = new File(filePath);
        if(trackingData.exists()){
            retainDataHelper();
        }
        else{
            try{
                if(trackingData.createNewFile()){
                    System.out.println("New data tracking file has been created.");
                }
                else{
                    System.out.println("Unknown error, file not created.");
                }
            }
            catch (IOException e){
                System.out.println(e.toString());
            }
        }
    }
    private static void retainDataHelper(){

    }

    private static String getTodayDate(){
        Calendar myCalendar = Calendar.getInstance();
        Calendar myCalendar2 = Calendar.getInstance();
        myCalendar2.add(Calendar.DATE, -3);
        threeDaysBef = myFormat.format(myCalendar2.getTime());
        return myFormat.format(myCalendar.getTime());
    }

    private static Queue<Sheet> getSheetNum(Workbook wb){
        int sheet_size = wb.getNumberOfSheets();
        Queue<Sheet> results = new LinkedList();
        for (int index = 0; index < sheet_size; index++){
            Sheet dataSheet;
            if(wb.getSheet(index).getName().contains("OpenPO")){
                dataSheet = wb.getSheet(index);
                results.add(dataSheet);
            }
        }
        return results;
    }

    private static void getBuyerPerformance(File poData){
        try{
            InputStream is = new FileInputStream(poData.getAbsolutePath());
            Workbook wb = Workbook.getWorkbook(is);
            WritableWorkbook wwb = Workbook.createWorkbook(poData, wb);
            Queue<Sheet> dataSheets = getSheetNum(wb);
            while(!dataSheets.isEmpty()){
                Sheet currDataSheet = dataSheets.poll();
                WritableSheet currDataSheetW = wwb.getSheet(currDataSheet.getName());
                currDataSheetW.insertColumn(remarkClnIdx);
                Label remarkTitle = new Label(remarkClnIdx, 0, "Remark");
                currDataSheetW.addCell(remarkTitle);
                int rowNums = currDataSheet.getRows();
                for(int i = 1; i < rowNums; i++){
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
                    String currVendor = currDataSheet.getCell(vendorClnIdx, i).getContents().toUpperCase();
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
                    if(currVendor.contains("BRANSON") || currVendor.contains("EMERSON") || currPromiseDate.equals("20150909")
                            || currVendor.contains("法埃龙") || currVendor.contains("惠恩")){
                        currBuyerPerformance.goodPromiseDateAdd();
                        totalPerformance.goodPromiseDateAdd();
                        Label remarkCell = new Label(remarkClnIdx, i, "Promise Date OK", goodFormat);
                        currDataSheetW.addCell(remarkCell);
                    }
                    else if(myFormat.parse(currPromiseDate).getTime() >= myFormat.parse(todayDate).getTime()){
                        currBuyerPerformance.goodPromiseDateAdd();
                        totalPerformance.goodPromiseDateAdd();
                        Label remarkCell = new Label(remarkClnIdx, i, "Promise Date OK", goodFormat);
                        currDataSheetW.addCell(remarkCell);
                    }
                    else if(currPromiseDate.equals("19900101")){
                        if(myFormat.parse(currOrderDate).getTime() >= myFormat.parse(threeDaysBef).getTime()){
                            currBuyerPerformance.goodPromiseDateAdd();
                            totalPerformance.goodPromiseDateAdd();
                            Label remarkCell = new Label(remarkClnIdx, i, "Promise Date OK", goodFormat);
                            currDataSheetW.addCell(remarkCell);
                        }
                        else{
                            currBuyerPerformance.nonePromiseDateAdd();
                            totalPerformance.nonePromiseDateAdd();
                            Label remarkCell = new Label(remarkClnIdx, i, "Promise Date Missed", noneFormat);
                            currDataSheetW.addCell(remarkCell);
                        }
                    }
                    else{
                        currBuyerPerformance.expiredPromiseDateAdd();
                        totalPerformance.expiredPromiseDateAdd();
                        Label remarkCell = new Label(remarkClnIdx, i, "Promise Date Expired", expiredFormat);
                        currDataSheetW.addCell(remarkCell);
                    }
                }
            }
            wwb.write();
            wwb.close();
        }
        catch (FileNotFoundException e){
            System.out.println("Err: 0, Sorry! File is not found, or is opened right now.");
            exit(1);
        }
        catch(BiffException e){
            System.out.println("ERR: 1, " + e.toString());
        }
        catch(IOException e){
            System.out.println("Err: 2, " + e.toString());
        }
        catch(ParseException e){
            System.out.println("Err: 3, " + e.toString());
        }
        catch(WriteException e){
            System.out.println("Err: 4, Unable to write to file.");
        }
    }

    private static void outputData(){
        String outputFilePath = "../PerformanceOutput/Performance" + todayDate + ".xls";
        for(int i = 0; i < buyerTracked.size(); i++){
            System.out.println(buyerPerformance[i].toString());
        }
        try{
            WritableWorkbook outputFile = Workbook.createWorkbook(new File(outputFilePath));
            WritableSheet sheet = outputFile.createSheet("Performance" + todayDate, 0);
            Label buyerLabel = new Label(0, 0, "Buyer", titleFormat);
            sheet.addCell(buyerLabel);
            Label goodLabel = new Label(1, 0, "Promise Date OK", titleFormat);
            sheet.addCell(goodLabel);
            Label expiredLabel = new Label(2, 0, "Promise Date Expired", titleFormat);
            sheet.addCell(expiredLabel);
            Label missedLabel = new Label(3, 0, "Promise Date Missed", titleFormat);
            sheet.addCell(missedLabel);
            Label totalLabel = new Label(4, 0, "Total", titleFormat);
            sheet.addCell(totalLabel);
            Label percentLabel = new Label(5, 0, "Performance Percent", titleFormat);
            sheet.addCell(percentLabel);
            int i = 0;
            for(i = 0; i < buyerTracked.size(); i++){
                Performance currPerformance = buyerPerformance[i];
                String currName = currPerformance.getName();
                int goodNum = currPerformance.getGoodPromiseDate();
                int expiredNum = currPerformance.getExpiredPromiseDate();
                int missedNum = currPerformance.getNonePromiseDate();
                int totalNum = goodNum+expiredNum+missedNum;
                int goodPercent = (int) ((double)goodNum/(double)totalNum*100.0);
                WritableCellFormat currFormat;
                Label currBuyer = new Label(0, i+1, currName, normalFormat);
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
                Label goodPromisePercent = new Label(5, i+1, ""+goodPercent+"%", currFormat);
                sheet.addCell(goodPromisePercent);
            }
            Label buyerTotal = new Label(0, buyerTracked.size()+1, totalPerformance.getName(), titleFormat);
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
            Label percentTotal = new Label(5, buyerTracked.size()+1, ""+totalPercent+"%", titleFormat);
            sheet.addCell(percentTotal);
            outputFile.write();
            outputFile.close();
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
}
