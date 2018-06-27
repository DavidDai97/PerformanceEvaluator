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

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class Eavluator {
    private final static int orderDateClnIdx = 0;
    private final static int promiseDateClnIdx = 23;
    private final static int buyerClnIdx = 21;
    private final static int vendorClnIdx = 11;
    private static SimpleDateFormat myFormat = new SimpleDateFormat("yyyyMMdd");

    private static Queue<String> buyerTracked = new LinkedList();
    private static Performance[] buyerPerformance = new Performance[2];
    private static Performance totalPerformance = new Performance("Total");
    private static String todayDate;
    private static String threeDaysBef;

    public static void main(String[] args){
        //todayDate = getTodayDate();       // Uncomment this line when actually use the tool
        todayDate = "20180621";             // Test usage, comment this line when actually use the tool
        threeDaysBef = "20180618";          // Test usage, comment this line when actually use the tool
        String filePath = "../SourceFile/OpenPO" + todayDate + ".xls";
        File openPoData = new File(filePath);
        getBuyerPerformance(openPoData);
        outputData();
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

    private static void getBuyerPerformance(File poData){
        try{
            InputStream is = new FileInputStream(poData.getAbsolutePath());
            Workbook wb = Workbook.getWorkbook(is);
            int sheet_size = wb.getNumberOfSheets();
            Queue<Sheet> dataSheets = new LinkedList();
            for (int index = 0; index < sheet_size; index++){
                Sheet dataSheet;
                if(wb.getSheet(index).getName().contains("OpenPO")){
                    dataSheet = wb.getSheet(index);
                    dataSheets.add(dataSheet);
                }
            }
            while(!dataSheets.isEmpty()){
                Sheet currDataSheet = dataSheets.poll();
                int rowNums = currDataSheet.getRows();
                for(int i = 1; i < rowNums; i++){
                    String currBuyer = currDataSheet.getCell(buyerClnIdx, i).getContents();
                    DateCell currOrderDateCell = (DateCell) currDataSheet.getCell(orderDateClnIdx, i);
                    Date currOrderDate_temp = currOrderDateCell.getDate();//new Date(currOrderDateCell.getDate().getTime()-8*60*60*1000L);
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
                    if(currVendor.contains("BRANSON") || currVendor.contains("EMERSON") || currPromiseDate.equals("20150909")){
                        currBuyerPerformance.goodPromiseDateAdd();
                        totalPerformance.goodPromiseDateAdd();
                    }
                    else if(myFormat.parse(currPromiseDate).getTime() > myFormat.parse(todayDate).getTime()){
                        currBuyerPerformance.goodPromiseDateAdd();
                        totalPerformance.goodPromiseDateAdd();
                    }
                    else if(currPromiseDate.equals("19900101")){
                        if(myFormat.parse(currOrderDate).getTime() >= myFormat.parse(threeDaysBef).getTime()){
                            currBuyerPerformance.goodPromiseDateAdd();
                            totalPerformance.goodPromiseDateAdd();
                        }
                        else{
                            currBuyerPerformance.nonePromiseDateAdd();
                            totalPerformance.nonePromiseDateAdd();
                        }
                    }
                    else{
                        //System.out.println(currPromiseDate);
                        currBuyerPerformance.expiredPromiseDateAdd();
                        totalPerformance.expiredPromiseDateAdd();
                    }
                }
            }
        }
        catch (FileNotFoundException e){
            System.out.println("Err: 0, Sorry! File not found.");
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
    }

    private static void outputData(){
        String outputFilePath = "../PerformanceOutput/Performance" + todayDate + ".xls";
        for(int i = 0; i < buyerTracked.size(); i++){
            System.out.println(buyerPerformance[i].toString());
        }
        try{
            WritableWorkbook outputFile = Workbook.createWorkbook(new File(outputFilePath));
            WritableSheet sheet = outputFile.createSheet("Performance" + todayDate, 0);
            WritableFont titleFont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD,false);
            WritableCellFormat titleFormat = new WritableCellFormat(titleFont);
            titleFormat.setAlignment(Alignment.CENTRE);
            titleFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
            WritableFont myFont = new WritableFont(WritableFont.ARIAL,10, WritableFont.NO_BOLD, false);
            WritableCellFormat goodFormat = new WritableCellFormat(myFont);
            WritableCellFormat expiredFormat = new WritableCellFormat(myFont);
            WritableCellFormat noneFormat = new WritableCellFormat(myFont);
            WritableCellFormat normalFormat = new WritableCellFormat(myFont);
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
            for(int i = 0; i < buyerTracked.size(); i++){
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
                if(goodNum == 0){
                    currFormat = normalFormat;
                }
                else{
                    currFormat = goodFormat;
                }
                Label goodPromise = new Label(1, i+1, ""+goodNum, currFormat);
                sheet.addCell(goodPromise);
                if(expiredNum == 0){
                    currFormat = normalFormat;
                }
                else{
                    currFormat = expiredFormat;
                }
                Label expiredPromise = new Label(2, i+1, ""+expiredNum, currFormat);
                sheet.addCell(expiredPromise);
                if(missedNum == 0){
                    currFormat = normalFormat;
                }
                else{
                    currFormat = noneFormat;
                }
                Label missedPromise = new Label(3, i+1, ""+missedNum, currFormat);
                sheet.addCell(missedPromise);
                Label totalPromise = new Label(4, i+1, ""+totalNum, normalFormat);
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
            Label goodTotal = new Label(1, buyerTracked.size()+1, ""+totalPerformance.getGoodPromiseDate(), titleFormat);
            sheet.addCell(goodTotal);
            Label expiredTotal = new Label(2, buyerTracked.size()+1, ""+totalPerformance.getExpiredPromiseDate(), titleFormat);
            sheet.addCell(expiredTotal);
            Label missedTotal = new Label(3, buyerTracked.size()+1, ""+totalPerformance.getNonePromiseDate(), titleFormat);
            sheet.addCell(missedTotal);
            int totalNum = totalPerformance.getGoodPromiseDate()+totalPerformance.getExpiredPromiseDate()+totalPerformance.getNonePromiseDate();
            Label totalTotal = new Label(4, buyerTracked.size()+1, ""+totalNum, titleFormat);
            sheet.addCell(totalTotal);
            int totalPercent = (int) ((double)totalPerformance.getGoodPromiseDate()/(double)totalNum*100);
            Label percentTotal = new Label(5, buyerTracked.size()+1, ""+totalPercent+"%", titleFormat);
            sheet.addCell(percentTotal);
            outputFile.write();
            outputFile.close();
        } catch (Exception e){
            System.out.println(e.toString());
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
