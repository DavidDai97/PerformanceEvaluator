import java.util.LinkedList;
import java.util.Queue;

public class PerformancePlotData {
    private Queue<Performance> performanceQueue;
    private String name;
    private Performance totalPerformance;

    public PerformancePlotData(String buyerName){
        this.performanceQueue = new LinkedList<>();
        this.name = buyerName;
        this.totalPerformance = new Performance(buyerName+"Total");
    }
    public PerformancePlotData(PerformancePlotData other){
        this.performanceQueue = new LinkedList<>(other.performanceQueue);
        this.name = other.name;
        this.totalPerformance = new Performance(other.totalPerformance);
    }

    public String getName(){
        return this.name;
    }
    public Queue<Performance> getPerformanceQueue() {
        return performanceQueue;
    }
    public Performance getTotalPerformance() {
        return totalPerformance;
    }

    public Performance poll(){
        return this.performanceQueue.poll();
    }
    public void add(String buyerMame, int goodNum, int expireNum, int missNum, String dataDate){
        Performance dataPerformanceNode = new Performance(buyerMame, goodNum, expireNum, missNum, dataDate);
        this.performanceQueue.add(dataPerformanceNode);
        this.totalPerformance.goodPromiseDateAdd(goodNum);
        this.totalPerformance.expiredPromiseDateAdd(expireNum);
        this.totalPerformance.nonePromiseDateAdd(missNum);
    }
    public boolean isEmpty(){
        return this.performanceQueue.isEmpty();
    }

}
