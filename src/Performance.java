public class Performance {
    private String name;
    private int goodPromiseDate;
    private int expiredPromiseDate;
    private int nonePromiseDate;

    public Performance(String buyerName){
        this.name = buyerName;
        this.goodPromiseDate = 0;
        this.expiredPromiseDate = 0;
        this.nonePromiseDate = 0;
    }
    public Performance(String buyerName, int goodDate, int expiredDate, int noneDate){
        this.name = buyerName;
        this.goodPromiseDate = goodDate;
        this.expiredPromiseDate = expiredDate;
        this.nonePromiseDate = noneDate;
    }
    public Performance(Performance other){
        this.name = other.name;
        this.goodPromiseDate = other.goodPromiseDate;
        this.expiredPromiseDate = other.expiredPromiseDate;
        this.nonePromiseDate = other.nonePromiseDate;
    }

    public boolean isThisBuyer(String buyerName){
        return this.name.equalsIgnoreCase(buyerName);
    }

    public String getName(){
        return this.name;
    }
    public int getGoodPromiseDate(){
        return this.goodPromiseDate;
    }
    public int getExpiredPromiseDate(){
        return this.expiredPromiseDate;
    }
    public int getNonePromiseDate(){
        return this.nonePromiseDate;
    }

    public String toString(){
        return this.name + "," + this.goodPromiseDate + "," + this.expiredPromiseDate + "," + this.nonePromiseDate + ".";
    }

    public void goodPromiseDateAdd(){
        this.goodPromiseDate++;
    }
    public void expiredPromiseDateAdd(){
        this.expiredPromiseDate++;
    }
    public void nonePromiseDateAdd(){
        this.nonePromiseDate++;
    }

}
