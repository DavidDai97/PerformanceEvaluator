import javax.swing.*;

public class MyRunnable2 implements Runnable{
    @Override
    public void run() {
        try {
            Eavluator.generatePlot();
            JOptionPane.showMessageDialog(null, "Plot generated successfully", "Progress",
                    JOptionPane.WARNING_MESSAGE);
        }
        catch (Exception el){
            System.out.println("Error: " + el);
            JOptionPane.showMessageDialog(null,"Error: "+ el +
                            "\nPlease ensure all files existed in the required folder, and remain closed when the program is running. Please run the program again.","Progress",
                    JOptionPane.WARNING_MESSAGE);
        }
    }
}
