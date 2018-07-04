import javax.swing.*;

public class MyRunnable implements Runnable{
    @Override
    public void run() {
        Eavluator.startMultipleGeneration(Eavluator.todayDate.substring(1, 9), Eavluator.todayDate.substring(10));
        JOptionPane.showMessageDialog(null,"Table generated successfully","Progress",
                JOptionPane.WARNING_MESSAGE);
    }
}
