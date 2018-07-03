import java.awt.FlowLayout;
import java.awt.Frame;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;

public class MyWin extends WindowAdapter {

    @Override
    public void windowClosing(WindowEvent e) {
        // TODO Auto-generated method stub
        //System.out.println("Window closing"+e.toString());
        System.out.println("Exit Program.");
        System.exit(0);
    }

    @Override
    public void windowActivated(WindowEvent e) {
        //每次获得焦点 就会触发
        System.out.println("Focused");
        //super.windowActivated(e);
    }

    @Override
    public void windowOpened(WindowEvent e) {
        // TODO Auto-generated method stub
        System.out.println("Opened");
        //super.windowOpened(e);
    }
}