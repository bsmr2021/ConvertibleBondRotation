import jxl.read.biff.BiffException;
import jxl.write.WriteException;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;

public class ConvertibleBondRotation {
    private JPanel JPanel1;
    private JPanel JPanel2;
    private JPanel JPanel3;
    private JPanel JPanel4;
    private JButton LowPremiumRank;
    private JButton LowPremiumNotBuy;
    private JButton DoubleLowRank;
    private JButton DoubleLowNotBuy;
    private JButton VipRank;
    private JButton VipNotbuy;
    private JScrollPane JScrollPane1;
    private JTextArea textArea1;
    // 我目前持有的低溢价可转债列表
    String[][] strMyLowPremium = new String[100][2];
    // 最新低溢价可转债列表
    String[][] strLastestLowPremium = new String[300][2];
    // 上次的VIP可转债列表
    String[][] strVipOld = new String[100][2];
    // 最新的VIP可转债列表
    String[][] strVipNew = new String[100][2];
    // 我目前持有的双低可转债列表
    String[][] strMyDoubleLow = new String[100][2];
    // 最新双低可转债列表
    String[][] strLastestDoubleLow = new String[300][2];
    StringBuffer stringBuffer = new StringBuffer();

    public ConvertibleBondRotation() throws BiffException, IOException, WriteException {
        textArea1.setText("1.下载GitHub里ConvertibleBondRotation\\out\\artifacts\\ConvertibleBondRotation_jar\n" +
                "2.获取低溢价可转债排名、双低可转债排名并粘贴到《可转债轮动.xls》。\n" +
                "3.从券商下载最新的持仓并粘贴到《可转债轮动.xls》。\n" +
                "4.执行相应按钮即可得到需要轮动的结果。");

        com.company.ExcelTools excelTools = new com.company.ExcelTools();
        //先删除我的低溢价可转债持仓的其他品种，只保留可转债
        String[][] strMyTemp = new String[100][2];
        excelTools.readExcel(strMyTemp, "我的低溢价可转债持仓", 1, 0, 2);
        excelTools.DeleteNotConvertibleBond(strMyTemp);
        //然后得到了纯粹的可转债持仓列表
        excelTools.readExcel(strMyLowPremium, "我的低溢价可转债持仓",  1, 0, 2);
        //excelTools.PrintData(strMyLowPremium, "最终strMy", 0 , strMyLowPremium.length, 0, 2);

        excelTools.readExcel(strLastestLowPremium, "最新低溢价可转债排名",  1, 0, 2);
        //excelTools.PrintData(strLastestLowPremium, "strLastestLowPremium", 0 , strLastestLowPremium.length, 0, 2);

        excelTools.readExcel(strVipOld, "VIP轮动old",  1, 0, 2);
        //excelTools.PrintData(strVipOld, "strVipOld", 0 , strVipOld.length, 0, 2);
        excelTools.readExcel(strVipNew, "VIP轮动new",  1, 0, 2);
        //excelTools.PrintData(strVipNew, "strVipNew", 0 , strVipNew.length, 0, 2);

        excelTools.readExcel(strMyDoubleLow, "我的双低可转债持仓",  1, 0, 2);
        //excelTools.PrintData(strMyDoubleLow, "strMyDoubleLow", 0 , strMyDoubleLow.length, 0, 2);
        excelTools.readExcel(strLastestDoubleLow, "最新双低可转债排名",  2, 0, 2);
        //excelTools.PrintData(strLastestDoubleLow, "strLastestDoubleLow", 0 , strLastestDoubleLow.length, 0, 2);

        LowPremiumRank.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                stringBuffer.setLength(0);
                System.out.println("我的低溢价可转债持仓在最新的低溢价可转债列表里的排名:");
                for (int i = 0; i < strMyLowPremium.length; i++) {
                    int rank = -1;  //排名
                    int row = i+1;  //注意：raw是行，实际排名需要-1
                    if (strMyLowPremium[i][1] != null) {
                        for (int j = 0; j < strLastestLowPremium.length; j++) {
                            if (strLastestLowPremium[j][1] != null) {
                                if (strMyLowPremium[i][0].contains(strLastestLowPremium[j][0]))
                                {
                                    rank = j-1;//实际在Excel的行是j+1
                                }
                            }
                        }
                        PrintString(strMyLowPremium[i][1]+ "["+i+"]" + "在最新低溢价可转债的排名是:" +rank);
                    }
                }
                textArea1.setText(stringBuffer.toString());
            }
        });
        LowPremiumNotBuy.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                stringBuffer.setLength(0);
                System.out.println("最新低溢价可转债排名前50里，我的低溢价可转债持仓未买入的");
                for (int i = 0; i < 50; i++) {
                    int isExist = 0;
                    int row = i-1;  //实际在Excel的行是i+1
                    if (strLastestLowPremium[i][1] != null) {
                        for (int j = 0; j < strMyLowPremium.length; j++) {
                            if (strMyLowPremium[j][1] != null) {
                                if (strMyLowPremium[j][0].contains(strLastestLowPremium[i][0]))
                                {
                                    isExist = 1;
                                }
                            }
                        }
                        if (isExist != 0) {
                            //PrintString("最新低溢价可转债排名前50已买:" + strLastestLowPremium[i][1]+ "["+row+"]");
                        } else {
                            PrintString(strLastestLowPremium[i][0] + strLastestLowPremium[i][1]+ "--排名:"+row);
                        }
                    }
                }
                textArea1.setText(stringBuffer.toString());
            }
        });
        DoubleLowRank.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {

            }
        });
    }

    public static void main(String[] args) throws BiffException, IOException, WriteException {
        JFrame frame = new JFrame("ConvertibleBondRotation");
        frame.setContentPane(new ConvertibleBondRotation().JPanel1);
        frame.setTitle("可转债轮动");
        Dimension d = Toolkit.getDefaultToolkit().getScreenSize();
        int width = 500;
        int height = 800;
        frame.setBounds((d.width - width) / 2, (d.height - height) / 2, width, height);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setVisible(true);
    }

    public void PrintString(String printData) {
        System.out.println(printData);
        stringBuffer.append(printData + "\n");
    }

}
