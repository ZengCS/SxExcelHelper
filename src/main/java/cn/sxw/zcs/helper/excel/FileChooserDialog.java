package cn.sxw.zcs.helper.excel;

import javax.swing.*;
import java.awt.*;
import java.io.File;

/**
 * @Author ZengCS
 * @Date 2023/3/28 14:52
 * @Copyright 四川生学教育科技有限公司
 * @Address 成都市天府软件园E3-3F
 */
public class FileChooserDialog extends JFrame {
    private static final Font FONT = new Font(null, Font.PLAIN, 16);
    private static final Font FONT_S = new Font(null, Font.PLAIN, 13);
    private static final Insets INSETS = new Insets(5, 5, 5, 5);
    private static final Dimension TV_DIMENSION = new Dimension(300, 36);
    private static final Dimension BTN_DIMENSION = new Dimension(100, 33);
    private static final Color THEME_COLOR = Color.decode("#13C1C1");

    private File lastDir = new File("/Users/zengcs/Documents/华西心理促进/");
    private File file1 = null;
    private File file2 = null;
    private JTextPane tooltips = makeTextPane("\n             \n", true, Color.RED);

    private void showTips(String tips, boolean error) {
        tooltips.setText("\n" + tips + "\n");
        tooltips.setForeground(error ? Color.RED : THEME_COLOR);
    }

    public static void main(String[] args) {
        new FileChooserDialog();
    }

    private JTextPane makeTextPane(String text, boolean large, Color color) {
        JTextPane textPane = new JTextPane();
        textPane.setText(text);
        textPane.setFont(large ? FONT : FONT_S);
        textPane.setForeground(color);
        textPane.setFocusable(false);

        return textPane;
    }

    private JTextField makeTextField(String text, Color color) {
        JTextField jTextField = new JTextField(text, color == THEME_COLOR ? 47 : 40);
        jTextField.setEditable(false);
        jTextField.setForeground(color);
        jTextField.setSelectedTextColor(Color.WHITE);
        jTextField.setFont(FONT);
        jTextField.setMargin(INSETS);
        return jTextField;
    }

    private JButton makeButton(String text) {
        JButton button = new JButton(text);
        button.setFont(FONT);
        button.setPreferredSize(BTN_DIMENSION);
        return button;
    }

    public FileChooserDialog() {
        super("华西心理促进文档分析助手");

        JPanel panel = new JPanel();
        panel.setBackground(Color.WHITE);

        // Welcome&Version
        panel.add(makeTextPane("\n欢迎使用 华西心理促进文档分析助手(Ver 1.0)\n", true, THEME_COLOR));

        // 文本框
        final JTextField tvInput1 = makeTextField(" 请选择信息采集表格", Color.GRAY);
        final JTextField tvInput2 = makeTextField(" 请选择平台导出表格", Color.GRAY);
        final JTextField tvAnalysis = makeTextField(" 请选择文件后点击'分析表格'按钮", THEME_COLOR);

        // 按钮组
        final JButton btnInput1 = makeButton("选择文件");
        final JButton btnInput2 = makeButton("选择文件");
        final JButton btnAnalysis = makeButton("分析表格");
        final JButton btnReset = makeButton("重置");
        btnAnalysis.setForeground(Color.RED);
        btnAnalysis.setEnabled(false);

        panel.add(tvInput1);
        panel.add(btnInput1);
        panel.add(tvInput2);
        panel.add(btnInput2);
        panel.add(tvAnalysis);

        JPanel childPanel = new JPanel();
        panel.add(childPanel);
        childPanel.setPreferredSize(new Dimension(700, 42));
        childPanel.add(btnAnalysis);
        childPanel.add(btnReset);
        childPanel.setBackground(Color.WHITE);

        // 创建文本面板
        panel.add(tooltips);

        // 监听按钮
        btnInput1.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setCurrentDirectory(lastDir);
            int returnValue = fileChooser.showOpenDialog(null);
            if (returnValue == JFileChooser.APPROVE_OPTION) {
                file1 = fileChooser.getSelectedFile();
                lastDir = file1.getParentFile();
                tvInput1.setText(file1.getAbsolutePath());
                btnAnalysis.setEnabled(file1 != null && file2 != null);
            }
        });
        btnInput2.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setCurrentDirectory(lastDir);
            int returnValue = fileChooser.showOpenDialog(null);
            if (returnValue == JFileChooser.APPROVE_OPTION) {
                file2 = fileChooser.getSelectedFile();
                tvInput2.setText(file2.getAbsolutePath());
                btnAnalysis.setEnabled(file1 != null && file2 != null);
            }
        });
        btnReset.addActionListener(e -> {
            file1 = null;
            file2 = null;
            tvInput1.setText(" 请选择信息采集表格");
            tvInput2.setText(" 请选择平台导出表格");
            tvAnalysis.setText(" 请选择文件后点击'分析表格'按钮");
            btnAnalysis.setEnabled(false);
            btnInput1.setEnabled(true);
            btnInput2.setEnabled(true);
            showTips("             ", false);
        });
        btnAnalysis.addActionListener(e -> {
            if (file1 == null) {
                showTips("请选择信息采集表格", true);
                return;
            }
            if (file2 == null) {
                showTips("请选择平台导出表格", true);
                return;
            }
            showTips("正在处理表格，请勿关闭程序...", true);

            tvAnalysis.setText("");
            btnInput1.setEnabled(false);
            btnInput2.setEnabled(false);
            btnAnalysis.setEnabled(false);
            btnReset.setEnabled(false);
            new Thread(() -> {
                try {
                    Thread.sleep(1200);
                } catch (Exception ex) {
                    ex.printStackTrace();
                    btnReset.setEnabled(true);
                }

                try {
                    ExcelReader excelReader = new ExcelReader(file1, file2);
                    File destFile = excelReader.analysis();

                    tvAnalysis.setText(destFile.getAbsolutePath());
                    StringBuilder sb = new StringBuilder();
                    sb.append("\n表格分析成功：");
                    sb.append("\n****************************************************");
                    sb.append("\n数据采集学生共：" + excelReader.stuList.size() + "人");
                    sb.append("\n其中已完成人数：" + excelReader.completeList.size() + "人");
                    sb.append("\n其中未完成人数：" + excelReader.unCompleteList.size() + "人");
                    sb.append("\n\n****************************************************");
                    sb.append("\n异常学生信息(不在数据采集范围内的学生)：");
                    sb.append("\n异常已完成人数：" + excelReader.errorCompleteList.size() + "人");
                    sb.append("\n异常未完成人数：" + excelReader.errorUnCompleteList.size() + "人");
                    showTips(sb.toString(), false);

                    btnInput1.setEnabled(true);
                    btnInput2.setEnabled(true);
                    btnAnalysis.setEnabled(true);
                    btnReset.setEnabled(true);
                } catch (Exception exception) {
                    showTips(exception.getMessage(), true);
                    btnReset.setEnabled(true);
                }
            }).start();
        });

        setContentPane(panel);
        setSize(740, 540);
        setResizable(false);
        setLocationRelativeTo(null);
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        setVisible(true);
    }
}
