package br.com.obra;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.List;

public class PlanilhaObraApp {
    
    public static void main(String[] args) {
        SwingUtilities.invokeLater(PlanilhaObraApp::createAndShowGUI);
    }
    
    private static void createAndShowGUI() {
        JFrame frame = new JFrame("Gerador de planilhas de controle de obra");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(500, 200);
        
        JPanel panel = new JPanel();
        panel.setLayout(new GridLayout(3, 1));
        
        JLabel titleLabel = new JLabel("Gerador de planilhas de controle de obra", SwingConstants.CENTER);
        titleLabel.setFont(new Font("Arial", Font.BOLD, 16));
        
        JButton perfisButton = new JButton("PERFIS");
        JButton componentesButton = new JButton("COMPONENTES");
        
        perfisButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                selecionarArquivos("RELAÇÃO DE BARRAS.PDF", true);
            }
        });
        
        componentesButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                selecionarArquivos("RELAÇÃO DE COMPONENTES.PDF", false);
            }
        });
        
        panel.add(titleLabel);
        panel.add(perfisButton);
        panel.add(componentesButton);
        
        frame.getContentPane().add(panel);
        frame.setLocationRelativeTo(null);
        frame.setVisible(true);
    }
    
    private static void selecionarArquivos(String filtroNomeArquivo, boolean isPerfis) {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setMultiSelectionEnabled(true);
        
        int option = fileChooser.showOpenDialog(null);
        
        if (option == JFileChooser.APPROVE_OPTION) {
            File[] selectedFiles = fileChooser.getSelectedFiles();
            
            for (File file : selectedFiles) {
                if (file.getName().toUpperCase().contains(filtroNomeArquivo.toUpperCase().replace(".PDF", ""))) {
                    try {
                        // Gerar nome do arquivo Excel baseado no nome do PDF
                        String nomeExcel = file.getName().replace(".pdf", ".xlsx").replace(".PDF", ".xlsx");
                        String caminhoExcel = file.getParent() + File.separator + nomeExcel;
                        
                        if (isPerfis) {
                            // Para PERFIS: passar null para componentes
                            List<String[]> dadosPerfis = PDFParserPerfis.extrairDados(file);
                            ExcelGenerator.gerarExcel(null, dadosPerfis, caminhoExcel);
                            JOptionPane.showMessageDialog(null, 
                                "Planilha de PERFIS gerada com sucesso!\n" + nomeExcel,
                                "Sucesso", JOptionPane.INFORMATION_MESSAGE);
                        } else {
                            // Para COMPONENTES: passar null para perfis
                            List<String[]> dadosComponentes = PDFParserComponentes.extrairDados(file);
                            ExcelGenerator.gerarExcel(dadosComponentes, null, caminhoExcel);
                            JOptionPane.showMessageDialog(null, 
                                "Planilha de COMPONENTES gerada com sucesso!\n" + nomeExcel,
                                "Sucesso", JOptionPane.INFORMATION_MESSAGE);
                        }
                    } catch (Exception ex) {
                        JOptionPane.showMessageDialog(null,
                                "Erro ao processar arquivo: " + file.getName() + "\n" + ex.getMessage(),
                                "Erro", JOptionPane.ERROR_MESSAGE);
                        ex.printStackTrace();
                    }
                } else {
                    JOptionPane.showMessageDialog(null,
                            "Arquivo inválido: " + file.getName() + "\nEsperado arquivo contendo: " + filtroNomeArquivo,
                            "Erro", JOptionPane.ERROR_MESSAGE);
                }
            }
        }
    }
}