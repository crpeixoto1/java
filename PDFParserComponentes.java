package br.com.obra;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Parser de componentes: agrega linhas multilinha e usa regex
 * para extrair código, descrição, cor, quantidade e unidade.
 */
public class PDFParserComponentes {

    // Pattern único para capturar: código, descrição, cor, quantidade e unidade
    private static final Pattern COMPONENT_PATTERN = Pattern.compile(
        "^\\s*(?:\\d+\\s+)?([+A-Za-z0-9-]+)\\s+(.+?)\\s+"
        + "(NATURAL|PRETO|PRETA|INOX|CINZA|BRANCO|BRANCA|FOSCO|FOSCA|ECOSEAL|S/A)\\s+"
        + "(\\d+(?:[.,]\\d+)?)\\s+"
        + "(PC|PR|MT|TB|KG|CJ)\\b"
    );

    /**
     * Extrai lista de componentes: [código, descrição, cor, quantidade, unidade]
     */
    public static List<String[]> extrairDados(File pdfFile) {
        List<String[]> linhas = new ArrayList<>();
        try (PDDocument document = Loader.loadPDF(pdfFile)) {
            PDFTextStripper stripper = new PDFTextStripper();
            stripper.setSortByPosition(true);
            String[] rawLines = stripper.getText(document).split("\\r?\\n");

            for (int i = 0; i < rawLines.length; i++) {
                String line = rawLines[i].trim();
                // processa apenas linhas que começam com número ou código
                if (!line.matches("^\\s*(?:\\d+\\s+)?[+A-Za-z0-9-]+\\s+.*")) continue;

                // agrega possíveis linhas abaixo até próximo item
                StringBuilder sb = new StringBuilder(line);
                int j = i + 1;
                while (j < rawLines.length && !rawLines[j].trim().matches("^\\s*(?:\\d+\\s+)?[+A-Za-z0-9-]+\\s+.*")) {
                    String nxt = rawLines[j].trim();
                    if (!nxt.isEmpty()) sb.append(' ').append(nxt);
                    j++;
                }
                i = j - 1;
                String row = sb.toString().replaceAll("\\s+", " ").trim();

                Matcher m = COMPONENT_PATTERN.matcher(row);
                if (!m.find()) continue;
                String codigo     = m.group(1);
                String descricao  = m.group(2).trim();
                String cor        = m.group(3);
                String quantidade = m.group(4).replace(',', '.');
                String unidade    = m.group(5);

                linhas.add(new String[]{codigo, descricao, cor, quantidade, unidade});
            }
        } catch (IOException e) {
            System.err.println("ERRO ao ler PDF Componentes: " + e.getMessage());
            e.printStackTrace();
        }
        return linhas;
    }
}
