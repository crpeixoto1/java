package br.com.obra;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.Loader;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class PDFParserPerfis {

    // Regex ajustado - ainda captura tudo mas vamos ignorar sobra e porcento
    private static final Pattern linhaPattern = Pattern.compile(
        "^(?<perfil>[A-Z0-9-]+)\\s+(?<qtde>\\d+)\\s+(?<barra>\\d+)\\s+(?<peso>\\d+,\\d{1,3})\\s+(?<sobra>\\d+,\\d{1,3})\\s+(?<porcento>\\d+,\\d{1,2})(?<tratCor>.*)$"
    );

    public static List<String[]> extrairDados(File pdfFile) {
        List<String[]> linhas = new ArrayList<>();
        String obraCod = "";
        String nomeObra = "";

        try (PDDocument document = Loader.loadPDF(pdfFile)) {
            PDFTextStripper stripper = new PDFTextStripper();
            String texto = stripper.getText(document);
            String[] linhasTexto = texto.split("\\r?\\n");

            // Extração dos dados da obra
            for (int i = 0; i < linhasTexto.length - 1; i++) {
                String linha = linhasTexto[i].trim();

                if (linha.startsWith("Obra cód.:")) {
                    for (int j = i + 1; j < i + 5 && j < linhasTexto.length; j++) {
                        String possivel = linhasTexto[j].trim();
                        if (!possivel.toLowerCase().startsWith("cliente") && !possivel.isEmpty()) {
                            String[] partes = possivel.split("\\s+", 2);
                            if (partes.length >= 2) {
                                obraCod = partes[0].trim();
                                nomeObra = partes[1].trim();
                            }
                            break;
                        }
                    }
                }
            }

            int numeroLinha = 1;

            for (String linha : linhasTexto) {
                linha = linha.trim().replaceAll("\\.(?=\\d{3}(\\D|$))", ""); // remove pontos de milhar

                Matcher matcher = linhaPattern.matcher(linha);
                if (matcher.matches()) {
                    String perfil = matcher.group("perfil").trim();
                    String qtde = matcher.group("qtde");
                    String barra = matcher.group("barra");
                    String peso = matcher.group("peso");
                    String tratCor = matcher.group("tratCor").trim();
                    
                    // Limpa o número de linha no final do tratamento/cor
                    tratCor = tratCor.replaceAll("\\s*\\d+\\s*$", "").trim();

                    // Array com apenas 8 elementos (removendo sobra e porcento)
                    linhas.add(new String[]{
                        String.valueOf(numeroLinha),
                        obraCod,
                        nomeObra,
                        perfil,
                        tratCor,
                        qtde,
                        barra,
                        peso
                    });

                    numeroLinha++;
                }
            }

            System.out.println("Total de linhas extraídas (PERFIS): " + linhas.size());

        } catch (IOException e) {
            e.printStackTrace();
        }

        return linhas;
    }
}