// Arquivo: ExcelGenerator.java
package br.com.obra;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ExcelGenerator {

    /**
     * Gera arquivo .xls ou .xlsx com abas COMPONENTES e PERFIS.
     * Fecha o FileOutputStream antes de fechar o Workbook.
     */
    public static void gerarExcel(List<String[]> dadosComponentes,
                                  List<String[]> dadosPerfis,
                                  String nomeArquivo) throws IOException {
        Workbook workbook;
        String lower = nomeArquivo.toLowerCase();
        if (lower.endsWith(".xls")) {
            workbook = new HSSFWorkbook();
        } else {
            if (!lower.endsWith(".xlsx")) nomeArquivo += ".xlsx";
            workbook = new XSSFWorkbook();
        }

        if (dadosComponentes != null && !dadosComponentes.isEmpty()) {
            Sheet sheet = workbook.createSheet("COMPONENTES");
            criarAbaComponentes(sheet, dadosComponentes, workbook);
        }
        if (dadosPerfis != null && !dadosPerfis.isEmpty()) {
            Sheet sheet = workbook.createSheet("PERFIS");
            criarAbaPerfis(sheet, dadosPerfis, workbook);
        }

        // Apenas o FileOutputStream em try-with-resources
        try (FileOutputStream out = new FileOutputStream(nomeArquivo)) {
            workbook.write(out);
        } finally {
            workbook.close();
        }
    }

    private static void criarAbaComponentes(Sheet sheet, List<String[]> dados, Workbook wb) {
        CellStyle hdr = criarEstiloCabecalho(wb);
        CellStyle dat = criarEstiloDados(wb);
        String[] cols = {"Código","Descrição","Cor","Qtd.","UN"};
        Row r = sheet.createRow(0);
        for (int i = 0; i < cols.length; i++) {
            Cell c = r.createCell(i);
            c.setCellValue(cols[i]);
            c.setCellStyle(hdr);
        }
        int rn = 1;
        for (String[] ln : dados) {
            Row row = sheet.createRow(rn++);
            for (int i = 0; i < cols.length; i++) {
                Cell cell = row.createCell(i);
                String v = i < ln.length ? ln[i] : "";
                if (i == 3) {
                    try {
                        cell.setCellValue(Double.parseDouble(v.replace(",", ".")));
                    } catch (Exception ex) {
                        cell.setCellValue(v);
                    }
                } else {
                    cell.setCellValue(v);
                }
                cell.setCellStyle(dat);
            }
        }
        sheet.setColumnWidth(0, 4500);
        sheet.setColumnWidth(1,15000);
        sheet.setColumnWidth(2, 3500);
        sheet.setColumnWidth(3, 2000);
        sheet.setColumnWidth(4, 1500);
        sheet.createFreezePane(0,1);
    }

    private static void criarAbaPerfis(Sheet sheet, List<String[]> dados, Workbook wb) {
        CellStyle hdr = criarEstiloCabecalho(wb);
        CellStyle dat = criarEstiloDados(wb);
        CellStyle num = criarEstiloNumero(wb);
        String[] cols = {"Número","Obra Cód.","Nome Obra","Perfil","Trat./Cor","Qtde","Barra","Peso(kg)"};
        Row r = sheet.createRow(0);
        for (int i = 0; i < cols.length; i++) {
            Cell c = r.createCell(i);
            c.setCellValue(cols[i]);
            c.setCellStyle(hdr);
        }
        int rn = 1;
        for (String[] ln : dados) {
            Row row = sheet.createRow(rn++);
            for (int i = 0; i < cols.length; i++) {
                Cell c = row.createCell(i);
                String v = i < ln.length ? ln[i] : "";
                if (i == 0 || i == 5 || i == 6) {
                    try {
                        c.setCellValue(Double.parseDouble(v.replace(",", ".")));
                        c.setCellStyle(num);
                    } catch (Exception ex) {
                        c.setCellValue(v);
                        c.setCellStyle(dat);
                    }
                } else if (i == 7) {
                    try {
                        c.setCellValue(Double.parseDouble(v.replace(",", ".")));
                        c.setCellStyle(num);
                    } catch (Exception ex) {
                        c.setCellValue(v);
                        c.setCellStyle(dat);
                    }
                } else {
                    c.setCellValue(v);
                    c.setCellStyle(dat);
                }
            }
        }
        sheet.setColumnWidth(0,2500);
        sheet.setColumnWidth(1,4500);
        sheet.setColumnWidth(2,10000);
        sheet.setColumnWidth(3,3500);
        sheet.setColumnWidth(4,8000);
        sheet.setColumnWidth(5,2000);
        sheet.setColumnWidth(6,2000);
        sheet.setColumnWidth(7,2500);
        sheet.createFreezePane(0,1);
    }

    private static CellStyle criarEstiloCabecalho(Workbook w) {
        CellStyle s = w.createCellStyle();
        Font f = w.createFont();
        f.setBold(true);
        f.setFontHeightInPoints((short)11);
        s.setFont(f);
        s.setAlignment(HorizontalAlignment.CENTER);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        s.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setBorderBottom(BorderStyle.THIN);
        s.setBorderTop(BorderStyle.THIN);
        s.setBorderRight(BorderStyle.THIN);
        s.setBorderLeft(BorderStyle.THIN);
        return s;
    }

    private static CellStyle criarEstiloDados(Workbook w) {
        CellStyle s = w.createCellStyle();
        s.setBorderBottom(BorderStyle.THIN);
        s.setBorderTop(BorderStyle.THIN);
        s.setBorderRight(BorderStyle.THIN);
        s.setBorderLeft(BorderStyle.THIN);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    private static CellStyle criarEstiloNumero(Workbook w) {
        CellStyle s = criarEstiloDados(w);
        DataFormat fmt = w.createDataFormat();
        s.setDataFormat(fmt.getFormat("#,##0.00"));
        s.setAlignment(HorizontalAlignment.RIGHT);
        return s;
    }
}
